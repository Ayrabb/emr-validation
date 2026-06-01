# app/blob/connector.py
# ─────────────────────────────────────────────────────────────────────────────
# Downloads the combined RADET Excel from Azure Blob Storage.
# Matches the DHIS2 pipeline blob_client.py conventions exactly.
#
# Blob path convention (ACE-1 programme)
# ──────────────────────────────────────
#   Container : combine-radets-xls
#   Path      : {YYYY-MM-DD}/{prefix}-{YYYY-MM-DD}.xlsx
#   Example   : 2026-05-09/ACE-1_Combined_RADET-2026-05-09.xlsx
#
# Authentication (tried in order)
# ────────────────────────────────
#   1. Full connection string  — AZURE_STORAGE_CONNECTION_STRING
#   2. Account + SAS token    — AZURE_STORAGE_ACCOUNT + AZURE_SAS_TOKEN
#
# Returns raw bytes — ValidationService.run() accepts file_bytes: bytes.
# Retries on transient errors AND ResourceModifiedError (in-flight replacement).
# ─────────────────────────────────────────────────────────────────────────────

import logging
import time
from datetime import date

from app.core.config import settings

logger = logging.getLogger(__name__)


class BlobConnector:

    def __init__(self):
        self._connection_string = settings.azure_storage_connection_string
        self._account           = settings.azure_storage_account
        self._sas_token         = settings.azure_sas_token
        self._container         = settings.azure_container_name
        self._blob_prefix       = settings.azure_blob_prefix
        self._max_retries       = settings.blob_download_retries
        self._timeout           = settings.blob_download_timeout

    # ── Blob path ─────────────────────────────────────────────────────────────

    def blob_path(self, for_date: date) -> str:
        """
        Build the blob path for a given date.
        Pattern : {YYYY-MM-DD}/{prefix}-{YYYY-MM-DD}.xlsx
        Example : 2026-05-09/ACE-1_Combined_RADET-2026-05-09.xlsx
        """
        ds = for_date.strftime("%Y-%m-%d")
        return f"{ds}/{self._blob_prefix}-{ds}.xlsx"

    # ── Auth ──────────────────────────────────────────────────────────────────

    def _service_client(self):
        from azure.storage.blob import BlobServiceClient

        # connection_timeout — seconds to establish the TCP connection
        # read_timeout      — seconds to wait for the next chunk of data on the socket
        #                     this is the key setting that prevents mid-stream stalls
        transport_kwargs = {
            "connection_timeout": 30,
            "read_timeout": self._timeout,
        }

        if self._connection_string:
            return BlobServiceClient.from_connection_string(
                self._connection_string, **transport_kwargs
            )

        if self._account and self._sas_token:
            token = self._sas_token.lstrip("?")
            account_url = f"https://{self._account}.blob.core.windows.net?{token}"
            return BlobServiceClient(account_url=account_url, **transport_kwargs)

        raise RuntimeError(
            "Azure credentials missing.\n"
            "Set either:\n"
            "  AZURE_STORAGE_CONNECTION_STRING  (full connection string)\n"
            "or both:\n"
            "  AZURE_STORAGE_ACCOUNT  (e.g. baycentral)\n"
            "  AZURE_SAS_TOKEN        (SAS token from Storage Explorer)"
        )

    # ── Download ──────────────────────────────────────────────────────────────

    _CHUNK_SIZE = 10 * 1_048_576   # 10 MB per range-request

    def download(self, for_date: date | None = None) -> bytes:
        """
        Download the RADET blob for for_date (defaults to today).

        Uses chunked range-requests (10 MB each) so a transient stall only
        retries the affected chunk rather than restarting the whole file.
        Each chunk is retried up to settings.blob_download_retries times
        with exponential back-off before the whole run is marked as failed.
        """
        from azure.core.exceptions import (
            HttpResponseError,
            ResourceModifiedError,
            ResourceNotFoundError,
            ServiceRequestError,
        )
        from http.client import IncompleteRead
        try:
            from requests.exceptions import (
                ReadTimeout,
                ConnectionError as RequestsConnectionError,
            )
            _net_errors = (IncompleteRead, ReadTimeout, RequestsConnectionError)
        except ImportError:
            _net_errors = (IncompleteRead,)

        # All transient errors that warrant a retry
        TRANSIENT = (ServiceRequestError, HttpResponseError, ResourceModifiedError) + _net_errors

        if for_date is None:
            for_date = date.today()

        blob_name = self.blob_path(for_date)
        service     = self._service_client()
        blob_client = (
            service
            .get_container_client(self._container)
            .get_blob_client(blob_name)
        )

        # ── Resolve blob size so we can issue range-requests ──────────────────
        _props_backoff = (5, 15, 30)
        _props_error: Exception = RuntimeError("No attempt made")
        total_size: int | None = None
        for _attempt in range(1, self._max_retries + 1):
            try:
                total_size = blob_client.get_blob_properties(timeout=30).size
                break
            except ResourceNotFoundError:
                raise RuntimeError(
                    f"Blob not found: {blob_name!r}\n"
                    f"Container      : {self._container!r}\n"
                    "Check that the combined RADET file has been uploaded for this date."
                )
            except Exception as exc:
                _props_error = exc
                _delay = _props_backoff[min(_attempt - 1, len(_props_backoff) - 1)]
                logger.warning(f"get_blob_properties failed (attempt {_attempt}/{self._max_retries}): {exc} — retrying in {_delay}s…")
                time.sleep(_delay)
        else:
            raise RuntimeError(f"Could not read blob properties after {self._max_retries} attempts: {_props_error}")
        assert total_size is not None

        total_chunks = max(1, -(-total_size // self._CHUNK_SIZE))   # ceiling division
        size_mb      = total_size / 1_048_576
        logger.info(
            f"Downloading blob: container={self._container!r}  blob={blob_name!r}  "
            f"{size_mb:.1f} MB  ({total_chunks} chunks)"
        )

        # ── Fetch each chunk with independent retry ───────────────────────────
        backoff  = (5, 15, 30)   # seconds between retry attempts
        chunks: list[bytes] = []
        offset = 0

        for chunk_idx in range(1, total_chunks + 1):
            length     = min(self._CHUNK_SIZE, total_size - offset)
            last_error: Exception = RuntimeError("No attempt made")

            for attempt in range(1, self._max_retries + 1):
                try:
                    data = blob_client.download_blob(
                        offset=offset, length=length, timeout=self._timeout
                    ).readall()
                    chunks.append(data)
                    logger.info(
                        f"Chunk {chunk_idx}/{total_chunks} OK  "
                        f"({(offset + length) / 1_048_576:.1f}/{size_mb:.1f} MB)"
                    )
                    break   # chunk succeeded — move to the next one

                except ResourceNotFoundError:
                    raise RuntimeError(
                        f"Blob not found: {blob_name!r}\n"
                        f"Container      : {self._container!r}\n"
                        "Check that the combined RADET file has been uploaded for this date."
                    )

                except TRANSIENT as exc:
                    last_error = exc
                    delay      = backoff[min(attempt - 1, len(backoff) - 1)]
                    logger.warning(
                        f"Chunk {chunk_idx}/{total_chunks} transient error "
                        f"(attempt {attempt}/{self._max_retries}): {exc}  "
                        f"— retrying in {delay}s…"
                    )
                    time.sleep(delay)

                except Exception as exc:
                    raise RuntimeError(
                        f"Unexpected error on chunk {chunk_idx}: {exc}"
                    ) from exc

            else:
                # All retries for this chunk exhausted
                raise RuntimeError(
                    f"Chunk {chunk_idx}/{total_chunks} failed after "
                    f"{self._max_retries} attempts: {last_error}"
                )

            offset += length

        data = b"".join(chunks)
        logger.info(f"Blob downloaded: {blob_name!r}  {len(data) / 1_048_576:.1f} MB")
        return data


# Module-level singleton — reused across all scheduled runs
blob_connector = BlobConnector()