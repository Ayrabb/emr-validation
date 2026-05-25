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

    def download(self, for_date: date | None = None) -> bytes:
        """
        Download the RADET blob for for_date (defaults to today).
        Returns raw bytes.  Retries up to settings.blob_download_retries times
        with back-off on transient errors and ResourceModifiedError.
        """
        from azure.core.exceptions import (
            HttpResponseError,
            ResourceModifiedError,
            ResourceNotFoundError,
            ServiceRequestError,
        )

        if for_date is None:
            for_date = date.today()

        blob_name = self.blob_path(for_date)
        logger.info(
            f"Downloading blob: container={self._container!r}  "
            f"blob={blob_name!r}"
        )

        service     = self._service_client()
        blob_client = (
            service
            .get_container_client(self._container)
            .get_blob_client(blob_name)
        )

        backoff = (1, 3, 6)   # seconds between retries — matches DHIS2 pipeline
        last_error: Exception = RuntimeError("No attempt made")

        for attempt in range(1, self._max_retries + 1):
            try:
                data    = blob_client.download_blob(timeout=self._timeout).readall()
                size_mb = len(data) / 1_048_576
                logger.info(
                    f"Blob downloaded: {blob_name!r}  "
                    f"{size_mb:.1f} MB  (attempt {attempt})"
                )
                return data

            except ResourceNotFoundError:
                # Not transient — file genuinely missing for this date
                raise RuntimeError(
                    f"Blob not found: {blob_name!r}\n"
                    f"Container      : {self._container!r}\n"
                    "Check that the combined RADET file has been uploaded for this date."
                )

            except ResourceModifiedError as exc:
                # File is being replaced upstream — wait and retry
                last_error = exc
                delay = backoff[min(attempt - 1, len(backoff) - 1)]
                logger.warning(
                    f"Blob modified mid-download (attempt {attempt}/{self._max_retries}). "
                    f"Retrying in {delay}s…"
                )
                time.sleep(delay)

            except (ServiceRequestError, HttpResponseError) as exc:
                last_error = exc
                delay = backoff[min(attempt - 1, len(backoff) - 1)]
                logger.warning(
                    f"Transient blob error (attempt {attempt}/{self._max_retries}): "
                    f"{exc}.  Retrying in {delay}s…"
                )
                time.sleep(delay)

            except Exception as exc:
                raise RuntimeError(f"Unexpected blob error: {exc}") from exc

        raise RuntimeError(
            f"Blob download failed after {self._max_retries} attempts: {last_error}"
        )


# Module-level singleton — reused across all scheduled runs
blob_connector = BlobConnector()