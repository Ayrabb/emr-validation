# app/core/config.py
# ─────────────────────────────────────────────────────────────────────────────
# Central settings.  All values read from .env with typed defaults.
# Desktop app .env (PORT=8765) still works — new keys have safe defaults.
#
# Azure auth: connection string preferred; falls back to account + SAS token.
# Blob path:  {container}/{YYYY-MM-DD}/{prefix}-{YYYY-MM-DD}.xlsx
# Example:    combine-radets-xls/2026-05-09/ACE-1_Combined_RADET-2026-05-09.xlsx
# ─────────────────────────────────────────────────────────────────────────────

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
    )


    # ── Server ────────────────────────────────────────────────────────────────
    port:           int = 8001          # 8765 for desktop, 8001 for central
    app_version:    str = "1.0.0"

    # ── File handling ─────────────────────────────────────────────────────────
    max_upload_mb:  int = 200
    report_dir:     str = r"C:\EMRValidation\reports"

    # ── SQLite database ───────────────────────────────────────────────────────
    database_path:  str = r"C:\EMRValidation\data\emr_validation.db"
    retention_days: int = 90

    # ── Azure Blob — auth (connection string preferred) ───────────────────────
    # Mode 1: full connection string
    azure_storage_connection_string: str = ""
    # Mode 2: account name + SAS token (alternative to connection string)
    azure_storage_account: str = ""
    azure_sas_token:        str = ""

    # ── Azure Blob — path config ──────────────────────────────────────────────
    # Container holds the combined RADET files
    azure_container_name:  str = "combine-radets-xls"
    # Filename prefix — path built as: {date}/{prefix}-{date}.xlsx
    # Example: 2026-05-09/ACE-1_Combined_RADET-2026-05-09.xlsx
    azure_blob_prefix:     str = "ACE-1_Combined_RADET"
    blob_download_retries: int = 3
    blob_download_timeout: int = 60   # seconds per attempt before giving up

    # ── APScheduler ───────────────────────────────────────────────────────────
    scheduler_timezone: str = "Africa/Lagos"

    # ── BayCentral CORS ───────────────────────────────────────────────────────
    baycentral_origin: str = (
        "http://localhost:3000,http://127.0.0.1,http://localhost"
    )

    # ── Helpers ───────────────────────────────────────────────────────────────

    @property
    def allowed_origins(self) -> list[str]:
        return [o.strip() for o in self.baycentral_origin.split(",") if o.strip()]

    @property
    def database_url(self) -> str:
        path = self.database_path.replace("\\", "/")
        return f"sqlite:///{path}"


settings = Settings()