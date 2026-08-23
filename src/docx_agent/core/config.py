"""
Core Configuration Management for Docx-Agent Platform.
Provides centralized, typed, environment-driven settings with safe defaults.
"""

import os
from pathlib import Path
from typing import Optional
from pydantic import BaseModel, Field


class Settings(BaseModel):
    """
    Application runtime configuration schema.
    """
    env: str = Field(default="development", description="Environment mode: development | testing | production")
    log_level: str = Field(default="INFO", description="Log level: DEBUG | INFO | WARNING | ERROR")
    structured_logs: bool = Field(default=True, description="Output structured JSON logs if enabled")

    # Workspace Server Settings
    workspace_port: int = Field(default=8765, description="Port for visual workspace server")
    workspace_host: str = Field(default="127.0.0.1", description="Host address for visual workspace server")
    auto_open_browser: bool = Field(default=True, description="Automatically open browser on workspace launch")

    # Engine & File Constraints
    max_file_size_mb: int = Field(default=50, description="Maximum allowed DOCX file size in megabytes")
    pagination_dpi: int = Field(default=96, description="Standard screen DPI for layout engine calculations")
    auto_backup: bool = Field(default=True, description="Automatically create atomic backup before saving")

    # Storage paths
    base_dir: Path = Field(default_factory=lambda: Path(__file__).parent.parent.parent.parent)

    @classmethod
    def from_env(cls) -> "Settings":
        """
        Loads configuration from environment variables with fallback to defaults.
        """
        env_val = os.getenv("DOCX_AGENT_ENV", "development")
        log_lvl = os.getenv("DOCX_AGENT_LOG_LEVEL", "INFO")
        struct_logs = os.getenv("DOCX_AGENT_STRUCTURED_LOGS", "true").lower() in ("true", "1", "yes")

        port = int(os.getenv("DOCX_AGENT_WORKSPACE_PORT", "8765"))
        host = os.getenv("DOCX_AGENT_WORKSPACE_HOST", "127.0.0.1")
        open_b = os.getenv("DOCX_AGENT_AUTO_OPEN_BROWSER", "true").lower() in ("true", "1", "yes")

        max_mb = int(os.getenv("DOCX_AGENT_MAX_FILE_SIZE_MB", "50"))
        dpi = int(os.getenv("DOCX_AGENT_PAGINATION_DPI", "96"))
        backup = os.getenv("DOCX_AGENT_AUTO_BACKUP", "true").lower() in ("true", "1", "yes")

        return cls(
            env=env_val,
            log_level=log_lvl,
            structured_logs=struct_logs,
            workspace_port=port,
            workspace_host=host,
            auto_open_browser=open_b,
            max_file_size_mb=max_mb,
            pagination_dpi=dpi,
            auto_backup=backup,
        )


_config_instance: Optional[Settings] = None


def get_config() -> Settings:
    """
    Returns the singleton application settings instance.
    """
    global _config_instance
    if _config_instance is None:
        _config_instance = Settings.from_env()
    return _config_instance
