"""RAO radio calculator entrypoint."""

from .rao_radio import run_calc_capture

CALCULATOR_ID = "rao_radio"
PROVIDER = "РАО"
STATUS = "ready"

__all__ = ["run_calc_capture", "CALCULATOR_ID", "PROVIDER", "STATUS"]
