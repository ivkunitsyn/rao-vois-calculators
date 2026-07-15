"""VOIS radio calculator entrypoint."""

from calculators.rao_radio.backend.calculator import run_calc_capture

CALCULATOR_ID = "vois_radio"
PROVIDER = "ВОИС"
STATUS = "ready"

__all__ = ["run_calc_capture", "CALCULATOR_ID", "PROVIDER", "STATUS"]
