"""Pipeline runner with step-level status callbacks for the dashboard."""

import subprocess
import time
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Callable, Optional

from config import (
    ACCESS_DB,
    DATA_SHEET,
    DUMP_FILE,
    INPUT_CSV,
    MACRO_NAME,
    MAIN_XLSB,
    PIVOT_SHEET,
    WORK_FILE,
)


class StepStatus(Enum):
    PENDING = "pending"
    RUNNING = "running"
    SUCCESS = "success"
    FAILED = "failed"
    SKIPPED = "skipped"


@dataclass
class PipelineStep:
    id: str
    title: str
    description: str
    status: StepStatus = StepStatus.PENDING
    message: str = ""


PIPELINE_STEPS = [
    PipelineStep(
        "step1",
        "Step 1 — Data Pipeline",
        "CSV → Dump.xlsx → Work.xlsx",
    ),
    PipelineStep(
        "step2",
        "Step 2 — Pivot Refresh",
        "Refresh dynamic pivot in Work.xlsx",
    ),
    PipelineStep(
        "step3",
        "Step 3 — Access Macro",
        "Run Access macro and export to main workbook",
    ),
    PipelineStep(
        "step4",
        "Step 4 — Main Workbook Refresh",
        "Refresh main .xlsb from Work.xlsx",
    ),
    PipelineStep(
        "step5",
        "Step 5 — Final Pivot Update",
        "Update pivots in main workbook",
    ),
]


LogCallback = Callable[[str], None]
StepCallback = Callable[[str, StepStatus, str], None]
DoneCallback = Callable[[bool, Optional[str]], None]


def log_banner(message: str) -> str:
    return f"\n{'=' * 60}\n{message}\n{'=' * 60}"


def force_close_excel(log: LogCallback) -> None:
    log("Checking for running Excel processes...")
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "EXCEL.EXE", "/T"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        time.sleep(2)
        log("Excel processes cleared.")
    except Exception as exc:
        log(f"Failed to close Excel: {exc}")


def run_pipeline(
    on_log: LogCallback,
    on_step: StepCallback,
    on_done: DoneCallback,
    stop_event=None,
) -> None:
    """Execute the full pipeline, reporting progress through callbacks."""

    def check_stop() -> bool:
        return stop_event is not None and stop_event.is_set()

    current_step_id: Optional[str] = None

    def run_step(step_id: str, fn) -> None:
        nonlocal current_step_id
        if check_stop():
            on_step(step_id, StepStatus.SKIPPED, "Stopped by user")
            raise InterruptedError("Pipeline stopped by user")

        current_step_id = step_id
        on_step(step_id, StepStatus.RUNNING, "In progress...")
        fn()
        on_step(step_id, StepStatus.SUCCESS, "Completed")
        current_step_id = None

    try:
        from BSSPKG import excel_refresh4, macro3, pipeline1, PivotUp5, workR2
    except ImportError as exc:
        on_done(False, f"BSSPKG import failed: {exc}")
        return

    try:
        # Step 1
        on_log(log_banner("STEP 1 : Running Data Pipeline"))
        run_step(
            "step1",
            lambda: pipeline1.main(
                input_csv_path=str(INPUT_CSV),
                dump_excel_path=str(DUMP_FILE),
                work_excel_path=str(WORK_FILE),
            ),
        )
        on_log("Pipeline completed successfully.")

        if check_stop():
            raise InterruptedError("Pipeline stopped by user")

        # Step 2
        on_log(log_banner("STEP 2 : Refreshing Pivot Table"))
        run_step(
            "step2",
            lambda: workR2.refresh_dynamic_pivot(
                str(WORK_FILE),
                DATA_SHEET,
                PIVOT_SHEET,
            ),
        )
        on_log("Pivot refresh completed.")

        if check_stop():
            raise InterruptedError("Pipeline stopped by user")

        # Step 3
        on_log(log_banner("STEP 3 : Running Access Macro"))
        run_step(
            "step3",
            lambda: (
                force_close_excel(on_log),
                macro3.run_macro_and_export_ordered(
                    str(ACCESS_DB),
                    MACRO_NAME,
                    str(MAIN_XLSB),
                ),
            ),
        )
        on_log("Access macro completed.")

        if check_stop():
            raise InterruptedError("Pipeline stopped by user")

        # Step 4
        on_log(log_banner("STEP 4 : Refreshing Main Workbook"))
        run_step(
            "step4",
            lambda: excel_refresh4.main(
                str(MAIN_XLSB),
                str(WORK_FILE),
            ),
        )
        on_log("Main workbook refresh completed.")

        if check_stop():
            raise InterruptedError("Pipeline stopped by user")

        # Step 5
        on_log(log_banner("STEP 5 : Updating Final Pivots"))
        run_step(
            "step5",
            lambda: PivotUp5.main(str(MAIN_XLSB)),
        )
        on_log("Final pivot update completed.")

        on_log(log_banner("PROCESS COMPLETED SUCCESSFULLY"))
        on_done(True, None)

    except InterruptedError as exc:
        on_log(log_banner("PROCESS STOPPED"))
        on_log(str(exc))
        on_done(False, str(exc))

    except Exception as exc:
        on_log(log_banner("PROCESS FAILED"))
        on_log(f"Error: {exc}")
        if current_step_id:
            on_step(current_step_id, StepStatus.FAILED, str(exc))
        on_done(False, str(exc))
