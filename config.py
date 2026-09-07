"""Pipeline configuration — update paths for your environment."""

from pathlib import Path

BASE_PATH = Path(
    r"C:\Users\smigaikw\OneDrive - Tata Communications\1.work\MIS\BSSE_ComValid_RuchiNew"
)

ACCESS_FOLDER = BASE_PATH / "Access"
MAIN_FOLDER = BASE_PATH / "September 26"

INPUT_CSV = Path(r"C:\Users\smigaikw\Downloads\report1788756359321.csv")
DUMP_FILE = ACCESS_FOLDER / "Dump.xlsx"
WORK_FILE = ACCESS_FOLDER / "wrok.xlsx"

ACCESS_DB = ACCESS_FOLDER / "BSSE CommVett Comvalid 1.accdb"
MAIN_XLSB = Path(
    r"C:\Users\smigaikw\OneDrive - Tata Communications\1.work\MIS\BSSE_ComValid_RuchiNew\September 26\BSSE_CommVett_COMVALID_base.xlsb"
)

DATA_SHEET = "Sheet1"
PIVOT_SHEET = "Sheet2"
MACRO_NAME = "Macro1"

# Mail settings (Outlook Classic integration — wire these up later)
GROUP_MAIL_ID = "bss-mis@tatacommunications.com"
MAIL_SIGNATURE = (
    "Best Regards,\n"
    "BSS MIS Team\n"
    "Tata Communications"
)
