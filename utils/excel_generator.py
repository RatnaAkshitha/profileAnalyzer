import pandas as pd
from io import BytesIO

def create_excel_report_bytes(df: pd.DataFrame) -> bytes:
    """
    Accepts a DataFrame and returns an bytes object representing the Excel file.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="profile_report")
    return output.getvalue()
