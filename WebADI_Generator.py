from openpyxl import load_workbook
import pandas as pd

# ==========================================================
# READ VALIDATED OUTPUT
# ==========================================================

output_df = pd.read_excel(
    "output/Supplier_Output.xlsx",
    dtype=str
)

output_df = output_df.fillna("")

# ==========================================================
# OPEN WEBADI TEMPLATE
# ==========================================================

wb = load_workbook(
    "template/SIT_LOAD_template_webadi.xlsm"
)

ws = wb["WebADI"]

header_row = 12
start_row = 14

# ==========================================================
# BUILD WEBADI COLUMN MAPPING
# ==========================================================

webadi_columns = {}

for col in range(1, ws.max_column + 1):

    header = ws.cell(
        header_row,
        col
    ).value

    if header:

        webadi_columns[
            str(header).strip()
        ] = col

# ==========================================================
# LOAD DATA
# ==========================================================

for idx, row in output_df.iterrows():

    target_row = start_row + idx

    for field in output_df.columns:

        if field in webadi_columns:

            target_col = webadi_columns[field]

            value = row[field]

            # --------------------------------------------------
            # SITE NAME - PREFIX ZEROES TO MAKE 4 DIGITS
            # --------------------------------------------------

            if field == "S_SUPPLIER_SITE_NAME":

                value = str(value).strip()

                if value.isdigit():

                    value = value.zfill(4)

            # --------------------------------------------------
            # PAYMENT PRIORITY MUST BE NUMERIC
            # --------------------------------------------------

            elif field == "S_SITE_PAY_PAYMENT_PRIORITY":

                try:

                    if str(value).strip() != "":

                        value = int(
                            float(
                                str(value).strip()
                            )
                        )

                except:

                    pass

            ws.cell(
                target_row,
                target_col
            ).value = value

# ==========================================================
# SAVE NEW FILE
# ==========================================================

wb.save(
    "output/Supplier_WebADI_Output.xlsx"
)

print(
    "WebADI file created."
)