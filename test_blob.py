from azure_blob_helper import download_blob

download_blob(
    "data_integration/supplier_R12/config/Validation_Rules.xlsx",
    "temp/Validation_Rules.xlsx"
)

download_blob(
    "data_integration/supplier_R12/reference/PAY_TERMS.xlsx",
    "temp/PAY_TERMS.xlsx"
)

download_blob(
    "data_integration/supplier_R12/reference/PAY_METHOD.xlsx",
    "temp/PAY_METHOD.xlsx"
)

download_blob(
    "data_integration/supplier_R12/reference/SUPPLIER_TYPE.xlsx",
    "temp/SUPPLIER_TYPE.xlsx"
)

download_blob(
    "data_integration/supplier_R12/reference/Country_Address_Rules.xlsx",
    "temp/Country_Address_Rules.xlsx"
)