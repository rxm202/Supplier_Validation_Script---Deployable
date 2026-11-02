from azure_blob_helper import upload_blob

upload_blob(
    "input/current_input.xlsx",
    "data_integration/supplier_R12/input/test_file.xlsx"
)

print("Upload Success")