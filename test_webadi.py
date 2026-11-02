from openpyxl import load_workbook

wb = load_workbook(
    "template/SIT_LOAD_template_webadi.xlsm"
)

wb.save(
    "output/Test_WebADI.xlsx"
)

print("Done")