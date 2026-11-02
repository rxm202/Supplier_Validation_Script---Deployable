import pandas as pd

# ==========================================
# CONFIGURATION
# ==========================================

INPUT_FILE = "input/Suvinil_sample_test_23rdSept.csv"

# ==========================================
# READ FILE
# ==========================================

df = pd.read_csv(INPUT_FILE, low_memory=False)

# ==========================================
# DEFINE HEADER COLUMNS
# ==========================================

header_cols = [
    "H_SUPPLIER_NAME",
    "H_SUPPLIER_NUMBER",
    "H_SUPPLIER_NAME_ALT",
    "H_TAX_PAYER_ID",
    "H_TAX_REGISTRATION_NUMBER",
    "H_SUPPLIER_TYPE",
    "H_ATTRIBUTE15"
]

# ==========================================
# CREATE HEADER COMBINATION KEY
# ==========================================

df["HEADER_COMBINATION"] = (
    df["H_SUPPLIER_NAME"].fillna("").astype(str).str.upper().str.strip()
    + "|"
    + df["H_SUPPLIER_NUMBER"].fillna("").astype(str).str.upper().str.strip()
    + "|"
    + df["H_SUPPLIER_NAME_ALT"].fillna("").astype(str).str.upper().str.strip()
    + "|"
    + df["H_TAX_PAYER_ID"].fillna("").astype(str).str.upper().str.strip()
    + "|"
    + df["H_TAX_REGISTRATION_NUMBER"].fillna("").astype(str).str.upper().str.strip()
    + "|"
    + df["H_SUPPLIER_TYPE"].fillna("").astype(str).str.upper().str.strip()
    + "|"
    + df["H_ATTRIBUTE15"].fillna("").astype(str).str.upper().str.strip()
)

# ==========================================
# FIND DUPLICATE COMBINATIONS
# ==========================================

duplicate_mask = df.duplicated(
    subset=["HEADER_COMBINATION"],
    keep=False
)

# ==========================================
# ERROR FILE
# ==========================================

error_df = df[duplicate_mask].copy()

if not error_df.empty:
    error_df["FAILURE_REASON"] = (
        "Duplicate Header Combination Found"
    )

# ==========================================
# OUTPUT FILE
# ==========================================

valid_df = df[~duplicate_mask].copy()

# remove technical field
valid_df.drop(
    columns=["HEADER_COMBINATION"],
    inplace=True,
    errors="ignore"
)

error_df.drop(
    columns=["HEADER_COMBINATION"],
    inplace=True,
    errors="ignore"
)

# ==========================================
# SAVE FILES
# ==========================================

valid_df.to_csv(
    "output/Supplier_Output.csv",
    index=False
)

error_df.to_csv(
    "output/Supplier_Validation_Error.csv",
    index=False
)

# ==========================================
# SUMMARY
# ==========================================

print("=" * 50)
print("VALIDATION COMPLETED")
print("=" * 50)
print("Total Records :", len(df))
print("Valid Records :", len(valid_df))
print("Rejected Records :", len(error_df))
print()
print("Files Created:")
print("output/Supplier_Output.csv")
print("output/Supplier_Validation_Error.csv")