import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

# ==========================================================
# CONFIGURATION
# ==========================================================

INPUT_FILE = "input/current_input.xlsx"

# ==========================================================
# READ FILE
# ==========================================================

df = pd.read_excel(
    INPUT_FILE,
    dtype=str
)

df = df.fillna("")

df.columns = (
    df.columns
    .astype(str)
    .str.strip()
)

df = df.reset_index(drop=True)

print(
    f"Input File Loaded : {INPUT_FILE}"
)

print(
    f"Total Records : {len(df)}"
)

# ==========================================================
# FAILURE REASON COLUMN
# ==========================================================

df["FAILURE_REASON"] = ""
df["WARNING_MESSAGE"] = ""
df["ERROR_COLUMNS"] = ""
df["WARNING_COLUMNS"] = ""



# ==========================================
# LOAD MANDATORY RULES
# ==========================================

rules_df = pd.read_excel(
    "config/Validation_Rules.xlsx"
)

mandatory_rules = rules_df[
    (
        rules_df["CATEGORY"]
        .astype(str)
        .str.strip()
        .str.upper()
        == "MANDATORY"
    )
    &
    (
        rules_df["ACTIVE"]
        .astype(str)
        .str.strip()
        .str.upper()
        == "Y"
    )
]

# ==========================================================
# RULE 0
# DYNAMIC MANDATORY VALIDATION
# ==========================================================

for _, rule in mandatory_rules.iterrows():

    field = str(
        rule["FIELD_NAME"]
    ).strip()

    severity = str(
        rule["SEVERITY"]
    ).strip().upper()

    # Skip if column does not exist
    if field not in df.columns:

        print(
            f"WARNING : Column {field} not found."
        )

        continue

    field_mask = (
        df[field].isna()
        |
        (
            df[field]
            .fillna("")
            .astype(str)
            .str.strip()
            == ""
        )
    )

    if severity == "ERROR":

        df.loc[
            field_mask,
            "FAILURE_REASON"
        ] += (
            f"{field} Missing | "
        )

        df.loc[
            field_mask,
            "ERROR_COLUMNS"
        ] += (
            f"{field}|"
        )

    elif severity == "WARNING":

        df.loc[
            field_mask,
            "WARNING_MESSAGE"
        ] += (
            f"{field} Missing | "
        )

        df.loc[
            field_mask,
            "WARNING_COLUMNS"
        ] += (
            f"{field}|"
        )

# ==========================================================
# RULE 1
# DUPLICATE SUPPLIER NUMBER + SITE NAME
# ==========================================================

duplicate_mask = df.duplicated(
    subset=[
        "H_ATTRIBUTE15",
        "S_SUPPLIER_SITE_NAME"
    ],
    keep=False
)

df.loc[
    duplicate_mask,
    "FAILURE_REASON"
] += (
    "Duplicate record found. The same H_ATTRIBUTE15 and "
    "S_SUPPLIER_SITE_NAME combination already exists. "
    "Each S_SUPPLIER_SITE_NAME must be unique for the "
    "specified H_ATTRIBUTE15 | "
)

df.loc[
    duplicate_mask,
    "ERROR_COLUMNS"
] += (
    "H_ATTRIBUTE15|"
    "S_SUPPLIER_SITE_NAME|"
)

# ==========================================================
# RULE 2
# SITE NAME LENGTH
# ==========================================================

site_len_mask = (
    df["S_SUPPLIER_SITE_NAME"]
    .fillna("")
    .astype(str)
    .str.len()
    > 15
)

df.loc[
    site_len_mask,
    "FAILURE_REASON"
] += (
    "S_SUPPLIER_SITE_NAME exceeds the maximum "
    "length of 15 characters | "
)

df.loc[
    site_len_mask,
    "ERROR_COLUMNS"
] += (
    "S_SUPPLIER_SITE_NAME|"
)

# ==========================================================
# RULE 3
# PAYMENT PRIORITY NUMERIC
# ==========================================================

payment_mask = ~pd.to_numeric(
    df["S_SITE_PAY_PAYMENT_PRIORITY"],
    errors="coerce"
).notna()

payment_mask = (
    payment_mask
    &
    df["S_SITE_PAY_PAYMENT_PRIORITY"].notna()
)

df.loc[
    payment_mask,
    "FAILURE_REASON"
] += (
    "S_SITE_PAY_PAYMENT_PRIORITY must be data type numeric | "
)

df.loc[
    payment_mask,
    "ERROR_COLUMNS"
] += (
    "S_SITE_PAY_PAYMENT_PRIORITY|"
)

# ==========================================================
# RULE 4
# CONTACT FIRST NAME UPPERCASE
# ==========================================================

first_name_mask = (
    df["CONTACT_FIRST_NAME"]
    .fillna("")
    .astype(str)
    !=
    df["CONTACT_FIRST_NAME"]
    .fillna("")
    .astype(str)
    .str.upper()
)

df.loc[
    first_name_mask,
    "FAILURE_REASON"
] += (
    "CONTACT_FIRST_NAME must be entered in uppercase | "
)

df.loc[
    first_name_mask,
    "ERROR_COLUMNS"
] += (
    "CONTACT_FIRST_NAME|"
)

# ==========================================================
# RULE 5
# CONTACT LAST NAME UPPERCASE
# ==========================================================

last_name_mask = (
    df["CONTACT_LAST_NAME"]
    .fillna("")
    .astype(str)
    !=
    df["CONTACT_LAST_NAME"]
    .fillna("")
    .astype(str)
    .str.upper()
)

df.loc[
    last_name_mask,
    "FAILURE_REASON"
] += (
    "CONTACT_LAST_NAME must be entered in uppercase | "
)

df.loc[
    last_name_mask,
    "ERROR_COLUMNS"
] += (
    "CONTACT_LAST_NAME|"
)

# ==========================================================
# RULE 6
# SUPPLIER NAME / ADDRESS NAME CLEANUP
# ==========================================================

df["H_SUPPLIER_NAME"] = (
    df["H_SUPPLIER_NAME"]
    .fillna("")
    .astype(str)
    .str.strip()
    .str.upper()
)

df["A_ADDRESS_NAME"] = (
    df["A_ADDRESS_NAME"]
    .fillna("")
    .astype(str)
    .str.strip()
    .str.upper()
)

# ==========================================================
# RULE 7
# DEFAULT SUPPLIER NOTIFICATION METHOD
# ==========================================================

df["S_SUPP_SITE_NOTIF_METHOD"] = df.apply(
    lambda row:
    ""
    if pd.isna(row["S_SITE_EMAIL_ADDRESS"])
       or str(row["S_SITE_EMAIL_ADDRESS"]).strip() == ""
    else (
        "EMAIL"
        if pd.isna(row["S_SUPP_SITE_NOTIF_METHOD"])
           or str(row["S_SUPP_SITE_NOTIF_METHOD"]).strip() == ""
        else row["S_SUPP_SITE_NOTIF_METHOD"]
    ),
    axis=1
)

# ----------------------------------------------------------
# VALIDATE NOTIFICATION METHOD
# ----------------------------------------------------------

invalid_notification_mask = (
    df["S_SUPP_SITE_NOTIF_METHOD"]
    .fillna("")
    .astype(str)
    .str.strip()
    .str.upper()
    != ""
) & (
    df["S_SUPP_SITE_NOTIF_METHOD"]
    .fillna("")
    .astype(str)
    .str.strip()
    .str.upper()
    != "EMAIL"
)

df.loc[
    invalid_notification_mask,
    "FAILURE_REASON"
] += (
    "S_SUPP_SITE_NOTIF_METHOD must be EMAIL when "
    "S_EMAIL_ADDRESS is provided | "
)

df.loc[
    invalid_notification_mask,
    "ERROR_COLUMNS"
] += (
    "S_SUPP_SITE_NOTIF_METHOD|"
)

df.loc[
    invalid_notification_mask,
    "WARNING_MESSAGE"
] += (
    "Invalid Supplier Notification Method | "
)

df.loc[
    invalid_notification_mask,
    "WARNING_COLUMNS"
] += (
    "S_SUPP_SITE_NOTIF_METHOD|"
)

# ==========================================================
# RULE 8
# OPERATING UNIT MUST BE SWBR SS OU
# ==========================================================

swbr_mask = (
    df["S_OPERATING_UNIT"]
    .fillna("")
    .astype(str)
    .str.strip()
    != "SWBR SS OU"
)

df.loc[
    swbr_mask,
    "FAILURE_REASON"
] += (
    "S_OPERATING_UNIT must be SWBR SS OU | "
)

df.loc[
    swbr_mask,
    "ERROR_COLUMNS"
] += (
    "S_OPERATING_UNIT|"
)


# ==========================================================
# RULE 9
# BRAZIL CATEGORY + ATTRIBUTE DEPENDENCY
# ==========================================================

# ----------------------------------------------------------
# Condition 1
# For Brazil records check category value
# ----------------------------------------------------------

br_country_mask = (
    df["S_COUNTRY"]
    .fillna("")
    .astype(str)
    .str.strip()
    .str.upper()
    == "BR"
)

invalid_category_mask = (
    br_country_mask
    &
    (
        df["S_SITE_GLOBAL_ATTR_CATEGORY"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        != "JL.BR.APXVDMVD.SITES"
    )
)

df.loc[
    invalid_category_mask,
    "FAILURE_REASON"
] += (
    "S_SITE_GLOBAL_ATTR_CATEGORY must be "
    "JL.BR.APXVDMVD.SITES for Brazil | "
)

df.loc[
    invalid_category_mask,
    "ERROR_COLUMNS"
] += (
    "S_SITE_GLOBAL_ATTR_CATEGORY|"
)

# ----------------------------------------------------------
# Condition 2
# Attribute Dependency Check
# ----------------------------------------------------------

ga9 = (
    df["S_SITE_GLOBAL_ATTRIBUTE9"]
    .fillna("")
    .astype(str)
    .str.strip()
)

ga10 = (
    df["S_SITE_GLOBAL_ATTRIBUTE10"]
    .fillna("")
    .astype(str)
    .str.strip()
)

ga11 = (
    df["S_SITE_GLOBAL_ATTRIBUTE11"]
    .fillna("")
    .astype(str)
    .str.strip()
)

ga12 = (
    df["S_SITE_GLOBAL_ATTRIBUTE12"]
    .fillna("")
    .astype(str)
    .str.strip()
)

dependency_mask = (
    (ga9 != "")
    &
    (
        (ga10 == "")
        |
        (ga11 == "")
        |
        (ga12 == "")
    )
)

df.loc[
    dependency_mask,
    "FAILURE_REASON"
] += (
    "S_SITE_GLOBAL_ATTRIBUTE10, "
    "S_SITE_GLOBAL_ATTRIBUTE11, and "
    "S_SITE_GLOBAL_ATTRIBUTE12 are required "
    "when S_SITE_GLOBAL_ATTRIBUTE9 is populated "
    "for Suppliers in Brazil | "
)

df.loc[
    dependency_mask,
    "ERROR_COLUMNS"
] += (
    "S_SITE_GLOBAL_ATTRIBUTE9|"
    "S_SITE_GLOBAL_ATTRIBUTE10|"
    "S_SITE_GLOBAL_ATTRIBUTE11|"
    "S_SITE_GLOBAL_ATTRIBUTE12|"
)

df.loc[
    dependency_mask,
    "WARNING_MESSAGE"
] += (
    "Missing dependent Site Global Attributes | "
)

df.loc[
    dependency_mask,
    "WARNING_COLUMNS"
] += (
    "S_SITE_GLOBAL_ATTRIBUTE10|"
    "S_SITE_GLOBAL_ATTRIBUTE11|"
    "S_SITE_GLOBAL_ATTRIBUTE12|"
)
# ==========================================================
# RULE 10
# BANKING ATTRIBUTE DEPENDENCY
# ==========================================================

bank_columns = [
    "S_BANK_NAME",
    "S_BANK_NUM",
    "S_BANK_COUNTRY",
    "S_BANK_BRANCH_NAME",
    "S_BANK_BRANCH_NUM",
    "S_BANK_BRANCH_TYPE",
    "S_ACCOUNT_NUM",
    "S_ACCOUNT_NAME"
]

# Check if ANY banking field is populated
any_bank_value = (
    df[bank_columns]
    .fillna("")
    .astype(str)
    .apply(lambda col: col.str.strip())
    .ne("")
    .any(axis=1)
)

# Check if ALL banking fields are populated
all_bank_values = (
    df[bank_columns]
    .fillna("")
    .astype(str)
    .apply(lambda col: col.str.strip())
    .ne("")
    .all(axis=1)
)

bank_dependency_mask = (
    any_bank_value
    &
    ~all_bank_values
)

df.loc[
    bank_dependency_mask,
    "FAILURE_REASON"
] += (
    "All bank-related fields are required when banking information is provided. | "
)

df.loc[
    bank_dependency_mask,
    "ERROR_COLUMNS"
] += (
    "S_BANK_NAME|"
    "S_BANK_NUM|"
    "S_BANK_COUNTRY|"
    "S_BANK_BRANCH_NAME|"
    "S_BANK_BRANCH_NUM|"
    "S_BANK_BRANCH_TYPE|"
    "S_ACCOUNT_NUM|"
    "S_ACCOUNT_NAME|"
)

# ==========================================================
# RULE 11
# A_ADDRESS_LINE1 SPECIAL CHARACTER VALIDATION
# ==========================================================

invalid_address_mask = (
    df["A_ADDRESS_LINE1"]
    .fillna("")
    .astype(str)
    .str.contains(
        r"[^A-Za-z0-9\s\.\,\-\/&()]",
        regex=True
    )
)

df.loc[
    invalid_address_mask,
    "FAILURE_REASON"
] += (
    "A_ADDRESS_LINE1 contains invalid characters. "
    "Special characters are not allowed. | "
)

df.loc[
    invalid_address_mask,
    "ERROR_COLUMNS"
] += (
    "A_ADDRESS_LINE1|"
)
# ==========================================================
# RULE 12
# PAYMENT PRIORITY / PAY GROUP WARNING
# ==========================================================

pay_warning_mask = (
    (
        df["S_SITE_PAY_PAYMENT_PRIORITY"]
        .fillna("")
        .astype(str)
        .str.strip()
        == ""
    )
    |
    (
        df["S_SITE_PAY_PAY_GROUP"]
        .fillna("")
        .astype(str)
        .str.strip()
        == ""
    )
)

df.loc[
    pay_warning_mask,
    "WARNING_MESSAGE"
] += (
    "S_SITE_PAY_PAYMENT_PRIORITY or "
    "S_SITE_PAY_PAY_GROUP is blank | "
)

df.loc[
    pay_warning_mask,
    "WARNING_COLUMNS"
] += (
    "S_SITE_PAY_PAYMENT_PRIORITY|"
    "S_SITE_PAY_PAY_GROUP|"
)

# ==========================================================
# RULE 13
# COUNTRY SPECIFIC ADDRESS VALIDATION
# ==========================================================

address_rule_df = pd.read_excel(
    "config/Country_Address_Ruless.xlsx"
)

mandatory_address_rules = (
    address_rule_df[
        address_rule_df["REQUIRED_FLAG"]
        .astype(str)
        .str.strip()
        .str.upper()
        == "Y"
    ]
)

for idx, row in df.iterrows():

    country = str(
        row["S_COUNTRY"]
    ).strip().upper()

    country_rules = mandatory_address_rules[
        mandatory_address_rules["S_COUNTRY"]
        .astype(str)
        .str.strip()
        .str.upper()
        == country
    ]

    for _, rule in country_rules.iterrows():

        field_name = str(
            rule["Column Name"]
        ).strip()

        if field_name not in df.columns:
            continue

        field_value = str(
            row.get(field_name, "")
        ).strip()

        if (
            pd.isna(row.get(field_name))
            or field_value == ""
            or field_value.upper() == "NAN"
        ):

            df.loc[
                idx,
                "FAILURE_REASON"
            ] += (
                f"{field_name} mandatory "
                f"for country {country} | "
            )

            df.loc[
                idx,
                "ERROR_COLUMNS"
            ] += (
                f"{field_name}|"
            )

# ==========================================================
# RULE 14
# PAYMENT TERMS LOV VALIDATION
# ==========================================================

payment_terms_df = pd.read_excel(
    "config/PAY_TERMS.xlsx"
)

valid_payment_terms = set(
    payment_terms_df["S_SITE_PAY_TERMS"]
    .fillna("")
    .astype(str)
    .str.strip()
    .str.upper()
)

source_payment_terms = (
    df["S_SITE_PAY_TERMS"]
    .fillna("")
    .astype(str)
    .str.strip()
    .str.upper()
)

invalid_pay_term_mask = (
    (source_payment_terms != "")
    &
    (~source_payment_terms.isin(valid_payment_terms))
)

df.loc[
    invalid_pay_term_mask,
    "FAILURE_REASON"
] += (
    "S_SITE_PAY_TERMS not found in PAYMENT_TERM LOV | "
)

df.loc[
    invalid_pay_term_mask,
    "ERROR_COLUMNS"
] += (
    "S_SITE_PAY_TERMS|"
)

df.loc[
    invalid_pay_term_mask,
    "WARNING_MESSAGE"
] += (
    "Invalid Payment Term - review PAYMENT_TERM reference | "
)

df.loc[
    invalid_pay_term_mask,
    "WARNING_COLUMNS"
] += (
    "S_SITE_PAY_TERMS|"
)
# ==========================================================
# RULE 15
# PAYMENT METHOD LOV VALIDATION
# ==========================================================

# Load Payment Method Reference File

payment_method_df = pd.read_excel(
    "config/PAY_METHOD.xlsx"
)

# Create LOV Set

valid_payment_methods = set(
    payment_method_df["S_DEF_SITE_PAY_METHOD"]
    .fillna("")
    .astype(str)
    .str.strip()
    .str.upper()
)

# Source Values

source_payment_methods = (
    df["S_DEF_SITE_PAY_METHOD"]
    .fillna("")
    .astype(str)
    .str.strip()
    .str.upper()
)

invalid_payment_method_mask = (
    (source_payment_methods != "")
    &
    (~source_payment_methods.isin(valid_payment_methods))
)

# Validation Error

df.loc[
    invalid_payment_method_mask,
    "FAILURE_REASON"
] += (
    "S_DEF_SITE_PAY_METHOD not found in PAYMENT_METHOD LOV | "
)

df.loc[
    invalid_payment_method_mask,
    "ERROR_COLUMNS"
] += (
    "S_DEF_SITE_PAY_METHOD|"
)

# Warning Message

df.loc[
    invalid_payment_method_mask,
    "WARNING_MESSAGE"
] += (
    "Invalid Payment Method - review PAYMENT_METHOD reference | "
)

df.loc[
    invalid_payment_method_mask,
    "WARNING_COLUMNS"
] += (
    "S_DEF_SITE_PAY_METHOD|"
)

# ==========================================================
# RULE 16
# SUPPLIER TYPE LOV VALIDATION
# ==========================================================

# Load Supplier Type Reference File

supplier_type_df = pd.read_excel(
    "config/SUPPLIER_TYPE.xlsx"
)

# Create LOV Set

valid_supplier_types = set(
    supplier_type_df["H_SUPPLIER_TYPE"]
    .fillna("")
    .astype(str)
    .str.strip()
    .str.upper()
)

# Source Values

source_supplier_types = (
    df["H_SUPPLIER_TYPE"]
    .fillna("")
    .astype(str)
    .str.strip()
    .str.upper()
)

invalid_supplier_type_mask = (
    (source_supplier_types != "")
    &
    (~source_supplier_types.isin(valid_supplier_types))
)

# Validation Error

df.loc[
    invalid_supplier_type_mask,
    "FAILURE_REASON"
] += (
    "H_SUPPLIER_TYPE not found in SUPPLIER_TYPE LOV | "
)

df.loc[
    invalid_supplier_type_mask,
    "ERROR_COLUMNS"
] += (
    "H_SUPPLIER_TYPE|"
)

# Warning Message

df.loc[
    invalid_supplier_type_mask,
    "WARNING_MESSAGE"
] += (
    "Invalid Supplier Type - review SUPPLIER_TYPE reference | "
)

df.loc[
    invalid_supplier_type_mask,
    "WARNING_COLUMNS"
] += (
    "H_SUPPLIER_TYPE|"
)

# ==========================================================
# RULE 17
# ARGENTINA LOCALIZATION VALIDATION
# ==========================================================

argentina_category_mask = (
    df["S_SITE_GLOBAL_ATTR_CATEGORY"]
    .fillna("")
    .astype(str)
    .str.strip()
    .str.upper()
    == "JL.AR.APXVDMVD.SUPPLIER_SITES"
)

argentina_attr17_mask = (
    argentina_category_mask
    &
    (
        df["S_SITE_GLOBAL_ATTRIBUTE17"]
        .fillna("")
        .astype(str)
        .str.strip()
        == ""
    )
)

df.loc[
    argentina_attr17_mask,
    "FAILURE_REASON"
] += (
    "S_SITE_GLOBAL_ATTRIBUTE17 mandatory when "
    "S_SITE_GLOBAL_ATTR_CATEGORY = "
    "JL.AR.APXVDMVD.SUPPLIER_SITES | "
)

df.loc[
    argentina_attr17_mask,
    "ERROR_COLUMNS"
] += (
    "S_SITE_GLOBAL_ATTR_CATEGORY|"
    "S_SITE_GLOBAL_ATTRIBUTE17|"
)
# ----------------------------------------------------------
# ATTRIBUTE17 MANDATORY
# ----------------------------------------------------------

attribute17_missing_mask = (
    argentina_category_mask
    &
    (
        df["S_SITE_GLOBAL_ATTRIBUTE17"]
        .fillna("")
        .astype(str)
        .str.strip()
        == ""
    )
)

df.loc[
    attribute17_missing_mask,
    "FAILURE_REASON"
] += (
    "S_SITE_GLOBAL_ATTRIBUTE17 mandatory when "
    "S_SITE_GLOBAL_ATTR_CATEGORY = "
    "JL.AR.APXVDMVD.SUPPLIER_SITES | "
)

df.loc[
    attribute17_missing_mask,
    "ERROR_COLUMNS"
] += (
    "S_SITE_GLOBAL_ATTR_CATEGORY|"
    "S_SITE_GLOBAL_ATTRIBUTE17|"
)

df.loc[
    attribute17_missing_mask,
    "WARNING_MESSAGE"
] += (
    "Argentina Localization Warning - "
    "S_SITE_GLOBAL_ATTRIBUTE17 missing | "
)

df.loc[
    attribute17_missing_mask,
    "WARNING_COLUMNS"
] += (
    "S_SITE_GLOBAL_ATTRIBUTE17|"
)


# ==========================================================
# RULE 18
# NO FIELD SHOULD CONTAIN #N/A
# ==========================================================

for column in df.columns:

    na_mask = (
        df[column]
        .astype(str)
        .str.strip()
        .str.upper()
        .isin([
            "#N/A",
            "N/A",
            "#NA"
        ])
    )

    df.loc[
        na_mask,
        "FAILURE_REASON"
    ] += (
        f"{column} contains #N/A | "
    )

    df.loc[
        na_mask,
        "ERROR_COLUMNS"
    ] += (
        f"{column}|"
    )

    df.loc[
        na_mask,
        "WARNING_MESSAGE"
    ] += (
        f"{column} contains invalid #N/A value | "
    )

    df.loc[
        na_mask,
        "WARNING_COLUMNS"
    ] += (
        f"{column}|"
    )

# ==========================================================
# RULE 19
# BRAZIL SUPPLIER NAME ALT WARNING
# ==========================================================

br_alt_name_mask = (
    (
        df["S_COUNTRY"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        == "BR"
    )
    &
    (
        df["H_SUPPLIER_NAME_ALT"]
        .fillna("")
        .astype(str)
        .str.strip()
        != ""
    )
)

df.loc[
    br_alt_name_mask,
    "WARNING_MESSAGE"
] += (
    "H_SUPPLIER_NAME_ALT should be blank for Brazil suppliers | "
)

df.loc[
    br_alt_name_mask,
    "WARNING_COLUMNS"
] += (
    "H_SUPPLIER_NAME_ALT|"
)

# ==========================================================
# RULE 20
# INACTIVE SUPPLIER WARNING
# ==========================================================

today = pd.Timestamp.today().normalize()

h_end_date = pd.to_datetime(
    df["H_END_DATE_ACTIVE"],
    errors="coerce"
)

s_inactive_date = pd.to_datetime(
    df["S_INACTIVE_DATE"],
    errors="coerce"
)

inactive_supplier_mask = (
    (
        h_end_date.notna()
    )
    &
    (
        h_end_date < today
    )
) | (
    (
        s_inactive_date.notna()
    )
    &
    (
        s_inactive_date < today
    )
)

df.loc[
    inactive_supplier_mask,
    "WARNING_MESSAGE"
] += (
    "Supplier appears to be inactive based on "
    "H_END_DATE_ACTIVE or S_INACTIVE_DATE. "
    "Please verify whether reactivation is required "
    "before submitting the request | "
)

df.loc[
    inactive_supplier_mask,
    "WARNING_COLUMNS"
] += (
    "H_END_DATE_ACTIVE|"
    "S_INACTIVE_DATE|"
)

# ==========================================================
# RULE 21
# BRAZIL DOCUMENT TYPE TRANSFORMATION
# ==========================================================

doc_type_df = pd.read_excel(
    "config/BR_DOCUMENT_TYPE.xlsx"
)

doc_type_mapping = dict(
    zip(
        doc_type_df["S_SITE_GLOBAL_ATTRIBUTE9"]
        .astype(str)
        .str.strip()
        .str.upper(),

        doc_type_df["Value"]
    )
)

br_mask = (
    df["S_COUNTRY"]
    .fillna("")
    .astype(str)
    .str.strip()
    .str.upper()
    == "BR"
)

for idx in df[br_mask].index:

    value = (
        str(
            df.at[
                idx,
                "S_SITE_GLOBAL_ATTRIBUTE9"
            ]
        )
        .strip()
        .upper()
    )

    if value in doc_type_mapping:

        df.at[
            idx,
            "S_SITE_GLOBAL_ATTRIBUTE9"
        ] = doc_type_mapping[value]

    else:

        df.at[
            idx,
            "S_SITE_GLOBAL_ATTRIBUTE9"
        ] = "3"

# ==========================================================
# RULE 22
# ATTRIBUTE10 MUST BE 9 DIGITS
# ==========================================================

attribute10_mask = (
    df["S_SITE_GLOBAL_ATTRIBUTE10"]
    .fillna("")
    .astype(str)
    .str.strip()
    != ""
)

# ------------------------------------------
# LENGTH > 9 → ERROR
# ------------------------------------------

attribute10_error_mask = (
    attribute10_mask
    &
    (
        df["S_SITE_GLOBAL_ATTRIBUTE10"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.len()
        > 9
    )
)

df.loc[
    attribute10_error_mask,
    "FAILURE_REASON"
] += (
    "S_SITE_GLOBAL_ATTRIBUTE10 must not exceed 9 digits | "
)

df.loc[
    attribute10_error_mask,
    "ERROR_COLUMNS"
] += (
    "S_SITE_GLOBAL_ATTRIBUTE10|"
)

# ------------------------------------------
# LENGTH <= 9 → PREFIX ZEROES
# ------------------------------------------

attribute10_transform_mask = (
    attribute10_mask
    &
    (
        df["S_SITE_GLOBAL_ATTRIBUTE10"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.len()
        <= 9
    )
)

df.loc[
    attribute10_transform_mask,
    "S_SITE_GLOBAL_ATTRIBUTE10"
] = (
    df.loc[
        attribute10_transform_mask,
        "S_SITE_GLOBAL_ATTRIBUTE10"
    ]
    .astype(str)
    .str.strip()
    .str.zfill(9)
)

# ==========================================================
# RULE 23
# RFQ FLAG VALIDATION
# ==========================================================

rfq_flag_mask = (

    df["S_RFQ_ONLY_SITE_FLAG(Y/N)"]
    .fillna("")
    .astype(str)
    .str.strip()
    .str.upper()
    == "Y"

) | (

    df["A_PURPOSE_RFQ_FLAG(Y/N)"]
    .fillna("")
    .astype(str)
    .str.strip()
    .str.upper()
    == "Y"

)

df.loc[
    rfq_flag_mask,
    "FAILURE_REASON"
] += (
    "RFQ enabled records are not allowed. "
    "S_RFQ_ONLY_SITE_FLAG or A_PURPOSE_RFQ_FLAG is set to Y | "
)

df.loc[
    rfq_flag_mask,
    "ERROR_COLUMNS"
] += (
    "S_RFQ_ONLY_SITE_FLAG(Y/N)|"
    "A_PURPOSE_RFQ_FLAG(Y/N)|"
)

df.loc[
    rfq_flag_mask,
    "WARNING_MESSAGE"
] += (
    "RFQ Flag Validation Failed | "
)

df.loc[
    rfq_flag_mask,
    "WARNING_COLUMNS"
] += (
    "S_RFQ_ONLY_SITE_FLAG(Y/N)|"
    "A_PURPOSE_RFQ_FLAG(Y/N)|"
)

# ==========================================================
# RULE 24
# SUPPLIER SITE NAME VS ADDRESS NAME VALIDATION
# ==========================================================

site_address_mismatch_mask = (

    df["S_SUPPLIER_SITE_NAME"]
    .fillna("")
    .astype(str)
    .str.strip()
    != ""

) & (

    df["A_ADDRESS_NAME"]
    .fillna("")
    .astype(str)
    .str.strip()
    != ""

) & (

    df["S_SUPPLIER_SITE_NAME"]
    .fillna("")
    .astype(str)
    .str.strip()
    .str.upper()

    !=

    df["A_ADDRESS_NAME"]
    .fillna("")
    .astype(str)
    .str.strip()
    .str.upper()

)

df.loc[
    site_address_mismatch_mask,
    "WARNING_MESSAGE"
] += (
    "Supplier Site Name and Address Name do not match | "
)

df.loc[
    site_address_mismatch_mask,
    "WARNING_COLUMNS"
] += (
    "S_SUPPLIER_SITE_NAME|"
    "A_ADDRESS_NAME|"
)
# ==========================================================
# RULE 25B
# COUNTRY / STATE / CITY LOOKUP VALIDATION
# ==========================================================

country_city_state_df = pd.read_excel(
    "config/COUNTRY_STATE_CITY_COMBINATION.xlsx"
)

valid_countries = set(
    country_city_state_df["S_COUNTRY"]
    .fillna("")
    .astype(str)
    .str.strip()
)

valid_cities = set(
    country_city_state_df["A_CITY"]
    .fillna("")
    .astype(str)
    .str.strip()
)

valid_states = set(
    country_city_state_df["A_STATE"]
    .fillna("")
    .astype(str)
    .str.strip()
)

# ----------------------------------------------------------
# COUNTRY VALIDATION
# ----------------------------------------------------------

invalid_country_mask = (
    df["S_COUNTRY"]
    .fillna("")
    .astype(str)
    .str.strip()
    != ""
) & (
    ~df["S_COUNTRY"]
    .fillna("")
    .astype(str)
    .str.strip()
    .isin(valid_countries)
)

df.loc[
    invalid_country_mask,
    "WARNING_MESSAGE"
] += (
    "Invalid Country in Country Lookup | "
)

df.loc[
    invalid_country_mask,
    "WARNING_COLUMNS"
] += (
    "S_COUNTRY|"
)

# ----------------------------------------------------------
# CITY VALIDATION
# ----------------------------------------------------------

invalid_city_mask = (
    df["A_CITY"]
    .fillna("")
    .astype(str)
    .str.strip()
    != ""
) & (
    ~df["A_CITY"]
    .fillna("")
    .astype(str)
    .str.strip()
    .isin(valid_cities)
)

df.loc[
    invalid_city_mask,
    "WARNING_MESSAGE"
] += (
    "Invalid City in City Lookup | "
)

df.loc[
    invalid_city_mask,
    "WARNING_COLUMNS"
] += (
    "A_CITY|"
)

# ----------------------------------------------------------
# STATE VALIDATION
# ----------------------------------------------------------

invalid_state_mask = (
    df["A_STATE"]
    .fillna("")
    .astype(str)
    .str.strip()
    != ""
) & (
    ~df["A_STATE"]
    .fillna("")
    .astype(str)
    .str.strip()
    .isin(valid_states)
)

df.loc[
    invalid_state_mask,
    "WARNING_MESSAGE"
] += (
    "Invalid State in State Lookup | "
)

df.loc[
    invalid_state_mask,
    "WARNING_COLUMNS"
] += (
    "A_STATE|"
)

# ==========================================================
# RULE 26
# HEADER ATTRIBUTE CONSISTENCY VALIDATION
# ==========================================================

consistency_columns = [
    "H_SUPPLIER_NAME_ALT",
    "H_TAX_PAYER_ID",
    "H_TAX_REGISTRATION_NUMBER",
    "H_SUPPLIER_TYPE"
]

for column in consistency_columns:

    inconsistent_suppliers = (

        df.groupby("H_ATTRIBUTE15")[column]
        .apply(
            lambda x:
            x.fillna("")
             .astype(str)
             .str.strip()
             .replace("", pd.NA)
             .dropna()
             .nunique()
        )
        .loc[lambda x: x > 1]
        .index
    )

    inconsistency_mask = (
        df["H_ATTRIBUTE15"]
        .isin(inconsistent_suppliers)
    )

    df.loc[
        inconsistency_mask,
        "FAILURE_REASON"
    ] += (
        f"{column} is inconsistent for the same supplier | "
    )

    df.loc[
        inconsistency_mask,
        "ERROR_COLUMNS"
    ] += (
        f"{column}|"
    )

# ==========================================================
# RULE 27
# DUPLICATE TAX PAYER ID ACROSS SUPPLIERS
# ==========================================================

tax_id_duplicates = (

    df.groupby("H_TAX_PAYER_ID")["H_SUPPLIER_NAME"]
    .nunique()

)

duplicate_tax_ids = tax_id_duplicates[
    tax_id_duplicates > 1
].index

duplicate_tax_id_mask = (

    df["H_TAX_PAYER_ID"]
    .fillna("")
    .astype(str)
    .str.strip()
    .isin(
        [
            str(x).strip()
            for x in duplicate_tax_ids
        ]
    )

) & (

    df["H_TAX_PAYER_ID"]
    .fillna("")
    .astype(str)
    .str.strip()
    != ""

)

df.loc[
    duplicate_tax_id_mask,
    "FAILURE_REASON"
] += (
    "H_TAX_PAYER_ID is assigned to multiple suppliers | "
)

df.loc[
    duplicate_tax_id_mask,
    "ERROR_COLUMNS"
] += (
    "H_TAX_PAYER_ID|"
)

from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

# ==========================================================
# SPLIT VALID / INVALID
# ==========================================================

error_df = df[
    df["FAILURE_REASON"]
    .fillna("")
    .astype(str)
    .str.strip()
    != ""
].copy()

valid_df = df[
    (
        df["FAILURE_REASON"]
        .fillna("")
        .astype(str)
        .str.strip()
        == ""
    )
    &
    (
        df["WARNING_MESSAGE"]
        .fillna("")
        .astype(str)
        .str.strip()
        == ""
    )
].copy()

# ==========================================================
# WARNING DATAFRAME
# ==========================================================

warning_df = df[
    df["WARNING_MESSAGE"]
    .fillna("")
    .astype(str)
    .str.strip()
    != ""
].copy()

# ==========================================================
# CONSOLIDATED VALIDATION REPORT
# ERRORS + WARNINGS
# ==========================================================

validation_report_df = df[
    (
        df["FAILURE_REASON"]
        .fillna("")
        .astype(str)
        .str.strip()
        != ""
    )
    |
    (
        df["WARNING_MESSAGE"]
        .fillna("")
        .astype(str)
        .str.strip()
        != ""
    )
].copy()

# ==========================================================
# WARNING DATAFRAME
# ==========================================================

warning_df = df[
    df["WARNING_MESSAGE"]
    .fillna("")
    .astype(str)
    .str.strip()
    != ""
].copy()

# ==========================================================
# ADD SOURCE ROW NUMBER
# ==========================================================

validation_report_df.insert(
    0,
    "SOURCE_ROW_NUMBER",
    validation_report_df.index + 2
)

warning_df.insert(
    0,
    "SOURCE_ROW_NUMBER",
    warning_df.index + 2
)

# ==========================================================
# SAVE OUTPUTS
# ==========================================================

valid_df = valid_df.drop(
    columns=[
        "FAILURE_REASON",
        "WARNING_MESSAGE",
        "ERROR_COLUMNS",
        "WARNING_COLUMNS"
    ],
    errors="ignore"
)

valid_df.to_excel(
    "output/Supplier_Output.xlsx",
    index=False
)

# ==========================================================
# CONSOLIDATED VALIDATION REPORT
# ==========================================================

validation_report_df.to_excel(
    "output/Supplier_Validation_Report.xlsx",
    index=False
)

# ==========================================================
# COLOR ERROR + WARNING CELLS
# ==========================================================

error_fill = PatternFill(
    start_color="F4B084",
    end_color="F4B084",
    fill_type="solid"
)

warning_fill = PatternFill(
    start_color="FFE699",
    end_color="FFE699",
    fill_type="solid"
)

wb = load_workbook(
    "output/Supplier_Validation_Report.xlsx"
)

ws = wb.active

headers = {}

for col in range(1, ws.max_column + 1):

    headers[
        ws.cell(1, col).value
    ] = col

# ----------------------------------------------------------
# COLOR ERROR CELLS (RED)
# ----------------------------------------------------------

if "ERROR_COLUMNS" in headers:

    for row in range(2, ws.max_row + 1):

        error_cols = str(
            ws.cell(
                row,
                headers["ERROR_COLUMNS"]
            ).value or ""
        )

        for col_name in error_cols.split("|"):

            col_name = col_name.strip()

            if (
                col_name != ""
                and
                col_name in headers
            ):

                ws.cell(
                    row,
                    headers[col_name]
                ).fill = error_fill

# ----------------------------------------------------------
# COLOR WARNING CELLS (YELLOW)
# ----------------------------------------------------------

if "WARNING_COLUMNS" in headers:

    for row in range(2, ws.max_row + 1):

        warning_cols = str(
            ws.cell(
                row,
                headers["WARNING_COLUMNS"]
            ).value or ""
        )

        for col_name in warning_cols.split("|"):

            col_name = col_name.strip()

            if (
                col_name != ""
                and
                col_name in headers
            ):

                existing_fill = ws.cell(
                    row,
                    headers[col_name]
                ).fill

                # Do not overwrite error color
                if (
                    existing_fill.fill_type != "solid"
                    or
                    existing_fill.start_color.rgb
                    not in [
                        "00F4CCCC",
                        "FFF4CCCC"
                    ]
                ):

                    ws.cell(
                        row,
                        headers[col_name]
                    ).fill = warning_fill

# ----------------------------------------------------------
# HIDE HELPER COLUMNS
# ----------------------------------------------------------

if "ERROR_COLUMNS" in headers:

    col_letter = get_column_letter(
        headers["ERROR_COLUMNS"]
    )

    ws.column_dimensions[
        col_letter
    ].hidden = True

if "WARNING_COLUMNS" in headers:

    col_letter = get_column_letter(
        headers["WARNING_COLUMNS"]
    )

    ws.column_dimensions[
        col_letter
    ].hidden = True

# ----------------------------------------------------------
# LEGEND
# ----------------------------------------------------------

from openpyxl.styles import Font

ws.insert_rows(1, 2)

# Legend Title

ws["A1"] = "Legend:"
ws["A1"].font = Font(
    bold=True
)

# Error Indicator Cell

ws["B1"] = ""
ws["B1"].fill = error_fill

ws["C1"] = "Validation Error (Fix Required)"
ws["C1"].font = Font(
    bold=True
)

# Warning Indicator Cell

ws["D1"] = ""
ws["D1"].fill = warning_fill

ws["E1"] = "Warning (Review Recommended)"
ws["E1"].font = Font(
    bold=True
)

# Make Color Boxes Visible

ws.column_dimensions["B"].width = 4
ws.column_dimensions["D"].width = 4

# Freeze Header Row

ws.freeze_panes = "A4"

# Enable Auto Filter

ws.auto_filter.ref = ws.dimensions

wb.save(
    "output/Supplier_Validation_Report.xlsx"
)

# ==========================================================
# WARNING COUNT
# ==========================================================

warning_count = len(warning_df)

# ==========================================================
# SUMMARY
# ==========================================================

summary_df = pd.DataFrame({
    "Metric": [
        "Total Records",
        "Valid Records",
        "Rejected Records",
        "Warning Records"
    ],
    "Value": [
        len(df),
        len(valid_df),
        len(error_df),
        warning_count
    ]
})

summary_df.to_csv(
    "output/Validation_Summary.csv",
    index=False
)

# ==========================================================
# SUMMARY
# ==========================================================

print("=" * 60)
print("SUPPLIER VALIDATION COMPLETED")
print("=" * 60)
print("Total Records    :", len(df))
print("Valid Records    :", len(valid_df))
print("Rejected Records :", len(error_df))
print("Warning Records  :", warning_count)
print("")
print("Generated Files")
print("----------------")
print("output/Supplier_Output.xlsx")
print("output/Supplier_Validation_Report.xlsx")
print("output/Validation_Summary.csv")
print("output/Supplier_WebADI_Output.xlsx")
print("=" * 60)