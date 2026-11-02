import pandas as pd

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

    elif severity == "WARNING":

        df.loc[
            field_mask,
            "WARNING_MESSAGE"
        ] += (
            f"{field} Missing | "
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
    "Duplicate H_ATTRIBUTE15 and "
    "S_SUPPLIER_SITE_NAME combination | "
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
] += "S_SUPPLIER_SITE_NAME exceeds 15 characters | "

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
] += "S_SITE_PAY_PAYMENT_PRIORITY should be numeric | "

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
] += "CONTACT_FIRST_NAME not uppercase | "

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
] += "CONTACT_LAST_NAME not uppercase | "

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
    "S_SUPP_SITE_NOTIF_METHOD must be EMAIL | "
)

df.loc[
    invalid_notification_mask,
    "WARNING_MESSAGE"
] += (
    "Invalid Supplier Notification Method | "
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
] += "S_OPERATING_UNIT should be SWBR SS OU | "


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
    "S_SITE_GLOBAL_ATTRIBUTE9 populated but "
    "S_SITE_GLOBAL_ATTRIBUTE10/11/12 missing | "
)

df.loc[
    dependency_mask,
    "WARNING_MESSAGE"
] += (
    "Missing dependent Site Global Attributes | "
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
    "Incomplete Banking Information - "
    "all bank fields must be populated | "
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
        r"[^A-Za-z0-9\s\.,\-/&()]",
        regex=True
    )
)

df.loc[
    invalid_address_mask,
    "FAILURE_REASON"
] += (
    "A_ADDRESS_LINE1 contains invalid special characters | "
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

# ==========================================================
# RULE 13
# COUNTRY SPECIFIC ADDRESS VALIDATION
# ==========================================================

# ==========================================================
# RULE 13
# COUNTRY SPECIFIC ADDRESS VALIDATION
# ==========================================================

address_rule_df = pd.read_excel(
    "input/Country_Address_Ruless.xlsx"
)

# Take only mandatory rows

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

# ==========================================================
# RULE 14
# PAYMENT TERMS LOV VALIDATION
# ==========================================================

# Load Payment Terms Reference File

payment_terms_df = pd.read_excel(
    "input/PAY_TERMS.xlsx"
)

# Create LOV Set

valid_payment_terms = set(
    payment_terms_df["S_SITE_PAY_TERMS"]
    .fillna("")
    .astype(str)
    .str.strip()
    .str.upper()
)

# Source Values

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

# Validation Error

df.loc[
    invalid_pay_term_mask,
    "FAILURE_REASON"
] += (
    "S_SITE_PAY_TERMS not found in PAYMENT_TERM LOV | "
)

# Warning Message

df.loc[
    invalid_pay_term_mask,
    "WARNING_MESSAGE"
] += (
    "Invalid Payment Term - review PAYMENT_TERM reference | "
)

# ==========================================================
# RULE 15
# PAYMENT METHOD LOV VALIDATION
# ==========================================================

# Load Payment Method Reference File

payment_method_df = pd.read_excel(
    "input/PAY_METHOD.xlsx"
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

# Warning Message

df.loc[
    invalid_payment_method_mask,
    "WARNING_MESSAGE"
] += (
    "Invalid Payment Method - review PAYMENT_METHOD reference | "
)

# ==========================================================
# RULE 16
# SUPPLIER TYPE LOV VALIDATION
# ==========================================================

# Load Supplier Type Reference File

supplier_type_df = pd.read_excel(
    "input/SUPPLIER_TYPE.xlsx"
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

# Warning Message

df.loc[
    invalid_supplier_type_mask,
    "WARNING_MESSAGE"
] += (
    "Invalid Supplier Type - review SUPPLIER_TYPE reference | "
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
    "WARNING_MESSAGE"
] += (
    "Argentina Localization Warning - "
    "S_SITE_GLOBAL_ATTRIBUTE17 missing | "
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
        "WARNING_MESSAGE"
    ] += (
        f"{column} contains invalid #N/A value | "
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

# ==========================================================
# RULE 20
# INACTIVE SUPPLIER WARNING
# ==========================================================

inactive_supplier_mask = (
    df["S_INACTIVE_DATE"]
    .fillna("")
    .astype(str)
    .str.strip()
    != ""
)

df.loc[
    inactive_supplier_mask,
    "WARNING_MESSAGE"
] += (
    "Supplier is already inactive. "
    "Please verify whether reactivation is required "
    "before submitting the request | "
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
# LENGTH > 9  → ERROR
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
    "S_SITE_GLOBAL_ATTRIBUTE10 cannot exceed 9 digits | "
)

# ------------------------------------------
# LENGTH < 9 → PREFIX ZEROES
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
    "WARNING_MESSAGE"
] += (
    "RFQ Flag Validation Failed | "
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

# ==========================================================
# RULE 25B
# COUNTRY / STATE / CITY LOOKUP VALIDATION
# ==========================================================

country_city_state_df = pd.read_excel(
    "temp/COUNTRY_STATE_CITY_COMBINATION.xlsx"
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

# ==========================================================
# SPLIT VALID / INVALID
# ==========================================================

error_df = df[
    df["FAILURE_REASON"] != ""
].copy()

valid_df = df[
    df["FAILURE_REASON"] == ""
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

error_df.insert(
    0,
    "SOURCE_ROW_NUMBER",
    error_df.index + 2
)

# ==========================================================
# SAVE OUTPUTS
# ==========================================================

valid_df = valid_df.drop(
    columns=[
        "FAILURE_REASON",
        "WARNING_MESSAGE"
    ],
    errors="ignore"
)

valid_df.to_excel(
    "output/Supplier_Output.xlsx",
    index=False
)

error_df.to_csv(
    "output/Supplier_Validation_Error.csv",
    index=False
)

warning_df.to_csv(
    "output/Supplier_Warning_Report.csv",
    index=False
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
print("output/Supplier_Validation_Error.csv")
print("output/Supplier_Warning_Report.csv")
print("output/Validation_Summary.csv")
print("output/Supplier_WebADI_Output.xlsx")
print("=" * 60)