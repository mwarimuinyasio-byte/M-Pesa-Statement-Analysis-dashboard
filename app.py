import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import re


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="M-Pesa Statement Analysis",
    page_icon="M",
    layout="wide"
)


# ============================================================
# TITLE
# ============================================================

st.title("M-Pesa Statement Analysis Dashboard")

st.write(
    "Analyze M-Pesa money in, money out, transaction amounts, "
    "transaction ranges, and transactions to or from the same person."
)


# ============================================================
# FUNCTIONS
# ============================================================

def clean_money(value):
    """
    Convert values such as:
    KSh 1,000
    1,000.00
    -500
    into numeric values.
    """

    if pd.isna(value):
        return np.nan

    value = str(value)

    value = (
        value
        .replace(",", "")
        .replace("KSh", "")
        .replace("KES", "")
        .strip()
    )

    value = re.sub(
        r"[^\d.\-]",
        "",
        value
    )

    try:
        return float(value)
    except ValueError:
        return np.nan


def clean_columns(df):

    df = df.copy()

    df.columns = [
        re.sub(
            r"[^a-z0-9]+",
            "_",
            str(column).strip().lower()
        ).strip("_")
        for column in df.columns
    ]

    return df


def find_column(df, names):

    for name in names:

        name = (
            name
            .lower()
            .replace(" ", "_")
        )

        if name in df.columns:
            return name

    return None


def identify_transaction(details):

    text = str(details).lower()

    # ------------------------------
    # MONEY IN
    # ------------------------------

    money_in_keywords = [

        "funds received",

        "received from",

        "money received",

        "receive international",

        "received international",

        "customer payment to small",

        "customer payment"
    ]

    for keyword in money_in_keywords:

        if keyword in text:

            return "Money In"


    # ------------------------------
    # MONEY OUT
    # ------------------------------

    money_out_keywords = [

        "customer transfer to",

        "merchant payment",

        "pay bill",

        "airtime",

        "withdrawal",

        "customer withdrawal",

        "send money",

        "funds transfer to"
    ]

    for keyword in money_out_keywords:

        if keyword in text:

            return "Money Out"


    return "Other"


def extract_person(details):

    """
    Attempts to extract the person or business
    involved in the transaction.
    """

    text = str(details)

    patterns = [

        r"Customer Transfer to\s*-\s*.*?\s+([A-Za-z][A-Za-z .'-]+)",

        r"Funds received from\s*-\s*.*?\s+([A-Za-z][A-Za-z .'-]+)",

        r"Merchant Payment to\s*.*?-\s*([A-Za-z][A-Za-z .'-]+)",

        r"Customer Payment to Small Business to\s*-\s*.*?\s+([A-Za-z][A-Za-z .'-]+)",

        r"Pay Bill\s*.*?-\s*([A-Za-z][A-Za-z .'-]+)"
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            text,
            flags=re.IGNORECASE
        )

        if match:

            person = match.group(1).strip()

            if person:

                return person


    # If extraction fails, return original detail

    return text.strip()


def get_amount_range(amount):

    if pd.isna(amount):

        return "Unknown"

    if amount < 500:

        return "Below 500"

    elif amount < 1000:

        return "500 - 999"

    elif amount < 2000:

        return "1,000 - 1,999"

    elif amount < 5000:

        return "2,000 - 4,999"

    elif amount < 10000:

        return "5,000 - 9,999"

    else:

        return "10,000+"


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.title("M-Pesa Dashboard")

uploaded_file = st.sidebar.file_uploader(
    "Upload your M-Pesa Excel file",
    type=["xlsx", "xls"]
)


# ============================================================
# IF NO FILE
# ============================================================

if uploaded_file is None:

    st.info(
        "Upload the Excel file created from your M-Pesa statement."
    )

    st.markdown(
        """
        ### Expected workbook

        The workbook should contain:

        - **Transactions**
        - **Statement Summary**
        - **OCR Raw Rows**

        The Transactions sheet should contain columns such as:

        - Receipt No.
        - Completion Time
        - Details
        - Transaction Status
        - Paid In
        - Withdrawn
        - Balance
        - Transaction Type
        """
    )

    st.stop()


# ============================================================
# READ EXCEL FILE
# ============================================================

try:

    excel_file = pd.ExcelFile(
        uploaded_file
    )

except Exception as error:

    st.error(
        f"Unable to open Excel file: {error}"
    )

    st.stop()


# ============================================================
# READ TRANSACTIONS SHEET
# ============================================================

if "Transactions" in excel_file.sheet_names:

    df = pd.read_excel(
        uploaded_file,
        sheet_name="Transactions"
    )

else:

    st.error(
        "The Excel file does not contain a Transactions sheet."
    )

    st.stop()


# ============================================================
# READ SUMMARY SHEET
# ============================================================

if "Statement Summary" in excel_file.sheet_names:

    summary_df = pd.read_excel(
        uploaded_file,
        sheet_name="Statement Summary"
    )

else:

    summary_df = pd.DataFrame()


# ============================================================
# CLEAN COLUMN NAMES
# ============================================================

df = clean_columns(df)


# ============================================================
# IDENTIFY COLUMNS
# ============================================================

details_column = find_column(
    df,
    [
        "details",
        "description",
        "transaction_details"
    ]
)

date_column = find_column(
    df,
    [
        "completion_time",
        "completion_date",
        "transaction_date",
        "date"
    ]
)

paid_in_column = find_column(
    df,
    [
        "paid_in",
        "money_in"
    ]
)

withdrawn_column = find_column(
    df,
    [
        "withdrawn",
        "paid_out",
        "money_out"
    ]
)

balance_column = find_column(
    df,
    [
        "balance"
    ]
)

transaction_type_column = find_column(
    df,
    [
        "transaction_type"
    ]
)


# ============================================================
# CHECK DETAILS
# ============================================================

if details_column is None:

    st.error(
        "The Details column could not be found."
    )

    st.write(
        "Columns found:"
    )

    st.write(
        list(df.columns)
    )

    st.stop()


# ============================================================
# PREPARE DATA
# ============================================================

df["details"] = (
    df[details_column]
    .fillna("")
    .astype(str)
)


# ============================================================
# DATE
# ============================================================

if date_column:

    df["completion_time"] = pd.to_datetime(
        df[date_column],
        errors="coerce"
    )

else:

    df["completion_time"] = pd.NaT


# ============================================================
# PAID IN
# ============================================================

if paid_in_column:

    df["paid_in"] = df[
        paid_in_column
    ].apply(clean_money)

else:

    df["paid_in"] = 0.0


# ============================================================
# WITHDRAWN
# ============================================================

if withdrawn_column:

    df["withdrawn"] = df[
        withdrawn_column
    ].apply(clean_money)

else:

    df["withdrawn"] = 0.0


# ============================================================
# BALANCE
# ============================================================

if balance_column:

    df["balance"] = df[
        balance_column
    ].apply(clean_money)

else:

    df["balance"] = np.nan


# ============================================================
# TRANSACTION TYPE
# ============================================================

if transaction_type_column:

    df["transaction_type"] = (
        df[transaction_type_column]
        .fillna("")
        .astype(str)
    )

    # Reclassify rows marked Other where possible

    detected_types = df[
        "details"
    ].apply(
        identify_transaction
    )

    df.loc[
        df["transaction_type"].isin(
            ["", "nan", "Other"]
        ),
        "transaction_type"
    ] = detected_types

else:

    df["transaction_type"] = (
        df["details"]
        .apply(identify_transaction)
    )


# ============================================================
# MONEY IN
# ============================================================

df["money_in"] = (
    df["paid_in"]
    .fillna(0)
    .clip(lower=0)
)


# ============================================================
# MONEY OUT
# ============================================================

df["money_out"] = (
    df["withdrawn"]
    .fillna(0)
    .abs()
)


# ============================================================
# USE DETAILS TO DETECT AMOUNTS WHEN OCR MIXED COLUMNS
# ============================================================

# If transaction is Money In and Paid In is zero,
# try to find a +amount in the details.

def extract_amount_from_details(text):

    text = str(text)

    matches = re.findall(
        r"\+\s*([\d,]+(?:\.\d{1,2})?)",
        text
    )

    if matches:

        try:

            return float(
                matches[-1].replace(",", "")
            )

        except:

            return np.nan

    return np.nan


details_amount = df[
    "details"
].apply(
    extract_amount_from_details
)


df.loc[
    (df["transaction_type"] == "Money In")
    & (df["money_in"] == 0)
    & details_amount.notna(),
    "money_in"
] = details_amount


df.loc[
    (df["transaction_type"] == "Money Out")
    & (df["money_out"] == 0)
    & details_amount.notna(),
    "money_out"
] = details_amount


# ============================================================
# TRANSACTION AMOUNT
# ============================================================

df["amount"] = np.where(

    df["transaction_type"] == "Money In",

    df["money_in"],

    np.where(

        df["transaction_type"] == "Money Out",

        df["money_out"],

        0
    )
)


# ============================================================
# AMOUNT RANGE
# ============================================================

df["amount_range"] = (
    df["amount"]
    .apply(get_amount_range)
)


# ============================================================
# PERSON / BUSINESS
# ============================================================

df["person"] = (
    df["details"]
    .apply(extract_person)
)


# ============================================================
# CHARGE DETECTION
# ============================================================

df["is_charge"] = (
    df["details"]
    .str.contains(
        "charge|fee",
        case=False,
        na=False
    )
)


# ============================================================
# SIDEBAR FILTERS
# ============================================================

st.sidebar.subheader(
    "Transaction Filters"
)


transaction_types = sorted(
    df["transaction_type"]
    .dropna()
    .unique()
)


selected_types = st.sidebar.multiselect(

    "Transaction Type",

    options=transaction_types,

    default=transaction_types
)


range_order = [

    "Below 500",

    "500 - 999",

    "1,000 - 1,999",

    "2,000 - 4,999",

    "5,000 - 9,999",

    "10,000+"
]


selected_ranges = st.sidebar.multiselect(

    "Amount Range",

    options=range_order,

    default=range_order
)


# ============================================================
# DATE FILTER
# ============================================================

valid_dates = df[
    "completion_time"
].dropna()


if not valid_dates.empty:

    minimum_date = (
        valid_dates
        .min()
        .date()
    )

    maximum_date = (
        valid_dates
        .max()
        .date()
    )

    selected_dates = st.sidebar.date_input(

        "Date Range",

        value=(
            minimum_date,
            maximum_date
        ),

        min_value=minimum_date,

        max_value=maximum_date
    )

else:

    selected_dates = None


# ============================================================
# APPLY FILTERS
# ============================================================

filtered_df = df[
    df["transaction_type"].isin(
        selected_types
    )
    &
    df["amount_range"].isin(
        selected_ranges
    )
].copy()


if (
    selected_dates
    and len(selected_dates) == 2
):

    start_date = pd.Timestamp(
        selected_dates[0]
    )

    end_date = (
        pd.Timestamp(
            selected_dates[1]
        )
        + pd.Timedelta(days=1)
    )


    filtered_df = filtered_df[
        (
            filtered_df[
                "completion_time"
            ] >= start_date
        )
        &
        (
            filtered_df[
                "completion_time"
            ] < end_date
        )
    ]


# ============================================================
# HEADER
# ============================================================

st.success(
    f"{len(filtered_df):,} transactions selected."
)


# ============================================================
# KPI SECTION
# ============================================================

total_money_in = filtered_df[
    "money_in"
].sum()


total_money_out = filtered_df[
    "money_out"
].sum()


net_flow = (
    total_money_in -
    total_money_out
)


transaction_count = len(
    filtered_df
)


average_transaction = (

    filtered_df["amount"]
    .mean()
    if transaction_count > 0
    else 0
)


maximum_transaction = (

    filtered_df["amount"]
    .max()
    if transaction_count > 0
    else 0
)


st.subheader(
    "Financial Summary"
)


col1, col2, col3, col4 = st.columns(4)


col1.metric(
    "Money In",
    f"KSh {total_money_in:,.2f}"
)


col2.metric(
    "Money Out",
    f"KSh {total_money_out:,.2f}"
)


col3.metric(
    "Net Flow",
    f"KSh {net_flow:,.2f}"
)


col4.metric(
    "Transactions",
    f"{transaction_count:,}"
)


col5, col6 = st.columns(2)


col5.metric(
    "Average Transaction",
    f"KSh {average_transaction:,.2f}"
)


col6.metric(
    "Largest Transaction",
    f"KSh {maximum_transaction:,.2f}"
)


# ============================================================
# MONEY IN VS MONEY OUT
# ============================================================

st.subheader(
    "Money In vs Money Out"
)


flow_df = pd.DataFrame({

    "Flow": [
        "Money In",
        "Money Out"
    ],

    "Amount": [
        total_money_in,
        total_money_out
    ]
})


fig_flow = px.bar(

    flow_df,

    x="Flow",

    y="Amount",

    text_auto=".2f",

    title="Total Money In vs Money Out",

    labels={
        "Amount": "Amount (KSh)"
    },

    template="plotly_white"
)


st.plotly_chart(
    fig_flow,
    use_container_width=True
)


# ============================================================
# AMOUNT RANGE
# ============================================================

st.subheader(
    "Transactions by Amount Range"
)


range_summary = (

    filtered_df

    .groupby(
        "amount_range",
        as_index=False
    )

    .agg(

        Money_In=(
            "money_in",
            "sum"
        ),

        Money_Out=(
            "money_out",
            "sum"
        ),

        Transactions=(
            "amount",
            "count"
        )
    )
)


range_summary[
    "amount_range"
] = pd.Categorical(

    range_summary[
        "amount_range"
    ],

    categories=range_order,

    ordered=True
)


range_summary = (
    range_summary
    .sort_values(
        "amount_range"
    )
)


st.dataframe(

    range_summary,

    use_container_width=True,

    hide_index=True
)


fig_range = px.bar(

    range_summary,

    x="amount_range",

    y=[
        "Money_In",
        "Money_Out"
    ],

    barmode="group",

    title="Money In and Money Out by Amount Range",

    labels={
        "amount_range": "Amount Range",
        "value": "Amount (KSh)",
        "variable": "Flow"
    },

    template="plotly_white"
)


st.plotly_chart(

    fig_range,

    use_container_width=True
)


# ============================================================
# MONEY IN SOURCES
# ============================================================

st.subheader(
    "Money In — Who Sent Money to You?"
)


incoming = filtered_df[
    (
        filtered_df[
            "transaction_type"
        ] == "Money In"
    )
    &
    (~filtered_df["is_charge"])
].copy()


incoming_people = (

    incoming

    .groupby(
        "person",
        as_index=False
    )

    .agg(

        Total_Received=(
            "money_in",
            "sum"
        ),

        Number_of_Transactions=(
            "money_in",
            "count"
        )
    )

    .sort_values(
        "Total_Received",
        ascending=False
    )
)


if incoming_people.empty:

    st.warning(
        "No money-in transactions were found."
    )

else:

    st.dataframe(

        incoming_people.head(30),

        use_container_width=True,

        hide_index=True
    )


    fig_in = px.bar(

        incoming_people
        .head(15)
        .sort_values(
            "Total_Received"
        ),

        x="Total_Received",

        y="person",

        orientation="h",

        text_auto=".2f",

        title="Top People / Sources Sending Money",

        labels={
            "Total_Received": "Amount Received (KSh)",
            "person": "Person / Source"
        },

        template="plotly_white"
    )


    st.plotly_chart(

        fig_in,

        use_container_width=True
    )


# ============================================================
# MONEY OUT DESTINATIONS
# ============================================================

st.subheader(
    "Money Out — Who Did You Send Money To?"
)


outgoing = filtered_df[
    (
        filtered_df[
            "transaction_type"
        ] == "Money Out"
    )
    &
    (~filtered_df["is_charge"])
].copy()


outgoing_people = (

    outgoing

    .groupby(
        "person",
        as_index=False
    )

    .agg(

        Total_Sent=(
            "money_out",
            "sum"
        ),

        Number_of_Transactions=(
            "money_out",
            "count"
        )
    )

    .sort_values(
        "Total_Sent",
        ascending=False
    )
)


if outgoing_people.empty:

    st.warning(
        "No money-out transactions were found."
    )

else:

    st.dataframe(

        outgoing_people.head(30),

        use_container_width=True,

        hide_index=True
    )


    fig_out = px.bar(

        outgoing_people
        .head(15)
        .sort_values(
            "Total_Sent"
        ),

        x="Total_Sent",

        y="person",

        orientation="h",

        text_auto=".2f",

        title="Top People / Businesses You Paid",

        labels={
            "Total_Sent": "Amount Sent (KSh)",
            "person": "Person / Destination"
        },

        template="plotly_white"
    )


    st.plotly_chart(

        fig_out,

        use_container_width=True
    )


# ============================================================
# SAME PERSON ANALYSIS
# ============================================================

st.subheader(
    "People Who Both Sent Money to You and Received Money From You"
)


received = (

    incoming

    .groupby(
        "person",
        as_index=False
    )

    .agg(
        Money_Received=(
            "money_in",
            "sum"
        )
    )
)


sent = (

    outgoing

    .groupby(
        "person",
        as_index=False
    )

    .agg(
        Money_Sent=(
            "money_out",
            "sum"
        )
    )
)


same_person = pd.merge(

    received,

    sent,

    on="person",

    how="inner"
)


if same_person.empty:

    st.warning(
        "No person appears on both sides of the transactions."
    )

else:

    same_person[
        "Total_Activity"
    ] = (

        same_person[
            "Money_Received"
        ]

        +

        same_person[
            "Money_Sent"
        ]
    )


    same_person[
        "Net"
    ] = (

        same_person[
            "Money_Received"
        ]

        -

        same_person[
            "Money_Sent"
        ]
    )


    same_person = same_person.sort_values(

        "Total_Activity",

        ascending=False
    )


    st.dataframe(

        same_person.head(50),

        use_container_width=True,

        hide_index=True
    )


    fig_same = px.bar(

        same_person.head(15),

        x="person",

        y=[
            "Money_Received",
            "Money_Sent"
        ],

        barmode="group",

        title="Money Received vs Money Sent — Same Person",

        labels={
            "value": "Amount (KSh)",
            "person": "Person",
            "variable": "Direction"
        },

        template="plotly_white"
    )


    fig_same.update_xaxes(
        tickangle=-45
    )


    st.plotly_chart(

        fig_same,

        use_container_width=True
    )


# ============================================================
# TRANSACTION TYPE ANALYSIS
# ============================================================

st.subheader(
    "Transaction Type Analysis"
)


type_summary = (

    filtered_df

    .groupby(
        "transaction_type",
        as_index=False
    )

    .agg(

        Amount=(
            "amount",
            "sum"
        ),

        Transactions=(
            "amount",
            "count"
        )
    )
)


col1, col2 = st.columns(2)


with col1:

    fig_type_amount = px.bar(

        type_summary,

        x="transaction_type",

        y="Amount",

        text_auto=".2f",

        title="Amount by Transaction Type",

        labels={
            "transaction_type": "Transaction Type",
            "Amount": "Amount (KSh)"
        },

        template="plotly_white"
    )


    st.plotly_chart(

        fig_type_amount,

        use_container_width=True
    )


with col2:

    fig_type_count = px.pie(

        type_summary,

        names="transaction_type",

        values="Transactions",

        title="Transaction Count by Type",

        hole=0.35
    )


    st.plotly_chart(

        fig_type_count,

        use_container_width=True
    )


# ============================================================
# DAILY TRANSACTION ANALYSIS
# ============================================================

st.subheader(
    "Daily Money Movement"
)


timeline = filtered_df.dropna(
    subset=["completion_time"]
).copy()


if timeline.empty:

    st.warning(
        "There are no valid transaction dates."
    )

else:

    timeline["Date"] = (
        timeline[
            "completion_time"
        ].dt.date
    )


    daily = (

        timeline

        .groupby(
            "Date",
            as_index=False
        )

        .agg(

            Money_In=(
                "money_in",
                "sum"
            ),

            Money_Out=(
                "money_out",
                "sum"
            )
        )
    )


    fig_daily = px.line(

        daily,

        x="Date",

        y=[
            "Money_In",
            "Money_Out"
        ],

        markers=True,

        title="Daily Money In and Money Out",

        labels={
            "Date": "Date",
            "value": "Amount (KSh)",
            "variable": "Flow"
        },

        template="plotly_white"
    )


    st.plotly_chart(

        fig_daily,

        use_container_width=True
    )


# ============================================================
# TRANSACTION AMOUNT DISTRIBUTION
# ============================================================

st.subheader(
    "Transaction Amount Distribution"
)


if not filtered_df.empty:

    fig_histogram = px.histogram(

        filtered_df,

        x="amount",

        color="transaction_type",

        nbins=30,

        title="Distribution of Transaction Amounts",

        labels={
            "amount": "Transaction Amount (KSh)",
            "transaction_type": "Transaction Type"
        },

        template="plotly_white"
    )


    st.plotly_chart(

        fig_histogram,

        use_container_width=True
    )


# ============================================================
# STATEMENT SUMMARY
# ============================================================

st.subheader(
    "Original M-Pesa Statement Summary"
)


if not summary_df.empty:

    st.dataframe(

        summary_df,

        use_container_width=True,

        hide_index=True
    )


    # Summary graph

    summary_clean = summary_df.copy()


    if "Paid In" in summary_clean.columns:

        summary_clean["Paid In"] = pd.to_numeric(
            summary_clean["Paid In"],
            errors="coerce"
        )


    if "Paid Out" in summary_clean.columns:

        summary_clean["Paid Out"] = pd.to_numeric(
            summary_clean["Paid Out"],
            errors="coerce"
        )


    if (
        "Transaction Type" in summary_clean.columns
        and
        "Paid In" in summary_clean.columns
        and
        "Paid Out" in summary_clean.columns
    ):

        summary_long = summary_clean.melt(

            id_vars=[
                "Transaction Type"
            ],

            value_vars=[
                "Paid In",
                "Paid Out"
            ],

            var_name="Flow",

            value_name="Amount"
        )


        fig_summary = px.bar(

            summary_long,

            x="Transaction Type",

            y="Amount",

            color="Flow",

            barmode="group",

            title="Original Statement Summary",

            labels={
                "Amount": "Amount (KSh)"
            },

            template="plotly_white"
        )


        st.plotly_chart(

            fig_summary,

            use_container_width=True
        )


# ============================================================
# FULL TRANSACTION TABLE
# ============================================================

st.subheader(
    "Detailed Transactions"
)


display_columns = [

    "completion_time",

    "details",

    "transaction_type",

    "money_in",

    "money_out",

    "amount",

    "amount_range",

    "person",

    "balance"
]


display_columns = [

    column

    for column in display_columns

    if column in filtered_df.columns
]


display_df = filtered_df[
    display_columns
].sort_values(

    "completion_time",

    ascending=False,

    na_position="last"
)


st.dataframe(

    display_df,

    use_container_width=True,

    hide_index=True
)


# ============================================================
# DOWNLOAD FILTERED DATA
# ============================================================

st.subheader(
    "Download Data"
)


csv_data = filtered_df.to_csv(
    index=False
).encode("utf-8")


st.download_button(

    label="Download Filtered Transactions CSV",

    data=csv_data,

    file_name="mpesa_filtered_transactions.csv",

    mime="text/csv"
)


# ============================================================
# DOWNLOAD SUMMARY
# ============================================================

summary_download = pd.DataFrame({

    "Metric": [

        "Total Money In",

        "Total Money Out",

        "Net Flow",

        "Number of Transactions",

        "Average Transaction",

        "Largest Transaction"
    ],

    "Value": [

        total_money_in,

        total_money_out,

        net_flow,

        transaction_count,

        average_transaction,

        maximum_transaction
    ]
})


st.download_button(

    label="Download Dashboard Summary",

    data=summary_download.to_csv(
        index=False
    ).encode("utf-8"),

    file_name="mpesa_analysis_summary.csv",

    mime="text/csv"
)


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "M-Pesa Statement Analysis Dashboard"
)