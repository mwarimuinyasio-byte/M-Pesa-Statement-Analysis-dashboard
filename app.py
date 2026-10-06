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
    "Analyze your M-Pesa transactions, money received, money sent, "
    "amount ranges, and transactions involving the same person."
)


# ============================================================
# FUNCTIONS
# ============================================================

def clean_columns(df):
    """
    Standardize column names.
    """

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


def find_column(df, possible_names):
    """
    Find a column even when its name has slightly different formatting.
    """

    normalized_columns = {
        str(column).lower().replace(" ", "_"): column
        for column in df.columns
    }

    for name in possible_names:

        key = name.lower().replace(" ", "_")

        if key in normalized_columns:
            return normalized_columns[key]

    return None


def convert_money(series):
    """
    Convert money values such as:
    KSh 1,000
    1,000
    -500
    into numbers.
    """

    return pd.to_numeric(
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("KSh", "", regex=False)
        .str.replace("KES", "", regex=False)
        .str.replace(r"[^\d.\-]", "", regex=True),
        errors="coerce"
    )


def classify_transaction(row):
    """
    Determine whether the transaction is Money In,
    Money Out, or Other.
    """

    details = str(row.get("details", "")).lower()

    # Money received
    money_in_words = [
        "funds received",
        "received from",
        "money received",
        "receive international",
        "received international",
        "cash deposit"
    ]

    for word in money_in_words:

        if word in details:
            return "Money In"

    # Money sent/spent
    money_out_words = [
        "customer transfer to",
        "merchant payment",
        "pay bill",
        "airtime",
        "withdrawal",
        "customer withdrawal",
        "send money",
        "funds transfer to"
    ]

    for word in money_out_words:

        if word in details:
            return "Money Out"

    # Use Paid In / Withdrawn columns as fallback

    paid_in = row.get("paid_in", 0)

    withdrawn = row.get("withdrawn", 0)

    if pd.notna(paid_in) and paid_in > 0:
        return "Money In"

    if pd.notna(withdrawn) and withdrawn < 0:
        return "Money Out"

    return "Other"


def extract_person(details):
    """
    Try to extract the person/business name
    from the M-Pesa Details column.
    """

    text = str(details).strip()

    patterns = [

        r"Customer Transfer to\s*-\s*.*?\s+(.+)",

        r"Funds received from\s*-\s*.*?\s+(.+)",

        r"Merchant Payment to\s*.*?-\s*(.+)",

        r"Customer Payment to Small Business to\s*-\s*.*?\s+(.+)",

        r"Pay Bill\s*.*?-\s*(.+)"
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            text,
            flags=re.IGNORECASE
        )

        if match:

            person = match.group(1).strip(" -")

            if person:
                return person

    return text


def amount_range(amount):

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
# FILE UPLOAD
# ============================================================

st.sidebar.header("Upload Statement")

uploaded_file = st.sidebar.file_uploader(
    "Upload M-Pesa CSV or Excel file",
    type=["csv", "xlsx", "xls"]
)


# ============================================================
# INSTRUCTIONS BEFORE UPLOAD
# ============================================================

if uploaded_file is None:

    st.info(
        "Upload your M-Pesa transaction CSV or Excel file "
        "using the sidebar."
    )

    st.subheader("Expected columns")

    example = pd.DataFrame({

        "Receipt No.": [
            "ABC001",
            "ABC002",
            "ABC003"
        ],

        "Completion Time": [
            "2026-10-01 10:00:00",
            "2026-10-01 12:00:00",
            "2026-10-02 09:00:00"
        ],

        "Details": [

            "Funds received from - 2547****123 JOHN DOE",

            "Customer Transfer to - 2547****456 JANE DOE",

            "Merchant Payment to 123456 - SHOP"
        ],

        "Transaction Status": [
            "Completed",
            "Completed",
            "Completed"
        ],

        "Paid In": [
            5000,
            0,
            0
        ],

        "Withdrawn": [
            0,
            -700,
            -250
        ],

        "Balance": [
            5000,
            4300,
            4050
        ]
    })

    st.dataframe(
        example,
        use_container_width=True,
        hide_index=True
    )

    st.stop()


# ============================================================
# READ FILE
# ============================================================

try:

    if uploaded_file.name.lower().endswith(".csv"):

        df = pd.read_csv(uploaded_file)

    else:

        df = pd.read_excel(uploaded_file)

except Exception as error:

    st.error(
        f"Unable to read the file: {error}"
    )

    st.stop()


# ============================================================
# CLEAN COLUMN NAMES
# ============================================================

df = clean_columns(df)


# ============================================================
# FIND IMPORTANT COLUMNS
# ============================================================

details_column = find_column(
    df,
    [
        "details",
        "description",
        "transaction_details"
    ]
)

paid_in_column = find_column(
    df,
    [
        "paid_in",
        "paidin",
        "money_in"
    ]
)

withdrawn_column = find_column(
    df,
    [
        "withdrawn",
        "withdrawal",
        "paid_out",
        "money_out"
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

balance_column = find_column(
    df,
    [
        "balance"
    ]
)


# ============================================================
# CHECK DETAILS COLUMN
# ============================================================

if details_column is None:

    st.error(
        "The Details column could not be found."
    )

    st.write(
        "Columns detected in your file:"
    )

    st.write(
        list(df.columns)
    )

    st.stop()


# ============================================================
# STANDARDIZE DATA
# ============================================================

df["details"] = df[details_column].astype(str)


# Paid In

if paid_in_column:

    df["paid_in"] = convert_money(
        df[paid_in_column]
    )

else:

    df["paid_in"] = 0.0


# Withdrawn

if withdrawn_column:

    df["withdrawn"] = convert_money(
        df[withdrawn_column]
    )

else:

    df["withdrawn"] = 0.0


# Balance

if balance_column:

    df["balance"] = convert_money(
        df[balance_column]
    )

else:

    df["balance"] = np.nan


# Completion Time

if date_column:

    df["completion_time"] = pd.to_datetime(
        df[date_column],
        errors="coerce"
    )

else:

    df["completion_time"] = pd.NaT


# ============================================================
# REMOVE EMPTY TRANSACTIONS
# ============================================================

df = df[
    df["details"].str.strip().ne("")
].copy()


# ============================================================
# CREATE MONEY IN / MONEY OUT
# ============================================================

df["money_in"] = (
    df["paid_in"]
    .fillna(0)
    .clip(lower=0)
)

df["money_out"] = (
    df["withdrawn"]
    .fillna(0)
    .abs()
)


# ============================================================
# TRANSACTION TYPE
# ============================================================

df["transaction_type"] = df.apply(
    classify_transaction,
    axis=1
)


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

df["amount_range"] = df[
    "amount"
].apply(amount_range)


# ============================================================
# PERSON / BUSINESS
# ============================================================

df["person"] = df[
    "details"
].apply(extract_person)


# ============================================================
# IDENTIFY CHARGES
# ============================================================

df["is_charge"] = df[
    "details"
].str.contains(
    "charge|fee",
    case=False,
    na=False
)


# ============================================================
# SIDEBAR FILTERS
# ============================================================

st.sidebar.header("Dashboard Filters")


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


# ============================================================
# DASHBOARD KPIs
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


number_transactions = len(
    filtered_df
)


st.subheader("Financial Summary")


col1, col2, col3, col4 = st.columns(4)


col1.metric(
    "Total Money In",
    f"KSh {total_money_in:,.2f}"
)


col2.metric(
    "Total Money Out",
    f"KSh {total_money_out:,.2f}"
)


col3.metric(
    "Net Flow",
    f"KSh {net_flow:,.2f}"
)


col4.metric(
    "Transactions",
    f"{number_transactions:,}"
)


# ============================================================
# MONEY IN VS MONEY OUT
# ============================================================

st.subheader(
    "Money In vs Money Out"
)


flow_data = pd.DataFrame({

    "Transaction Type": [
        "Money In",
        "Money Out"
    ],

    "Amount": [
        total_money_in,
        total_money_out
    ]
})


fig_flow = px.bar(

    flow_data,

    x="Transaction Type",

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
# AMOUNT RANGE ANALYSIS
# ============================================================

st.subheader(
    "Money In and Money Out by Amount Range"
)


range_summary = (

    filtered_df

    .groupby(
        "amount_range",
        as_index=False
    )

    .agg(

        money_in=(
            "money_in",
            "sum"
        ),

        money_out=(
            "money_out",
            "sum"
        ),

        transactions=(
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


range_summary = range_summary.sort_values(
    "amount_range"
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
        "money_in",
        "money_out"
    ],

    barmode="group",

    title="Money In vs Money Out by Amount Range",

    labels={
        "amount_range": "Amount Range",
        "value": "Amount (KSh)",
        "variable": "Transaction Flow"
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
    "Money In — Sources Ranked by Amount"
)


incoming = filtered_df[

    (filtered_df["transaction_type"] == "Money In")

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

        total_received=(
            "money_in",
            "sum"
        ),

        transactions=(
            "money_in",
            "count"
        )
    )

    .sort_values(
        "total_received",
        ascending=False
    )
)


if incoming_people.empty:

    st.warning(
        "No money-in transactions found."
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
            "total_received"
        ),

        x="total_received",

        y="person",

        orientation="h",

        text_auto=".2f",

        title="Top Sources of Money Received",

        labels={
            "total_received": "Money Received (KSh)",
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
    "Money Out — Destinations Ranked by Amount"
)


outgoing = filtered_df[

    (filtered_df["transaction_type"] == "Money Out")

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

        total_sent=(
            "money_out",
            "sum"
        ),

        transactions=(
            "money_out",
            "count"
        )
    )

    .sort_values(
        "total_sent",
        ascending=False
    )
)


if outgoing_people.empty:

    st.warning(
        "No money-out transactions found."
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
            "total_sent"
        ),

        x="total_sent",

        y="person",

        orientation="h",

        text_auto=".2f",

        title="Top Money-Out Destinations",

        labels={
            "total_sent": "Money Sent (KSh)",
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
    "Transactions To and From the Same Person"
)


received_people = (

    incoming

    .groupby(
        "person",
        as_index=False
    )["money_in"]

    .sum()

    .rename(
        columns={
            "money_in": "received"
        }
    )
)


sent_people = (

    outgoing

    .groupby(
        "person",
        as_index=False
    )["money_out"]

    .sum()

    .rename(
        columns={
            "money_out": "sent"
        }
    )
)


same_person = pd.merge(

    received_people,

    sent_people,

    on="person",

    how="inner"
)


if same_person.empty:

    st.warning(
        "No person appears on both the money-in and money-out sides."
    )

else:

    same_person[
        "total_activity"
    ] = (

        same_person["received"]

        +

        same_person["sent"]
    )


    same_person[
        "net_from_person"
    ] = (

        same_person["received"]

        -

        same_person["sent"]
    )


    same_person = same_person.sort_values(

        "total_activity",

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
            "received",
            "sent"
        ],

        barmode="group",

        title=(
            "People Who Both Sent Money To You "
            "and Received Money From You"
        ),

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
# DAILY MONEY FLOW
# ============================================================

st.subheader(
    "Daily Money In and Money Out"
)


timeline = filtered_df.dropna(
    subset=["completion_time"]
).copy()


if timeline.empty:

    st.warning(
        "No valid transaction dates were found."
    )

else:

    timeline["date"] = (
        timeline["completion_time"]
        .dt.date
    )


    daily = (

        timeline

        .groupby(
            "date",
            as_index=False
        )

        .agg(

            money_in=(
                "money_in",
                "sum"
            ),

            money_out=(
                "money_out",
                "sum"
            )
        )
    )


    fig_daily = px.line(

        daily,

        x="date",

        y=[
            "money_in",
            "money_out"
        ],

        markers=True,

        title="Daily Money In and Money Out",

        labels={
            "date": "Date",
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
            "amount": "Transaction Amount (KSh)"
        },

        template="plotly_white"
    )


    st.plotly_chart(

        fig_histogram,

        use_container_width=True
    )


# ============================================================
# TRANSACTION TYPE COUNT
# ============================================================

st.subheader(
    "Number of Transactions by Type"
)


type_count = (

    filtered_df

    .groupby(
        "transaction_type",
        as_index=False
    )

    .size()

    .rename(
        columns={
            "size": "transactions"
        }
    )
)


fig_count = px.pie(

    type_count,

    names="transaction_type",

    values="transactions",

    title="Transaction Count by Type",

    hole=0.35
)


st.plotly_chart(

    fig_count,

    use_container_width=True
)


# ============================================================
# FILTERED TRANSACTION TABLE
# ============================================================

st.subheader(
    "Filtered M-Pesa Transactions"
)


display_columns = [

    "completion_time",

    "details",

    "transaction_type",

    "amount",

    "amount_range",

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
    "Download Analysis"
)


csv_data = filtered_df.to_csv(
    index=False
).encode("utf-8")


st.download_button(

    label="Download Filtered Transactions",

    data=csv_data,

    file_name="mpesa_filtered_transactions.csv",

    mime="text/csv"
)


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "M-Pesa Statement Analysis Dashboard"
)