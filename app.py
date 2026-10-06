import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import re


# ==========================================================
# PAGE CONFIGURATION
# ==========================================================

st.set_page_config(
    page_title="M-Pesa Statement Analysis",
    page_icon="M",
    layout="wide"
)


# ==========================================================
# CUSTOM CSS
# ==========================================================

st.markdown(
    """
    <style>

    .main-title {
        font-size: 36px;
        font-weight: 700;
        margin-bottom: 5px;
    }

    .subtitle {
        color: #666;
        font-size: 16px;
        margin-bottom: 25px;
    }

    .section-title {
        font-size: 25px;
        font-weight: 600;
        margin-top: 30px;
        margin-bottom: 15px;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# ==========================================================
# TITLE
# ==========================================================

st.markdown(
    '<div class="main-title">M-Pesa Statement Analysis Dashboard</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitle">'
    'Analyze money in, money out, transaction amounts, '
    'transaction ranges, and transactions between the same people.'
    '</div>',
    unsafe_allow_html=True
)


# ==========================================================
# SIDEBAR
# ==========================================================

st.sidebar.header("M-Pesa Statement")

uploaded_file = st.sidebar.file_uploader(
    "Upload M-Pesa Excel File",
    type=["xlsx", "xls"]
)


# ==========================================================
# HELPER FUNCTIONS
# ==========================================================

def clean_money(value):

    if pd.isna(value):
        return 0.0

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

    except:
        return 0.0


def amount_range(amount):

    if pd.isna(amount):
        return "Unknown"

    amount = abs(float(amount))

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


def detect_transaction_type(details):

    text = str(details).lower()

    # MONEY IN

    money_in_words = [
        "funds received",
        "received from",
        "money received",
        "receive international",
        "received international"
    ]

    for word in money_in_words:

        if word in text:
            return "Money In"


    # MONEY OUT

    money_out_words = [
        "customer transfer to",
        "merchant payment",
        "pay bill",
        "airtime purchase",
        "airtime purchase",
        "withdrawal",
        "send money",
        "customer payment"
    ]

    for word in money_out_words:

        if word in text:
            return "Money Out"


    # CHARGES

    charge_words = [
        "charge",
        "fee"
    ]

    for word in charge_words:

        if word in text:
            return "Charge"


    return "Other"


def extract_person(details):

    """
    Attempts to identify the person or business
    involved in the transaction.
    """

    text = str(details).strip()

    patterns = [

        r"Customer Transfer to\s*-\s*(?:\d+\*+\d+\s*)?([A-Za-z][A-Za-z .'-]+)",

        r"Funds received from\s*-\s*(?:\d+\*+\d+\s*)?([A-Za-z][A-Za-z .'-]+)",

        r"Customer Transfer to\s*(?:-\s*)?(?:\d+\*+\d+\s*)?([A-Za-z][A-Za-z .'-]+)",

        r"Merchant Payment to\s*(?:-\s*)?(?:\d+\*+\d+\s*)?([A-Za-z][A-Za-z .'-]+)",

        r"Pay Bill.*?-\s*([A-Za-z][A-Za-z .'-]+)"
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            text,
            re.IGNORECASE
        )

        if match:

            person = match.group(1).strip()

            person = re.sub(
                r"\s+",
                " ",
                person
            )

            return person

    return "Unknown"


# ==========================================================
# CHECK FILE
# ==========================================================

if uploaded_file is None:

    st.info(
        "Upload the M-Pesa Excel workbook to start the analysis."
    )

    st.markdown(
        """
        ### Required workbook structure

        The Excel workbook should contain:

        **Transactions**
        - Receipt No.
        - Completion Time
        - Details
        - Transaction Status
        - Paid In
        - Withdrawn
        - Balance
        - Transaction Type

        **Statement Summary**
        - Transaction Type
        - Paid In
        - Paid Out

        **OCR Raw Rows**
        - Page
        - OCR_Row
        """
    )

    st.stop()


# ==========================================================
# LOAD EXCEL FILE
# ==========================================================

try:

    excel = pd.ExcelFile(
        uploaded_file
    )

except Exception as e:

    st.error(
        f"Could not read the Excel file: {e}"
    )

    st.stop()


# ==========================================================
# CHECK REQUIRED SHEETS
# ==========================================================

required_sheets = [
    "Transactions",
    "Statement Summary",
    "OCR Raw Rows"
]

missing_sheets = [
    sheet
    for sheet in required_sheets
    if sheet not in excel.sheet_names
]


if missing_sheets:

    st.error(
        "The following sheets are missing: "
        + ", ".join(missing_sheets)
    )

    st.stop()


# ==========================================================
# LOAD SHEETS
# ==========================================================

df = pd.read_excel(
    uploaded_file,
    sheet_name="Transactions"
)

summary_df = pd.read_excel(
    uploaded_file,
    sheet_name="Statement Summary"
)

ocr_df = pd.read_excel(
    uploaded_file,
    sheet_name="OCR Raw Rows"
)


# ==========================================================
# CHECK TRANSACTION COLUMNS
# ==========================================================

required_columns = [

    "Receipt No.",
    "Completion Time",
    "Details",
    "Transaction Status",
    "Paid In",
    "Withdrawn",
    "Balance",
    "Transaction Type"
]


missing_columns = [

    column
    for column in required_columns
    if column not in df.columns
]


if missing_columns:

    st.error(
        "The following columns are missing from "
        "the Transactions sheet:"
    )

    st.write(
        missing_columns
    )

    st.stop()


# ==========================================================
# CLEAN DATA
# ==========================================================

df["Completion Time"] = pd.to_datetime(
    df["Completion Time"],
    errors="coerce"
)


df["Paid In"] = df[
    "Paid In"
].apply(clean_money)


df["Withdrawn"] = df[
    "Withdrawn"
].apply(clean_money)


df["Balance"] = df[
    "Balance"
].apply(clean_money)


df["Details"] = (
    df["Details"]
    .fillna("")
    .astype(str)
)


# ==========================================================
# CLEAN TRANSACTION TYPE
# ==========================================================

df["Transaction Type"] = (

    df["Transaction Type"]
    .fillna("")
    .astype(str)
)


detected_type = df[
    "Details"
].apply(
    detect_transaction_type
)


for i in df.index:

    current = (
        df.loc[
            i,
            "Transaction Type"
        ]
        .strip()
        .lower()
    )

    if current in [
        "",
        "nan",
        "unknown",
        "other"
    ]:

        df.loc[
            i,
            "Transaction Type"
        ] = detected_type.loc[i]


# ==========================================================
# CREATE MONEY IN COLUMN
# ==========================================================

df["Money In"] = df[
    "Paid In"
].clip(
    lower=0
)


# ==========================================================
# CREATE MONEY OUT COLUMN
# ==========================================================

df["Money Out"] = df[
    "Withdrawn"
].abs()


# ==========================================================
# CREATE TRANSACTION AMOUNT
# ==========================================================

df["Transaction Amount"] = np.where(

    df["Money In"] > 0,

    df["Money In"],

    df["Money Out"]
)


# ==========================================================
# AMOUNT RANGE
# ==========================================================

df["Amount Range"] = df[
    "Transaction Amount"
].apply(
    amount_range
)


# ==========================================================
# PERSON / BUSINESS
# ==========================================================

df["Person / Business"] = df[
    "Details"
].apply(
    extract_person
)


# ==========================================================
# CHARGE IDENTIFICATION
# ==========================================================

df["Is Charge"] = (

    df["Details"]
    .str.contains(
        "charge|fee",
        case=False,
        na=False
    )
)


# ==========================================================
# SIDEBAR FILTERS
# ==========================================================

st.sidebar.subheader(
    "Filters"
)


transaction_types = sorted(
    df["Transaction Type"]
    .dropna()
    .unique()
    .tolist()
)


selected_types = st.sidebar.multiselect(

    "Transaction Type",

    transaction_types,

    default=transaction_types
)


amount_ranges = [

    "Below 500",

    "500 - 999",

    "1,000 - 1,999",

    "2,000 - 4,999",

    "5,000 - 9,999",

    "10,000+"
]


selected_ranges = st.sidebar.multiselect(

    "Amount Range",

    amount_ranges,

    default=amount_ranges
)


# ==========================================================
# DATE FILTER
# ==========================================================

valid_dates = df[
    "Completion Time"
].dropna()


if not valid_dates.empty:

    min_date = valid_dates.min().date()

    max_date = valid_dates.max().date()

    date_range = st.sidebar.date_input(

        "Transaction Date",

        value=(
            min_date,
            max_date
        ),

        min_value=min_date,

        max_value=max_date
    )

else:

    date_range = None


# ==========================================================
# APPLY FILTERS
# ==========================================================

filtered_df = df[
    df["Transaction Type"].isin(
        selected_types
    )
    &
    df["Amount Range"].isin(
        selected_ranges
    )
].copy()


if date_range and len(date_range) == 2:

    start_date = pd.Timestamp(
        date_range[0]
    )

    end_date = (
        pd.Timestamp(
            date_range[1]
        )
        + pd.Timedelta(days=1)
    )


    filtered_df = filtered_df[
        (
            filtered_df[
                "Completion Time"
            ] >= start_date
        )
        &
        (
            filtered_df[
                "Completion Time"
            ] < end_date
        )
    ]


# ==========================================================
# MAIN KPIs
# ==========================================================

st.markdown(
    '<div class="section-title">Financial Summary</div>',
    unsafe_allow_html=True
)


total_money_in = filtered_df[
    "Money In"
].sum()


total_money_out = filtered_df[
    "Money Out"
].sum()


net_flow = (
    total_money_in
    -
    total_money_out
)


number_transactions = len(
    filtered_df
)


average_transaction = (

    filtered_df[
        "Transaction Amount"
    ].mean()

    if not filtered_df.empty

    else 0
)


largest_transaction = (

    filtered_df[
        "Transaction Amount"
    ].max()

    if not filtered_df.empty

    else 0
)


# ==========================================================
# KPI CARDS
# ==========================================================

col1, col2, col3, col4 = st.columns(4)


with col1:

    st.metric(
        "Total Money In",
        f"KSh {total_money_in:,.2f}"
    )


with col2:

    st.metric(
        "Total Money Out",
        f"KSh {total_money_out:,.2f}"
    )


with col3:

    st.metric(
        "Net Cash Flow",
        f"KSh {net_flow:,.2f}"
    )


with col4:

    st.metric(
        "Transactions",
        f"{number_transactions:,}"
    )


col5, col6 = st.columns(2)


with col5:

    st.metric(
        "Average Transaction",
        f"KSh {average_transaction:,.2f}"
    )


with col6:

    st.metric(
        "Largest Transaction",
        f"KSh {largest_transaction:,.2f}"
    )


# ==========================================================
# MONEY IN VS MONEY OUT
# ==========================================================

st.markdown(
    '<div class="section-title">Money In vs Money Out</div>',
    unsafe_allow_html=True
)


flow_df = pd.DataFrame({

    "Type": [
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

    x="Type",

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


# ==========================================================
# MONEY IN PIE CHART
# ==========================================================

st.markdown(
    '<div class="section-title">Money Movement</div>',
    unsafe_allow_html=True
)


col1, col2 = st.columns(2)


with col1:

    fig_pie = px.pie(

        flow_df,

        names="Type",

        values="Amount",

        title="Money In vs Money Out",

        hole=0.4
    )

    st.plotly_chart(
        fig_pie,
        use_container_width=True
    )


with col2:

    type_summary = (

        filtered_df

        .groupby(
            "Transaction Type",
            as_index=False
        )

        .agg(
            Amount=(
                "Transaction Amount",
                "sum"
            ),

            Transactions=(
                "Transaction Amount",
                "count"
            )
        )
    )


    fig_type = px.bar(

        type_summary,

        x="Transaction Type",

        y="Amount",

        text_auto=".2f",

        title="Amount by Transaction Type",

        template="plotly_white"
    )


    st.plotly_chart(
        fig_type,
        use_container_width=True
    )


# ==========================================================
# AMOUNT RANGE ANALYSIS
# ==========================================================

st.markdown(
    '<div class="section-title">Transaction Amount Ranges</div>',
    unsafe_allow_html=True
)


range_summary = (

    filtered_df

    .groupby(
        "Amount Range",
        as_index=False
    )

    .agg(

        Money_In=(
            "Money In",
            "sum"
        ),

        Money_Out=(
            "Money Out",
            "sum"
        ),

        Transactions=(
            "Transaction Amount",
            "count"
        )
    )
)


range_summary["Amount Range"] = pd.Categorical(

    range_summary["Amount Range"],

    categories=amount_ranges,

    ordered=True
)


range_summary = (
    range_summary
    .sort_values(
        "Amount Range"
    )
)


st.dataframe(
    range_summary,
    use_container_width=True,
    hide_index=True
)


fig_ranges = px.bar(

    range_summary,

    x="Amount Range",

    y=[
        "Money_In",
        "Money_Out"
    ],

    barmode="group",

    text_auto=".2f",

    title="Money In and Money Out by Amount Range",

    labels={
        "value": "Amount (KSh)",
        "variable": "Transaction Flow"
    },

    template="plotly_white"
)


st.plotly_chart(
    fig_ranges,
    use_container_width=True
)


# ==========================================================
# PEOPLE WHO SENT MONEY
# ==========================================================

st.markdown(
    '<div class="section-title">Money In — People / Sources</div>',
    unsafe_allow_html=True
)


money_in_df = filtered_df[
    (
        filtered_df[
            "Money In"
        ] > 0
    )
    &
    (
        ~filtered_df[
            "Is Charge"
        ]
    )
].copy()


received_people = (

    money_in_df

    .groupby(
        "Person / Business",
        as_index=False
    )

    .agg(

        Total_Received=(
            "Money In",
            "sum"
        ),

        Transactions=(
            "Money In",
            "count"
        )
    )

    .sort_values(
        "Total_Received",
        ascending=False
    )
)


if received_people.empty:

    st.info(
        "No money-in sources found."
    )

else:

    st.dataframe(
        received_people.head(30),
        use_container_width=True,
        hide_index=True
    )


    fig_received = px.bar(

        received_people.head(15)
        .sort_values(
            "Total_Received"
        ),

        x="Total_Received",

        y="Person / Business",

        orientation="h",

        text_auto=".2f",

        title="Top Sources of Money Received",

        labels={
            "Total_Received":
                "Money Received (KSh)"
        },

        template="plotly_white"
    )


    st.plotly_chart(
        fig_received,
        use_container_width=True
    )


# ==========================================================
# PEOPLE WHO RECEIVED MONEY
# ==========================================================

st.markdown(
    '<div class="section-title">Money Out — People / Businesses</div>',
    unsafe_allow_html=True
)


money_out_df = filtered_df[
    (
        filtered_df[
            "Money Out"
        ] > 0
    )
    &
    (
        ~filtered_df[
            "Is Charge"
        ]
    )
].copy()


sent_people = (

    money_out_df

    .groupby(
        "Person / Business",
        as_index=False
    )

    .agg(

        Total_Sent=(
            "Money Out",
            "sum"
        ),

        Transactions=(
            "Money Out",
            "count"
        )
    )

    .sort_values(
        "Total_Sent",
        ascending=False
    )
)


if sent_people.empty:

    st.info(
        "No money-out destinations found."
    )

else:

    st.dataframe(
        sent_people.head(30),
        use_container_width=True,
        hide_index=True
    )


    fig_sent = px.bar(

        sent_people.head(15)
        .sort_values(
            "Total_Sent"
        ),

        x="Total_Sent",

        y="Person / Business",

        orientation="h",

        text_auto=".2f",

        title="Top People / Businesses You Paid",

        labels={
            "Total_Sent":
                "Money Sent (KSh)"
        },

        template="plotly_white"
    )


    st.plotly_chart(
        fig_sent,
        use_container_width=True
    )


# ==========================================================
# SAME PERSON ANALYSIS
# ==========================================================

st.markdown(
    '<div class="section-title">'
    'Transactions To and From the Same Person'
    '</div>',
    unsafe_allow_html=True
)


received = (

    money_in_df

    .groupby(
        "Person / Business",
        as_index=False
    )

    .agg(
        Money_Received=(
            "Money In",
            "sum"
        ),

        Received_Transactions=(
            "Money In",
            "count"
        )
    )
)


sent = (

    money_out_df

    .groupby(
        "Person / Business",
        as_index=False
    )

    .agg(
        Money_Sent=(
            "Money Out",
            "sum"
        ),

        Sent_Transactions=(
            "Money Out",
            "count"
        )
    )
)


same_people = pd.merge(

    received,

    sent,

    on="Person / Business",

    how="inner"
)


if same_people.empty:

    st.info(
        "No person appears in both money-in and money-out transactions."
    )

else:

    same_people["Net Difference"] = (

        same_people[
            "Money_Received"
        ]

        -

        same_people[
            "Money_Sent"
        ]
    )


    same_people["Total Activity"] = (

        same_people[
            "Money_Received"
        ]

        +

        same_people[
            "Money_Sent"
        ]
    )


    same_people = same_people.sort_values(

        "Total Activity",

        ascending=False
    )


    st.dataframe(

        same_people,

        use_container_width=True,

        hide_index=True
    )


    fig_same = px.bar(

        same_people.head(15),

        x="Person / Business",

        y=[
            "Money_Received",
            "Money_Sent"
        ],

        barmode="group",

        title="Money Received vs Money Sent by the Same Person",

        labels={
            "value": "Amount (KSh)",
            "Person / Business": "Person / Business",
            "variable": "Transaction Direction"
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


# ==========================================================
# DAILY TRANSACTION GRAPH
# ==========================================================

st.markdown(
    '<div class="section-title">Daily Money Movement</div>',
    unsafe_allow_html=True
)


daily_df = filtered_df.dropna(
    subset=["Completion Time"]
).copy()


if not daily_df.empty:

    daily_df["Date"] = (

        daily_df[
            "Completion Time"
        ].dt.date
    )


    daily_summary = (

        daily_df

        .groupby(
            "Date",
            as_index=False
        )

        .agg(

            Money_In=(
                "Money In",
                "sum"
            ),

            Money_Out=(
                "Money Out",
                "sum"
            )
        )
    )


    fig_daily = px.line(

        daily_summary,

        x="Date",

        y=[
            "Money_In",
            "Money_Out"
        ],

        markers=True,

        title="Daily Money In and Money Out",

        labels={
            "value": "Amount (KSh)",
            "variable": "Transaction Flow"
        },

        template="plotly_white"
    )


    st.plotly_chart(

        fig_daily,

        use_container_width=True
    )


# ==========================================================
# TRANSACTION COUNT BY DAY
# ==========================================================

st.markdown(
    '<div class="section-title">Number of Transactions by Day</div>',
    unsafe_allow_html=True
)


if not daily_df.empty:

    count_daily = (

        daily_df

        .groupby(
            "Date",
            as_index=False
        )

        .size()

        .rename(
            columns={
                "size":
                    "Transactions"
            }
        )
    )


    fig_count = px.bar(

        count_daily,

        x="Date",

        y="Transactions",

        title="Daily Transaction Count",

        text_auto=True,

        template="plotly_white"
    )


    st.plotly_chart(

        fig_count,

        use_container_width=True
    )


# ==========================================================
# TRANSACTION AMOUNT DISTRIBUTION
# ==========================================================

st.markdown(
    '<div class="section-title">Transaction Amount Distribution</div>',
    unsafe_allow_html=True
)


if not filtered_df.empty:

    fig_hist = px.histogram(

        filtered_df,

        x="Transaction Amount",

        color="Transaction Type",

        nbins=30,

        title="Distribution of Transaction Amounts",

        labels={
            "Transaction Amount":
                "Amount (KSh)"
        },

        template="plotly_white"
    )


    st.plotly_chart(

        fig_hist,

        use_container_width=True
    )


# ==========================================================
# BALANCE TREND
# ==========================================================

st.markdown(
    '<div class="section-title">M-Pesa Balance Trend</div>',
    unsafe_allow_html=True
)


balance_df = filtered_df.dropna(
    subset=[
        "Completion Time",
        "Balance"
    ]
).sort_values(
    "Completion Time"
)


if not balance_df.empty:

    fig_balance = px.line(

        balance_df,

        x="Completion Time",

        y="Balance",

        markers=True,

        title="Balance After Transactions",

        labels={
            "Completion Time":
                "Transaction Time",
            "Balance":
                "Balance (KSh)"
        },

        template="plotly_white"
    )


    st.plotly_chart(

        fig_balance,

        use_container_width=True
    )


# ==========================================================
# TRANSACTION STATUS
# ==========================================================

st.markdown(
    '<div class="section-title">Transaction Status</div>',
    unsafe_allow_html=True
)


status_summary = (

    filtered_df

    .groupby(
        "Transaction Status",
        as_index=False
    )

    .size()

    .rename(
        columns={
            "size":
                "Transactions"
        }
    )
)


fig_status = px.pie(

    status_summary,

    names="Transaction Status",

    values="Transactions",

    title="Transaction Status Distribution",

    hole=0.4
)


st.plotly_chart(

    fig_status,

    use_container_width=True
)


# ==========================================================
# ORIGINAL STATEMENT SUMMARY
# ==========================================================

st.markdown(
    '<div class="section-title">'
    'Original Statement Summary'
    '</div>',
    unsafe_allow_html=True
)


if not summary_df.empty:

    summary_display = summary_df.copy()


    for column in [
        "Paid In",
        "Paid Out"
    ]:

        if column in summary_display.columns:

            summary_display[column] = (
                summary_display[column]
                .apply(clean_money)
            )


    st.dataframe(

        summary_display,

        use_container_width=True,

        hide_index=True
    )


    if (
        "Transaction Type"
        in summary_display.columns
        and
        "Paid In"
        in summary_display.columns
        and
        "Paid Out"
        in summary_display.columns
    ):

        summary_long = summary_display.melt(

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

            title="Statement Summary",

            text_auto=".2f",

            template="plotly_white"
        )


        st.plotly_chart(

            fig_summary,

            use_container_width=True
        )


# ==========================================================
# OCR RAW DATA
# ==========================================================

with st.expander(
    "View OCR Raw Rows"
):

    st.dataframe(

        ocr_df,

        use_container_width=True,

        hide_index=True
    )


# ==========================================================
# FULL TRANSACTION DATA
# ==========================================================

st.markdown(
    '<div class="section-title">'
    'Detailed Transaction Data'
    '</div>',
    unsafe_allow_html=True
)


display_columns = [

    "Receipt No.",

    "Completion Time",

    "Details",

    "Transaction Status",

    "Paid In",

    "Withdrawn",

    "Balance",

    "Transaction Type",

    "Transaction Amount",

    "Amount Range",

    "Person / Business"
]


st.dataframe(

    filtered_df[
        display_columns
    ].sort_values(
        "Completion Time",
        ascending=False
    ),

    use_container_width=True,

    hide_index=True
)


# ==========================================================
# DOWNLOAD FILTERED DATA
# ==========================================================

st.markdown(
    '<div class="section-title">'
    'Download Results'
    '</div>',
    unsafe_allow_html=True
)


download_df = filtered_df[
    display_columns
].copy()


csv = download_df.to_csv(
    index=False
).encode(
    "utf-8"
)


st.download_button(

    label="Download Filtered Transactions",

    data=csv,

    file_name="mpesa_filtered_transactions.csv",

    mime="text/csv"
)


# ==========================================================
# DOWNLOAD SUMMARY
# ==========================================================

dashboard_summary = pd.DataFrame({

    "Metric": [

        "Total Money In",

        "Total Money Out",

        "Net Cash Flow",

        "Number of Transactions",

        "Average Transaction",

        "Largest Transaction"
    ],

    "Value": [

        total_money_in,

        total_money_out,

        net_flow,

        number_transactions,

        average_transaction,

        largest_transaction
    ]
})


summary_csv = dashboard_summary.to_csv(
    index=False
).encode(
    "utf-8"
)


st.download_button(

    label="Download Financial Summary",

    data=summary_csv,

    file_name="mpesa_financial_summary.csv",

    mime="text/csv"
)


# ==========================================================
# FOOTER
# ==========================================================

st.divider()

st.caption(
    "M-Pesa Statement Analysis Dashboard"
)