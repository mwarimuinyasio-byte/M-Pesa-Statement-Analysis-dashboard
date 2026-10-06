import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="M-Pesa Statement Analysis",
    page_icon="",
    layout="wide"
)


# ============================================================
# TITLE
# ============================================================

st.title("M-Pesa Statement Analysis Dashboard")

st.write(
    "Analyze your M-Pesa transactions, money movements, balances, "
    "transaction types and spending patterns."
)


# ============================================================
# LOAD EXCEL FILE
# ============================================================

@st.cache_data
def load_data():

    file_path = "mpesa_statement.xlsx"

    try:

        df = pd.read_excel(
            file_path,
            sheet_name="Transactions"
        )

    except Exception as e:

        st.error(
            f"Unable to load the Excel file: {e}"
        )

        return pd.DataFrame()

    return df


df = load_data()


# ============================================================
# CHECK DATA
# ============================================================

if df.empty:

    st.warning(
        "No transaction data was found."
    )

    st.stop()


# ============================================================
# CLEAN COLUMN NAMES
# ============================================================

df.columns = (
    df.columns
    .astype(str)
    .str.strip()
)


# ============================================================
# REQUIRED COLUMNS
# ============================================================

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
        "The following required columns are missing:"
    )

    st.write(missing_columns)

    st.stop()


# ============================================================
# CLEAN DATE COLUMN
# ============================================================

df["Completion Time"] = pd.to_datetime(
    df["Completion Time"],
    errors="coerce"
)


# ============================================================
# CLEAN MONEY COLUMNS
# ============================================================

money_columns = [
    "Paid In",
    "Withdrawn",
    "Balance"
]


for column in money_columns:

    df[column] = (
        df[column]
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("KSh", "", regex=False)
        .str.replace("KES", "", regex=False)
        .str.strip()
    )

    df[column] = pd.to_numeric(
        df[column],
        errors="coerce"
    )


# ============================================================
# SORT DATA
# ============================================================

df = df.sort_values(
    "Completion Time"
).reset_index(drop=True)


# ============================================================
# MONEY IN
# ============================================================

df["Money In"] = (
    df["Paid In"]
    .fillna(0)
)


# ============================================================
# MONEY OUT
# ============================================================

df["Money Out"] = (
    df["Withdrawn"]
    .fillna(0)
)


# ============================================================
# TRANSACTION AMOUNT
# ============================================================

df["Transaction Amount"] = (
    df["Money In"] +
    df["Money Out"]
)


# ============================================================
# AMOUNT RANGE
# ============================================================

def amount_range(amount):

    if pd.isna(amount):

        return "Unknown"

    elif amount < 100:

        return "Below KSh 100"

    elif amount < 500:

        return "KSh 100 - 499"

    elif amount < 1000:

        return "KSh 500 - 999"

    elif amount < 5000:

        return "KSh 1,000 - 4,999"

    elif amount < 10000:

        return "KSh 5,000 - 9,999"

    elif amount < 50000:

        return "KSh 10,000 - 49,999"

    else:

        return "KSh 50,000+"


df["Amount Range"] = (
    df["Transaction Amount"]
    .apply(amount_range)
)


# ============================================================
# PERSON / BUSINESS
# ============================================================

df["Person / Business"] = (
    df["Details"]
    .astype(str)
    .str.strip()
)


# ============================================================
# IDENTIFY CHARGES
# ============================================================

df["Is Charge"] = (
    df["Details"]
    .astype(str)
    .str.contains(
        "charge|fee",
        case=False,
        na=False
    )
)


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.header("Dashboard Filters")


# ============================================================
# DATE NAVIGATION
# ============================================================

st.sidebar.subheader("Date Navigation")


valid_dates = (
    df["Completion Time"]
    .dropna()
)


if len(valid_dates) > 0:

    min_date = valid_dates.min().date()

    max_date = valid_dates.max().date()

    preset = st.sidebar.selectbox(
        "Quick Date Range",
        [
            "All Dates",
            "Today",
            "Last 7 Days",
            "Last 30 Days",
            "This Month",
            "Custom"
        ]
    )


    # --------------------------------------------------------
    # ALL DATES
    # --------------------------------------------------------

    if preset == "All Dates":

        start_date = min_date

        end_date = max_date


    # --------------------------------------------------------
    # TODAY
    # --------------------------------------------------------

    elif preset == "Today":

        start_date = max_date

        end_date = max_date


    # --------------------------------------------------------
    # LAST 7 DAYS
    # --------------------------------------------------------

    elif preset == "Last 7 Days":

        start_date = max(
            min_date,
            max_date - pd.Timedelta(days=6)
        )

        end_date = max_date


    # --------------------------------------------------------
    # LAST 30 DAYS
    # --------------------------------------------------------

    elif preset == "Last 30 Days":

        start_date = max(
            min_date,
            max_date - pd.Timedelta(days=29)
        )

        end_date = max_date


    # --------------------------------------------------------
    # THIS MONTH
    # --------------------------------------------------------

    elif preset == "This Month":

        start_date = max(
            min_date,
            max_date.replace(day=1)
        )

        end_date = max_date


    # --------------------------------------------------------
    # CUSTOM DATE
    # --------------------------------------------------------

    else:

        selected_dates = st.sidebar.date_input(
            "Select Date Range",
            value=(
                min_date,
                max_date
            ),
            min_value=min_date,
            max_value=max_date
        )


        if (
            isinstance(selected_dates, tuple)
            and len(selected_dates) == 2
        ):

            start_date = selected_dates[0]

            end_date = selected_dates[1]

        else:

            start_date = min_date

            end_date = max_date


else:

    start_date = None

    end_date = None


# ============================================================
# TRANSACTION TYPE FILTER
# ============================================================

transaction_types = sorted(
    df["Transaction Type"]
    .dropna()
    .astype(str)
    .unique()
)


selected_types = st.sidebar.multiselect(
    "Transaction Type",
    transaction_types,
    default=transaction_types
)


# ============================================================
# AMOUNT RANGE FILTER
# ============================================================

amount_ranges = [
    "Below KSh 100",
    "KSh 100 - 499",
    "KSh 500 - 999",
    "KSh 1,000 - 4,999",
    "KSh 5,000 - 9,999",
    "KSh 10,000 - 49,999",
    "KSh 50,000+",
    "Unknown"
]


available_ranges = [
    x
    for x in amount_ranges
    if x in df["Amount Range"].unique()
]


selected_ranges = st.sidebar.multiselect(
    "Amount Range",
    available_ranges,
    default=available_ranges
)


# ============================================================
# APPLY FILTERS
# ============================================================

filtered_df = df.copy()


# ------------------------------------------------------------
# DATE FILTER
# ------------------------------------------------------------

if (
    start_date is not None
    and end_date is not None
):

    filtered_df = filtered_df[
        (
            filtered_df["Completion Time"].dt.date
            >= start_date
        )
        &
        (
            filtered_df["Completion Time"].dt.date
            <= end_date
        )
    ]


# ------------------------------------------------------------
# TRANSACTION TYPE
# ------------------------------------------------------------

if selected_types:

    filtered_df = filtered_df[
        filtered_df["Transaction Type"].isin(
            selected_types
        )
    ]


# ------------------------------------------------------------
# AMOUNT RANGE
# ------------------------------------------------------------

if selected_ranges:

    filtered_df = filtered_df[
        filtered_df["Amount Range"].isin(
            selected_ranges
        )
    ]


# ============================================================
# SELECTED PERIOD MESSAGE
# ============================================================

if (
    start_date is not None
    and end_date is not None
):

    st.info(
        f"Showing transactions from "
        f"{start_date.strftime('%d %B %Y')} "
        f"to "
        f"{end_date.strftime('%d %B %Y')}"
    )


# ============================================================
# NO RESULTS
# ============================================================

if filtered_df.empty:

    st.warning(
        "No transactions match the selected filters."
    )

    st.stop()


# ============================================================
# SECTION: ACCOUNT BALANCE OVERVIEW
# ============================================================

st.header("Account Balance Overview")


filtered_balance_df = (
    filtered_df[
        filtered_df["Balance"].notna()
    ]
    .copy()
)


if not filtered_balance_df.empty:

    filtered_balance_df = (
        filtered_balance_df
        .sort_values("Completion Time")
    )


    opening_balance = (
        filtered_balance_df.iloc[0]["Balance"]
    )

    closing_balance = (
        filtered_balance_df.iloc[-1]["Balance"]
    )

    highest_balance = (
        filtered_balance_df["Balance"].max()
    )

    lowest_balance = (
        filtered_balance_df["Balance"].min()
    )

else:

    opening_balance = 0

    closing_balance = 0

    highest_balance = 0

    lowest_balance = 0


col1, col2, col3, col4 = st.columns(4)


col1.metric(
    "Opening Balance",
    f"KSh {opening_balance:,.2f}"
)


col2.metric(
    "Closing Balance",
    f"KSh {closing_balance:,.2f}"
)


col3.metric(
    "Highest Balance",
    f"KSh {highest_balance:,.2f}"
)


col4.metric(
    "Lowest Balance",
    f"KSh {lowest_balance:,.2f}"
)


# ============================================================
# SECTION: BALANCE HISTORY
# ============================================================

st.header("Balance History")


balance_df = (
    filtered_df[
        filtered_df["Balance"].notna()
    ]
    .copy()
)


balance_df = (
    balance_df
    .sort_values("Completion Time")
)


if not balance_df.empty:

    fig_balance = px.line(
        balance_df,
        x="Completion Time",
        y="Balance",
        markers=True,
        title="M-Pesa Balance History"
    )

    fig_balance.update_layout(
        template="plotly_white",
        xaxis_title="Date",
        yaxis_title="Balance (KSh)"
    )

    st.plotly_chart(
        fig_balance,
        use_container_width=True
    )


# ============================================================
# BALANCE TABLE
# ============================================================

with st.expander("View Balance History Table"):

    st.dataframe(
        balance_df[
            [
                "Completion Time",
                "Receipt No.",
                "Details",
                "Balance"
            ]
        ],
        use_container_width=True
    )


# ============================================================
# SECTION: MONEY IN AND MONEY OUT
# ============================================================

st.header("Money In and Money Out")


total_money_in = (
    filtered_df["Money In"]
    .sum()
)


total_money_out = (
    filtered_df["Money Out"]
    .sum()
)


total_transactions = len(
    filtered_df
)


net_movement = (
    total_money_in -
    total_money_out
)


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
    "Net Movement",
    f"KSh {net_movement:,.2f}"
)


col4.metric(
    "Transactions",
    f"{total_transactions:,}"
)


# ============================================================
# MONEY IN / OUT CHART
# ============================================================

money_summary = pd.DataFrame(
    {
        "Category": [
            "Money In",
            "Money Out"
        ],
        "Amount": [
            total_money_in,
            total_money_out
        ]
    }
)


fig_money = px.bar(
    money_summary,
    x="Category",
    y="Amount",
    text="Amount",
    title="Money In vs Money Out"
)


fig_money.update_traces(
    texttemplate="KSh %{text:,.2f}",
    textposition="outside"
)


fig_money.update_layout(
    template="plotly_white",
    yaxis_title="Amount (KSh)"
)


st.plotly_chart(
    fig_money,
    use_container_width=True
)


# ============================================================
# SECTION: BALANCE CHANGES
# ============================================================

st.header("Balance Changes")


balance_change_df = (
    filtered_df
    .sort_values("Completion Time")
    .copy()
)


balance_change_df["Balance Change"] = (
    balance_change_df["Balance"]
    .diff()
)


fig_change = px.bar(
    balance_change_df,
    x="Completion Time",
    y="Balance Change",
    title="Balance Change Per Transaction"
)


fig_change.update_layout(
    template="plotly_white",
    xaxis_title="Date",
    yaxis_title="Balance Change (KSh)"
)


st.plotly_chart(
    fig_change,
    use_container_width=True
)


# ============================================================
# SECTION: TRANSACTION TYPE ANALYSIS
# ============================================================

st.header("Transaction Type Analysis")


type_analysis = (
    filtered_df
    .groupby("Transaction Type")
    .agg(
        Transactions=(
            "Receipt No.",
            "count"
        ),
        Total_Amount=(
            "Transaction Amount",
            "sum"
        ),
        Average_Amount=(
            "Transaction Amount",
            "mean"
        )
    )
    .reset_index()
)


col1, col2 = st.columns(2)


with col1:

    fig_type_count = px.bar(
        type_analysis,
        x="Transaction Type",
        y="Transactions",
        title="Number of Transactions by Type"
    )

    fig_type_count.update_layout(
        template="plotly_white"
    )

    st.plotly_chart(
        fig_type_count,
        use_container_width=True
    )


with col2:

    fig_type_amount = px.bar(
        type_analysis,
        x="Transaction Type",
        y="Total_Amount",
        title="Total Amount by Transaction Type"
    )

    fig_type_amount.update_layout(
        template="plotly_white",
        yaxis_title="Amount (KSh)"
    )

    st.plotly_chart(
        fig_type_amount,
        use_container_width=True
    )


st.dataframe(
    type_analysis,
    use_container_width=True
)


# ============================================================
# SECTION: AMOUNT RANGE ANALYSIS
# ============================================================

st.header("Amount Range Analysis")


range_analysis = (
    filtered_df
    .groupby("Amount Range")
    .agg(
        Transactions=(
            "Receipt No.",
            "count"
        ),
        Total_Amount=(
            "Transaction Amount",
            "sum"
        )
    )
    .reset_index()
)


fig_range = px.bar(
    range_analysis,
    x="Amount Range",
    y="Transactions",
    title="Transactions by Amount Range"
)


fig_range.update_layout(
    template="plotly_white",
    xaxis_title="Amount Range",
    yaxis_title="Number of Transactions"
)


st.plotly_chart(
    fig_range,
    use_container_width=True
)


# ============================================================
# SECTION: DAILY MOVEMENT
# ============================================================

st.header("Daily Transaction Movement")


daily_df = (
    filtered_df
    .copy()
)


daily_df["Date"] = (
    daily_df["Completion Time"]
    .dt.date
)


daily_analysis = (
    daily_df
    .groupby("Date")
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
            "Receipt No.",
            "count"
        )
    )
    .reset_index()
)


fig_daily = go.Figure()


fig_daily.add_trace(
    go.Scatter(
        x=daily_analysis["Date"],
        y=daily_analysis["Money_In"],
        mode="lines+markers",
        name="Money In"
    )
)


fig_daily.add_trace(
    go.Scatter(
        x=daily_analysis["Date"],
        y=daily_analysis["Money_Out"],
        mode="lines+markers",
        name="Money Out"
    )
)


fig_daily.update_layout(
    title="Daily Money Movement",
    template="plotly_white",
    xaxis_title="Date",
    yaxis_title="Amount (KSh)"
)


st.plotly_chart(
    fig_daily,
    use_container_width=True
)


# ============================================================
# SECTION: TRANSACTION AMOUNT DISTRIBUTION
# ============================================================

st.header("Transaction Amount Distribution")


fig_hist = px.histogram(
    filtered_df,
    x="Transaction Amount",
    nbins=30,
    title="Transaction Amount Distribution"
)


fig_hist.update_layout(
    template="plotly_white",
    xaxis_title="Transaction Amount (KSh)",
    yaxis_title="Number of Transactions"
)


st.plotly_chart(
    fig_hist,
    use_container_width=True
)


# ============================================================
# SECTION: TRANSACTION STATUS
# ============================================================

st.header("Transaction Status")


status_analysis = (
    filtered_df[
        "Transaction Status"
    ]
    .value_counts()
    .reset_index()
)


status_analysis.columns = [
    "Status",
    "Transactions"
]


fig_status = px.pie(
    status_analysis,
    names="Status",
    values="Transactions",
    title="Transaction Status Distribution"
)


st.plotly_chart(
    fig_status,
    use_container_width=True
)


# ============================================================
# SECTION: TOP SOURCES / DESTINATIONS
# ============================================================

st.header("Top Sources / Destinations")


person_analysis = (
    filtered_df
    .groupby("Person / Business")
    .agg(
        Transactions=(
            "Receipt No.",
            "count"
        ),
        Total_Amount=(
            "Transaction Amount",
            "sum"
        )
    )
    .reset_index()
)


person_analysis = (
    person_analysis
    .sort_values(
        "Total_Amount",
        ascending=False
    )
    .head(15)
)


fig_person = px.bar(
    person_analysis,
    x="Total_Amount",
    y="Person / Business",
    orientation="h",
    title="Top People / Businesses by Transaction Amount"
)


fig_person.update_layout(
    template="plotly_white",
    yaxis_title="Person / Business",
    xaxis_title="Amount (KSh)"
)


st.plotly_chart(
    fig_person,
    use_container_width=True
)


# ============================================================
# SECTION: SAME PERSON / BUSINESS ANALYSIS
# ============================================================

st.header("Repeated People / Businesses")


repeat_analysis = (
    filtered_df
    .groupby("Person / Business")
    .size()
    .reset_index(
        name="Transaction Count"
    )
)


repeat_analysis = (
    repeat_analysis
    .sort_values(
        "Transaction Count",
        ascending=False
    )
    .head(20)
)


st.dataframe(
    repeat_analysis,
    use_container_width=True
)


# ============================================================
# SECTION: CHARGES / FEES
# ============================================================

st.header("M-Pesa Charges and Fees")


charges_df = filtered_df[
    filtered_df["Is Charge"]
].copy()


if not charges_df.empty:

    total_charges = (
        charges_df["Transaction Amount"]
        .sum()
    )

    charge_count = len(
        charges_df
    )

else:

    total_charges = 0

    charge_count = 0


col1, col2 = st.columns(2)


col1.metric(
    "Total Charges",
    f"KSh {total_charges:,.2f}"
)


col2.metric(
    "Number of Charges",
    f"{charge_count:,}"
)


if not charges_df.empty:

    st.dataframe(
        charges_df[
            [
                "Completion Time",
                "Receipt No.",
                "Details",
                "Transaction Amount"
            ]
        ],
        use_container_width=True
    )


# ============================================================
# SECTION: LARGEST TRANSACTIONS
# ============================================================

st.header("Largest Transactions")


largest_transactions = (
    filtered_df
    .sort_values(
        "Transaction Amount",
        ascending=False
    )
    .head(20)
)


st.dataframe(
    largest_transactions[
        [
            "Completion Time",
            "Receipt No.",
            "Details",
            "Transaction Type",
            "Transaction Amount",
            "Balance"
        ]
    ],
    use_container_width=True
)


# ============================================================
# SECTION: FULL TRANSACTION TABLE
# ============================================================

st.header("Full Transaction Table")


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


available_display_columns = [
    column
    for column in display_columns
    if column in filtered_df.columns
]


st.dataframe(
    filtered_df[
        available_display_columns
    ],
    use_container_width=True,
    height=500
)


# ============================================================
# DOWNLOAD FILTERED DATA
# ============================================================

st.header("Download Data")


download_df = filtered_df.copy()


csv_data = download_df.to_csv(
    index=False
).encode("utf-8")


st.download_button(
    label="Download Filtered Transactions CSV",
    data=csv_data,
    file_name="mpesa_filtered_transactions.csv",
    mime="text/csv"
)


# ============================================================
# SUMMARY
# ============================================================

st.header("Statement Summary")


summary_data = {
    "Metric": [
        "Statement Start Date",
        "Statement End Date",
        "Selected Start Date",
        "Selected End Date",
        "Total Transactions",
        "Total Money In",
        "Total Money Out",
        "Net Movement",
        "Opening Balance",
        "Closing Balance",
        "Highest Balance",
        "Lowest Balance",
        "Total Charges"
    ],
    "Value": [
        (
            min_date
            if start_date is not None
            else "N/A"
        ),
        (
            max_date
            if end_date is not None
            else "N/A"
        ),
        (
            start_date
            if start_date is not None
            else "N/A"
        ),
        (
            end_date
            if end_date is not None
            else "N/A"
        ),
        total_transactions,
        f"KSh {total_money_in:,.2f}",
        f"KSh {total_money_out:,.2f}",
        f"KSh {net_movement:,.2f}",
        f"KSh {opening_balance:,.2f}",
        f"KSh {closing_balance:,.2f}",
        f"KSh {highest_balance:,.2f}",
        f"KSh {lowest_balance:,.2f}",
        f"KSh {total_charges:,.2f}"
    ]
}


summary_df = pd.DataFrame(
    summary_data
)


st.dataframe(
    summary_df,
    use_container_width=True
)


# ============================================================
# RAW / OCR DATA
# ============================================================

if "OCR Raw Rows" in pd.ExcelFile(
    "mpesa_statement.xlsx"
).sheet_names:

    st.header("OCR Raw Rows")

    try:

        raw_df = pd.read_excel(
            "mpesa_statement.xlsx",
            sheet_name="OCR Raw Rows"
        )

        st.dataframe(
            raw_df,
            use_container_width=True,
            height=400
        )

    except Exception as e:

        st.warning(
            f"Unable to load OCR Raw Rows: {e}"
        )


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "M-Pesa Statement Analysis Dashboard"
)