import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO


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

st.markdown(
    """
    Upload your M-Pesa Excel workbook and analyze:

    - Money received
    - Money sent
    - Withdrawals
    - Merchant payments
    - Airtime purchases
    - Charges and fees
    - Transaction amounts
    - Transaction ranges
    - People and businesses
    - Same-person money movement
    - Daily transaction activity
    - Balance trends
    - Transaction status
    """
)


# ============================================================
# REQUIRED SHEETS
# ============================================================

REQUIRED_SHEETS = [
    "Transactions",
    "Statement Summary",
    "OCR Raw Rows",
    "Amount Range Analysis",
    "Same Person Analysis"
]


# ============================================================
# REQUIRED TRANSACTION COLUMNS
# ============================================================

REQUIRED_COLUMNS = [
    "Receipt No.",
    "Completion Time",
    "Details",
    "Transaction Status",
    "Paid In",
    "Withdrawn",
    "Balance",
    "Transaction Type"
]


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def clean_money(value):
    """
    Convert money values such as:
    Ksh 1,000.00
    1,000
    1000
    to numeric values.
    """

    if pd.isna(value):
        return 0.0

    value = str(value)

    value = (
        value.replace("Ksh", "")
        .replace("KES", "")
        .replace("ksh", "")
        .replace(",", "")
        .replace(" ", "")
    )

    value = value.replace("(", "-").replace(")", "")

    try:
        return float(value)
    except:
        return 0.0


def detect_transaction_type(details):

    text = str(details).lower()

    # Money received
    money_in_words = [
        "funds received",
        "received from",
        "money received",
        "receive international",
        "received international",
        "cash deposit",
        "deposit"
    ]

    for word in money_in_words:
        if word in text:
            return "Money In"

    # Withdrawals
    withdrawal_words = [
        "withdraw",
        "withdrawal",
        "cash withdrawal"
    ]

    for word in withdrawal_words:
        if word in text:
            return "Withdrawal"

    # Merchant payments
    merchant_words = [
        "merchant payment",
        "buy goods",
        "till",
        "merchant"
    ]

    for word in merchant_words:
        if word in text:
            return "Merchant Payment"

    # Paybill
    paybill_words = [
        "pay bill",
        "paybill",
        "business number"
    ]

    for word in paybill_words:
        if word in text:
            return "Paybill"

    # Airtime
    airtime_words = [
        "airtime"
    ]

    for word in airtime_words:
        if word in text:
            return "Airtime"

    # Transfers
    transfer_words = [
        "customer transfer to",
        "send money",
        "sent to",
        "customer payment",
        "transfer to"
    ]

    for word in transfer_words:
        if word in text:
            return "Money Out"

    # Charges
    charge_words = [
        "charge",
        "fee",
        "transaction cost"
    ]

    for word in charge_words:
        if word in text:
            return "Charge"

    return "Other"


def extract_person_or_business(details):

    text = str(details)

    patterns = [
        "Funds received from",
        "Funds received",
        "Customer Transfer to",
        "Customer transfer to",
        "Merchant Payment",
        "Pay Bill",
        "Airtime Purchase",
        "Withdrawal"
    ]

    for pattern in patterns:

        if pattern.lower() in text.lower():

            parts = text.split(pattern)

            if len(parts) > 1:

                name = parts[1]

                name = name.strip()
                name = name.strip("-:")
                name = name.strip()

                if name:
                    return name

    return text


def assign_amount_range(amount):

    if amount < 500:
        return "Below 500"

    elif amount < 1000:
        return "500 - 999"

    elif amount < 5000:
        return "1,000 - 4,999"

    elif amount < 10000:
        return "5,000 - 9,999"

    elif amount < 50000:
        return "10,000 - 49,999"

    else:
        return "50,000+"


def safe_divide(a, b):

    if b == 0:
        return 0

    return a / b


# ============================================================
# EXCEL UPLOAD
# ============================================================

st.sidebar.header("Upload Excel File")

uploaded_file = st.sidebar.file_uploader(
    "Upload your M-Pesa Excel workbook",
    type=["xlsx", "xls"]
)


if uploaded_file is None:

    st.info(
        "Please upload your M-Pesa Excel workbook using the sidebar."
    )

    st.stop()


# ============================================================
# READ EXCEL
# ============================================================

try:

    excel_file = pd.ExcelFile(uploaded_file)

    available_sheets = excel_file.sheet_names

except Exception as e:

    st.error("The Excel file could not be opened.")

    st.error(str(e))

    st.stop()


# ============================================================
# DISPLAY AVAILABLE SHEETS
# ============================================================

st.sidebar.subheader("Excel Sheets")

for sheet in available_sheets:

    st.sidebar.write("✓", sheet)


# ============================================================
# LOAD TRANSACTIONS
# ============================================================

if "Transactions" not in available_sheets:

    st.error(
        "The Excel workbook must contain a sheet called 'Transactions'."
    )

    st.stop()


try:

    df = pd.read_excel(
        uploaded_file,
        sheet_name="Transactions"
    )

except Exception as e:

    st.error("Could not read the Transactions sheet.")

    st.error(str(e))

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
# CHECK REQUIRED COLUMNS
# ============================================================

missing_columns = [
    column
    for column in REQUIRED_COLUMNS
    if column not in df.columns
]


if missing_columns:

    st.error("The Transactions sheet is missing these columns:")

    for column in missing_columns:
        st.write("-", column)

    st.write("Columns found in your Excel file:")

    st.write(list(df.columns))

    st.stop()


# ============================================================
# CLEAN DATA
# ============================================================

df["Paid In"] = df["Paid In"].apply(clean_money)

df["Withdrawn"] = df["Withdrawn"].apply(clean_money)

df["Balance"] = df["Balance"].apply(clean_money)


# ============================================================
# DATE CLEANING
# ============================================================

df["Completion Time"] = pd.to_datetime(
    df["Completion Time"],
    errors="coerce"
)


# ============================================================
# CLEAN STATUS
# ============================================================

df["Transaction Status"] = (
    df["Transaction Status"]
    .fillna("Unknown")
    .astype(str)
    .str.strip()
)


# ============================================================
# TRANSACTION TYPE
# ============================================================

df["Transaction Type"] = (
    df["Transaction Type"]
    .fillna("")
    .astype(str)
    .str.strip()
)


# Detect transaction type where missing or Other

for index in df.index:

    current_type = df.loc[index, "Transaction Type"]

    if current_type == "" or current_type.lower() == "other":

        df.loc[index, "Transaction Type"] = detect_transaction_type(
            df.loc[index, "Details"]
        )


# ============================================================
# MONEY IN
# ============================================================

df["Money In"] = df["Paid In"].fillna(0)


# ============================================================
# MONEY OUT
# ============================================================

df["Money Out"] = df["Withdrawn"].fillna(0)


# ============================================================
# TRANSACTION AMOUNT
# ============================================================

df["Transaction Amount"] = (
    df["Money In"] + df["Money Out"]
)


# ============================================================
# AMOUNT RANGE
# ============================================================

df["Amount Range"] = (
    df["Transaction Amount"]
    .apply(assign_amount_range)
)


# ============================================================
# PERSON / BUSINESS
# ============================================================

df["Person / Business"] = (
    df["Details"]
    .apply(extract_person_or_business)
)


# ============================================================
# CHARGE FLAG
# ============================================================

df["Is Charge"] = (
    df["Transaction Type"]
    .astype(str)
    .str.contains(
        "charge|fee",
        case=False,
        na=False
    )
)


# ============================================================
# SORT DATA
# ============================================================

df = df.sort_values(
    by="Completion Time",
    ascending=True
).reset_index(drop=True)


# ============================================================
# SIDEBAR FILTERS
# ============================================================

st.sidebar.header("Filters")


# Transaction Type filter

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


# Amount Range filter

amount_ranges = [
    "Below 500",
    "500 - 999",
    "1,000 - 4,999",
    "5,000 - 9,999",
    "10,000 - 49,999",
    "50,000+"
]


existing_ranges = [
    x for x in amount_ranges
    if x in df["Amount Range"].unique()
]


selected_ranges = st.sidebar.multiselect(
    "Amount Range",
    existing_ranges,
    default=existing_ranges
)


# Date filter

valid_dates = df["Completion Time"].dropna()


if len(valid_dates) > 0:

    min_date = valid_dates.min().date()

    max_date = valid_dates.max().date()

    selected_dates = st.sidebar.date_input(
        "Date Range",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date
    )

else:

    selected_dates = None


# ============================================================
# APPLY FILTERS
# ============================================================

filtered_df = df.copy()


if selected_types:

    filtered_df = filtered_df[
        filtered_df["Transaction Type"].isin(
            selected_types
        )
    ]


if selected_ranges:

    filtered_df = filtered_df[
        filtered_df["Amount Range"].isin(
            selected_ranges
        )
    ]


if selected_dates and len(selected_dates) == 2:

    start_date = pd.Timestamp(
        selected_dates[0]
    )

    end_date = pd.Timestamp(
        selected_dates[1]
    ) + pd.Timedelta(days=1)

    filtered_df = filtered_df[
        (
            filtered_df["Completion Time"] >= start_date
        )
        &
        (
            filtered_df["Completion Time"] < end_date
        )
    ]


# ============================================================
# HEADER
# ============================================================

st.header("M-Pesa Financial Overview")


# ============================================================
# KPI CALCULATIONS
# ============================================================

total_money_in = filtered_df["Money In"].sum()

total_money_out = filtered_df["Money Out"].sum()

net_cash_flow = (
    total_money_in - total_money_out
)

transaction_count = len(filtered_df)

average_transaction = (
    filtered_df["Transaction Amount"].mean()
    if transaction_count > 0
    else 0
)

largest_transaction = (
    filtered_df["Transaction Amount"].max()
    if transaction_count > 0
    else 0
)


# ============================================================
# KPI CARDS
# ============================================================

col1, col2, col3, col4, col5, col6 = st.columns(6)


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
        f"KSh {net_cash_flow:,.2f}"
    )


with col4:

    st.metric(
        "Transactions",
        f"{transaction_count:,}"
    )


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


# ============================================================
# MONEY IN VS MONEY OUT
# ============================================================

st.subheader("Money In vs Money Out")


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


fig = px.bar(
    money_summary,
    x="Category",
    y="Amount",
    text_auto=".2f",
    title="Total Money In vs Money Out"
)


fig.update_layout(
    template="plotly_white",
    yaxis_title="Amount (KSh)",
    xaxis_title=""
)


st.plotly_chart(
    fig,
    use_container_width=True
)


# ============================================================
# MONEY MOVEMENT PIE CHART
# ============================================================

st.subheader("Money Movement Distribution")


fig = px.pie(
    money_summary,
    names="Category",
    values="Amount",
    hole=0.4,
    title="Money Movement"
)


fig.update_layout(
    template="plotly_white"
)


st.plotly_chart(
    fig,
    use_container_width=True
)


# ============================================================
# TRANSACTION TYPE ANALYSIS
# ============================================================

st.subheader("Amount by Transaction Type")


type_analysis = (
    filtered_df
    .groupby("Transaction Type", as_index=False)
    .agg(
        Total_Amount=("Transaction Amount", "sum"),
        Transaction_Count=("Transaction Type", "count")
    )
    .sort_values(
        "Total_Amount",
        ascending=False
    )
)


fig = px.bar(
    type_analysis,
    x="Transaction Type",
    y="Total_Amount",
    text_auto=".2f",
    title="Total Amount by Transaction Type"
)


fig.update_layout(
    template="plotly_white",
    xaxis_tickangle=-45,
    yaxis_title="Amount (KSh)"
)


st.plotly_chart(
    fig,
    use_container_width=True
)


# ============================================================
# TRANSACTION COUNT BY TYPE
# ============================================================

fig = px.bar(
    type_analysis,
    x="Transaction Type",
    y="Transaction_Count",
    text_auto=True,
    title="Number of Transactions by Type"
)


fig.update_layout(
    template="plotly_white",
    xaxis_tickangle=-45,
    yaxis_title="Number of Transactions"
)


st.plotly_chart(
    fig,
    use_container_width=True
)


# ============================================================
# AMOUNT RANGE ANALYSIS
# ============================================================

st.subheader("Amount Range Analysis")


range_analysis = (
    filtered_df
    .groupby("Amount Range", as_index=False)
    .agg(
        Total_Amount=("Transaction Amount", "sum"),
        Transaction_Count=("Transaction Amount", "count")
    )
)


range_analysis["Amount Range"] = pd.Categorical(
    range_analysis["Amount Range"],
    categories=amount_ranges,
    ordered=True
)


range_analysis = range_analysis.sort_values(
    "Amount Range"
)


fig = px.bar(
    range_analysis,
    x="Amount Range",
    y="Transaction_Count",
    text_auto=True,
    title="Transactions by Amount Range"
)


fig.update_layout(
    template="plotly_white",
    xaxis_title="Amount Range",
    yaxis_title="Number of Transactions"
)


st.plotly_chart(
    fig,
    use_container_width=True
)


# ============================================================
# AMOUNT HISTOGRAM
# ============================================================

st.subheader("Transaction Amount Distribution")


fig = px.histogram(
    filtered_df,
    x="Transaction Amount",
    nbins=30,
    title="Distribution of Transaction Amounts"
)


fig.update_layout(
    template="plotly_white",
    xaxis_title="Transaction Amount (KSh)",
    yaxis_title="Number of Transactions"
)


st.plotly_chart(
    fig,
    use_container_width=True
)


# ============================================================
# TOP MONEY-IN SOURCES
# ============================================================

st.subheader("Top Money-In Sources")


money_in_df = filtered_df[
    filtered_df["Money In"] > 0
].copy()


if len(money_in_df) > 0:

    top_sources = (
        money_in_df
        .groupby(
            "Person / Business",
            as_index=False
        )["Money In"]
        .sum()
        .sort_values(
            "Money In",
            ascending=False
        )
        .head(10)
    )

    fig = px.bar(
        top_sources,
        x="Money In",
        y="Person / Business",
        orientation="h",
        title="Top 10 Money-In Sources",
        text_auto=".2f"
    )

    fig.update_layout(
        template="plotly_white",
        xaxis_title="Money Received (KSh)",
        yaxis_title=""
    )

    st.plotly_chart(
        fig,
        use_container_width=True
    )

else:

    st.info("No money-in transactions found.")


# ============================================================
# TOP MONEY-OUT DESTINATIONS
# ============================================================

st.subheader("Top Money-Out Destinations")


money_out_df = filtered_df[
    filtered_df["Money Out"] > 0
].copy()


if len(money_out_df) > 0:

    top_destinations = (
        money_out_df
        .groupby(
            "Person / Business",
            as_index=False
        )["Money Out"]
        .sum()
        .sort_values(
            "Money Out",
            ascending=False
        )
        .head(10)
    )

    fig = px.bar(
        top_destinations,
        x="Money Out",
        y="Person / Business",
        orientation="h",
        title="Top 10 Money-Out Destinations",
        text_auto=".2f"
    )

    fig.update_layout(
        template="plotly_white",
        xaxis_title="Money Sent (KSh)",
        yaxis_title=""
    )

    st.plotly_chart(
        fig,
        use_container_width=True
    )

else:

    st.info("No money-out transactions found.")


# ============================================================
# SAME PERSON ANALYSIS
# ============================================================

st.subheader("Same Person / Business Analysis")


received_by_person = (
    filtered_df[
        filtered_df["Money In"] > 0
    ]
    .groupby(
        "Person / Business"
    )["Money In"]
    .sum()
)


sent_by_person = (
    filtered_df[
        filtered_df["Money Out"] > 0
    ]
    .groupby(
        "Person / Business"
    )["Money Out"]
    .sum()
)


same_people = sorted(
    set(received_by_person.index)
    &
    set(sent_by_person.index)
)


if same_people:

    same_person_data = pd.DataFrame(
        {
            "Person / Business": same_people,
            "Received": [
                received_by_person.get(
                    person,
                    0
                )
                for person in same_people
            ],
            "Sent": [
                sent_by_person.get(
                    person,
                    0
                )
                for person in same_people
            ]
        }
    )


    same_person_data["Net"] = (
        same_person_data["Received"]
        -
        same_person_data["Sent"]
    )


    fig = go.Figure()


    fig.add_trace(
        go.Bar(
            x=same_person_data[
                "Person / Business"
            ],
            y=same_person_data["Received"],
            name="Received"
        )
    )


    fig.add_trace(
        go.Bar(
            x=same_person_data[
                "Person / Business"
            ],
            y=same_person_data["Sent"],
            name="Sent"
        )
    )


    fig.update_layout(
        barmode="group",
        title="Money Received and Sent to the Same Person / Business",
        template="plotly_white",
        xaxis_tickangle=-45,
        yaxis_title="Amount (KSh)"
    )


    st.plotly_chart(
        fig,
        use_container_width=True
    )


    st.dataframe(
        same_person_data,
        use_container_width=True
    )

else:

    st.info(
        "No person or business appears on both money-in and money-out transactions."
    )


# ============================================================
# DAILY MONEY MOVEMENT
# ============================================================

st.subheader("Daily Money Movement")


daily_df = filtered_df.dropna(
    subset=["Completion Time"]
).copy()


if len(daily_df) > 0:

    daily_df["Date"] = (
        daily_df["Completion Time"]
        .dt.date
    )


    daily_analysis = (
        daily_df
        .groupby("Date", as_index=False)
        .agg(
            Money_In=("Money In", "sum"),
            Money_Out=("Money Out", "sum"),
            Transactions=("Transaction Amount", "count")
        )
    )


    fig = go.Figure()


    fig.add_trace(
        go.Scatter(
            x=daily_analysis["Date"],
            y=daily_analysis["Money_In"],
            mode="lines+markers",
            name="Money In"
        )
    )


    fig.add_trace(
        go.Scatter(
            x=daily_analysis["Date"],
            y=daily_analysis["Money_Out"],
            mode="lines+markers",
            name="Money Out"
        )
    )


    fig.update_layout(
        title="Daily Money In and Money Out",
        template="plotly_white",
        xaxis_title="Date",
        yaxis_title="Amount (KSh)"
    )


    st.plotly_chart(
        fig,
        use_container_width=True
    )


# ============================================================
# DAILY TRANSACTION COUNT
# ============================================================

st.subheader("Daily Transaction Count")


if len(daily_df) > 0:

    daily_count = (
        daily_df
        .groupby("Date")
        .size()
        .reset_index(
            name="Transactions"
        )
    )


    fig = px.bar(
        daily_count,
        x="Date",
        y="Transactions",
        text_auto=True,
        title="Transactions per Day"
    )


    fig.update_layout(
        template="plotly_white",
        xaxis_title="Date",
        yaxis_title="Number of Transactions"
    )


    st.plotly_chart(
        fig,
        use_container_width=True
    )


# ============================================================
# BALANCE TREND
# ============================================================

st.subheader("M-Pesa Balance Trend")


balance_df = filtered_df.dropna(
    subset=["Completion Time"]
).copy()


if len(balance_df) > 0:

    balance_df = balance_df.sort_values(
        "Completion Time"
    )


    fig = px.line(
        balance_df,
        x="Completion Time",
        y="Balance",
        markers=True,
        title="Balance Over Time"
    )


    fig.update_layout(
        template="plotly_white",
        xaxis_title="Completion Time",
        yaxis_title="Balance (KSh)"
    )


    st.plotly_chart(
        fig,
        use_container_width=True
    )


# ============================================================
# TRANSACTION STATUS
# ============================================================

st.subheader("Transaction Status Analysis")


status_analysis = (
    filtered_df
    .groupby(
        "Transaction Status"
    )
    .size()
    .reset_index(
        name="Transactions"
    )
)


fig = px.pie(
    status_analysis,
    names="Transaction Status",
    values="Transactions",
    hole=0.4,
    title="Transaction Status"
)


fig.update_layout(
    template="plotly_white"
)


st.plotly_chart(
    fig,
    use_container_width=True
)


# ============================================================
# CHARGES AND FEES
# ============================================================

st.subheader("Charges and Fees")


charges_df = filtered_df[
    filtered_df["Is Charge"] == True
]


total_charges = (
    charges_df["Transaction Amount"].sum()
)


charge_count = len(charges_df)


col1, col2 = st.columns(2)


with col1:

    st.metric(
        "Total Charges",
        f"KSh {total_charges:,.2f}"
    )


with col2:

    st.metric(
        "Number of Charges",
        f"{charge_count:,}"
    )


if len(charges_df) > 0:

    charge_analysis = (
        charges_df
        .groupby(
            "Transaction Type",
            as_index=False
        )["Transaction Amount"]
        .sum()
    )


    fig = px.bar(
        charge_analysis,
        x="Transaction Type",
        y="Transaction Amount",
        text_auto=".2f",
        title="Charges by Type"
    )


    fig.update_layout(
        template="plotly_white",
        yaxis_title="Amount (KSh)"
    )


    st.plotly_chart(
        fig,
        use_container_width=True
    )


# ============================================================
# ORIGINAL STATEMENT SUMMARY SHEET
# ============================================================

st.subheader("Statement Summary")


if "Statement Summary" in available_sheets:

    try:

        summary_df = pd.read_excel(
            uploaded_file,
            sheet_name="Statement Summary"
        )


        summary_df = summary_df.dropna(
            how="all"
        )


        st.dataframe(
            summary_df,
            use_container_width=True
        )


    except Exception as e:

        st.warning(
            f"Could not read Statement Summary: {e}"
        )


# ============================================================
# AMOUNT RANGE SHEET
# ============================================================

st.subheader("Amount Range Analysis Sheet")


if "Amount Range Analysis" in available_sheets:

    try:

        amount_range_df = pd.read_excel(
            uploaded_file,
            sheet_name="Amount Range Analysis"
        )


        amount_range_df = amount_range_df.dropna(
            how="all"
        )


        st.dataframe(
            amount_range_df,
            use_container_width=True
        )


    except Exception as e:

        st.warning(
            f"Could not read Amount Range Analysis: {e}"
        )


# ============================================================
# SAME PERSON SHEET
# ============================================================

st.subheader("Same Person Analysis Sheet")


if "Same Person Analysis" in available_sheets:

    try:

        same_person_sheet = pd.read_excel(
            uploaded_file,
            sheet_name="Same Person Analysis"
        )


        same_person_sheet = same_person_sheet.dropna(
            how="all"
        )


        st.dataframe(
            same_person_sheet,
            use_container_width=True
        )


    except Exception as e:

        st.warning(
            f"Could not read Same Person Analysis: {e}"
        )


# ============================================================
# OCR RAW DATA
# ============================================================

st.subheader("OCR Raw Rows")


if "OCR Raw Rows" in available_sheets:

    try:

        ocr_df = pd.read_excel(
            uploaded_file,
            sheet_name="OCR Raw Rows"
        )


        ocr_df = ocr_df.dropna(
            how="all"
        )


        with st.expander(
            "Show OCR Raw Data"
        ):

            st.dataframe(
                ocr_df,
                use_container_width=True
            )


    except Exception as e:

        st.warning(
            f"Could not read OCR Raw Rows: {e}"
        )


# ============================================================
# COMPLETE TRANSACTION TABLE
# ============================================================

st.subheader("Complete Transaction Table")


display_columns = [
    "Receipt No.",
    "Completion Time",
    "Details",
    "Transaction Status",
    "Paid In",
    "Withdrawn",
    "Balance",
    "Transaction Type",
    "Money In",
    "Money Out",
    "Transaction Amount",
    "Amount Range",
    "Person / Business",
    "Is Charge"
]


display_columns = [
    column
    for column in display_columns
    if column in filtered_df.columns
]


st.dataframe(
    filtered_df[display_columns],
    use_container_width=True,
    height=600
)


# ============================================================
# DESCRIPTIVE STATISTICS
# ============================================================

st.subheader("Transaction Statistics")


numeric_columns = [
    "Paid In",
    "Withdrawn",
    "Balance",
    "Money In",
    "Money Out",
    "Transaction Amount"
]


available_numeric = [
    column
    for column in numeric_columns
    if column in filtered_df.columns
]


if available_numeric:

    statistics = (
        filtered_df[available_numeric]
        .describe()
        .T
    )


    statistics = statistics.round(2)


    st.dataframe(
        statistics,
        use_container_width=True
    )


# ============================================================
# TOP TRANSACTIONS
# ============================================================

st.subheader("Largest Transactions")


largest_transactions = (
    filtered_df
    .sort_values(
        "Transaction Amount",
        ascending=False
    )
    .head(10)
)


st.dataframe(
    largest_transactions[display_columns],
    use_container_width=True
)


# ============================================================
# MONEY-IN TRANSACTIONS TABLE
# ============================================================

st.subheader("Money-In Transactions")


money_in_display = filtered_df[
    filtered_df["Money In"] > 0
]


st.dataframe(
    money_in_display[display_columns],
    use_container_width=True
)


# ============================================================
# MONEY-OUT TRANSACTIONS TABLE
# ============================================================

st.subheader("Money-Out Transactions")


money_out_display = filtered_df[
    filtered_df["Money Out"] > 0
]


st.dataframe(
    money_out_display[display_columns],
    use_container_width=True
)


# ============================================================
# DOWNLOAD FILTERED DATA
# ============================================================

st.subheader("Download Data")


csv_data = filtered_df.to_csv(
    index=False
).encode("utf-8")


st.download_button(
    label="Download Filtered Transactions CSV",
    data=csv_data,
    file_name="filtered_mpesa_transactions.csv",
    mime="text/csv"
)


# ============================================================
# DOWNLOAD FINANCIAL SUMMARY
# ============================================================

financial_summary = pd.DataFrame(
    {
        "Metric": [
            "Total Money In",
            "Total Money Out",
            "Net Cash Flow",
            "Number of Transactions",
            "Average Transaction",
            "Largest Transaction",
            "Total Charges",
            "Number of Charges"
        ],
        "Value": [
            total_money_in,
            total_money_out,
            net_cash_flow,
            transaction_count,
            average_transaction,
            largest_transaction,
            total_charges,
            charge_count
        ]
    }
)


summary_csv = financial_summary.to_csv(
    index=False
).encode("utf-8")


st.download_button(
    label="Download Financial Summary",
    data=summary_csv,
    file_name="mpesa_financial_summary.csv",
    mime="text/csv"
)


# ============================================================
# FOOTER
# ============================================================

st.markdown("---")

st.caption(
    "M-Pesa Statement Analysis Dashboard"
)

st.caption(
    "Important: If the original statement was created from photos/OCR, "
    "check important amounts and names against the original statement."
)