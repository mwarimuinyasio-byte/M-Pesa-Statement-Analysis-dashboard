import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go


# ============================================================
# PAGE CONFIG
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
    "Upload your M-Pesa Excel workbook to analyze transactions, "
    "money in, money out, charges, transaction ranges, people, "
    "and account balances."
)


# ============================================================
# REQUIRED COLUMNS
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
# MONEY CLEANING FUNCTION
# ============================================================

def clean_money(value):

    if pd.isna(value):
        return np.nan

    value = str(value).strip()

    # Remove currency labels
    value = value.replace("KSh", "")
    value = value.replace("KES", "")
    value = value.replace("ksh", "")
    value = value.replace("Kes", "")

    # Remove commas and spaces
    value = value.replace(",", "")
    value = value.replace(" ", "")

    # Handle brackets
    value = value.replace("(", "-")
    value = value.replace(")", "")

    try:
        return float(value)

    except:
        return np.nan


# ============================================================
# TRANSACTION TYPE DETECTION
# ============================================================

def detect_transaction_type(details):

    text = str(details).lower()

    if any(
        word in text
        for word in [
            "funds received",
            "received from",
            "money received",
            "receive international",
            "received international"
        ]
    ):
        return "Money In"

    if any(
        word in text
        for word in [
            "withdraw",
            "withdrawal"
        ]
    ):
        return "Withdrawal"

    if any(
        word in text
        for word in [
            "merchant payment",
            "buy goods",
            "till"
        ]
    ):
        return "Merchant Payment"

    if any(
        word in text
        for word in [
            "pay bill",
            "paybill"
        ]
    ):
        return "Paybill"

    if "airtime" in text:
        return "Airtime"

    if any(
        word in text
        for word in [
            "customer transfer",
            "send money",
            "sent to",
            "transfer to"
        ]
    ):
        return "Money Out"

    if any(
        word in text
        for word in [
            "charge",
            "fee"
        ]
    ):
        return "Charge"

    return "Other"


# ============================================================
# PERSON / BUSINESS EXTRACTION
# ============================================================

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

            parts = text.lower().split(
                pattern.lower()
            )

            if len(parts) > 1:

                result = parts[1].strip()

                return result

    return text


# ============================================================
# AMOUNT RANGE
# ============================================================

def amount_range(amount):

    if pd.isna(amount):
        return "Unknown"

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


# ============================================================
# UPLOAD FILE
# ============================================================

st.sidebar.header("Upload Excel Workbook")

uploaded_file = st.sidebar.file_uploader(
    "Choose your M-Pesa Excel file",
    type=["xlsx", "xls"]
)


if uploaded_file is None:

    st.info(
        "Upload your M-Pesa Excel workbook from the sidebar."
    )

    st.stop()


# ============================================================
# READ EXCEL
# ============================================================

try:

    excel = pd.ExcelFile(uploaded_file)

except Exception as e:

    st.error("Unable to open the Excel file.")

    st.error(str(e))

    st.stop()


# ============================================================
# SHOW SHEETS
# ============================================================

st.sidebar.subheader("Available Sheets")

for sheet in excel.sheet_names:

    st.sidebar.write("✓", sheet)


# ============================================================
# LOAD TRANSACTIONS
# ============================================================

if "Transactions" not in excel.sheet_names:

    st.error(
        "Your Excel workbook does not contain a 'Transactions' sheet."
    )

    st.stop()


df = pd.read_excel(
    uploaded_file,
    sheet_name="Transactions"
)


# ============================================================
# CLEAN COLUMN NAMES
# ============================================================

df.columns = (
    df.columns
    .astype(str)
    .str.strip()
)


# ============================================================
# CHECK COLUMNS
# ============================================================

missing = [
    column
    for column in REQUIRED_COLUMNS
    if column not in df.columns
]


if missing:

    st.error("The following columns are missing:")

    for column in missing:
        st.write(f"- {column}")

    st.write("Columns found:")

    st.write(list(df.columns))

    st.stop()


# ============================================================
# CLEAN DATE
# ============================================================

df["Completion Time"] = pd.to_datetime(
    df["Completion Time"],
    errors="coerce"
)


# ============================================================
# CLEAN MONEY COLUMNS
# ============================================================

df["Paid In"] = df["Paid In"].apply(
    clean_money
)

df["Withdrawn"] = df["Withdrawn"].apply(
    clean_money
)

df["Balance"] = df["Balance"].apply(
    clean_money
)


# ============================================================
# IMPORTANT:
# SORT BY TIME BEFORE ANALYZING BALANCE
# ============================================================

df = df.sort_values(
    by="Completion Time",
    ascending=True
).reset_index(drop=True)


# ============================================================
# FILL MISSING MONEY VALUES
# ============================================================

df["Paid In"] = df["Paid In"].fillna(0)

df["Withdrawn"] = df["Withdrawn"].fillna(0)


# ============================================================
# TRANSACTION TYPE
# ============================================================

df["Transaction Type"] = (
    df["Transaction Type"]
    .fillna("")
    .astype(str)
    .str.strip()
)


for i in df.index:

    if (
        df.loc[i, "Transaction Type"] == ""
        or
        df.loc[i, "Transaction Type"].lower()
        == "other"
    ):

        df.loc[i, "Transaction Type"] = (
            detect_transaction_type(
                df.loc[i, "Details"]
            )
        )


# ============================================================
# MONEY IN
# ============================================================

df["Money In"] = df["Paid In"]


# ============================================================
# MONEY OUT
# ============================================================

df["Money Out"] = df["Withdrawn"]


# ============================================================
# TRANSACTION AMOUNT
# ============================================================

df["Transaction Amount"] = (
    df["Money In"]
    +
    df["Money Out"]
)


# ============================================================
# AMOUNT RANGE
# ============================================================

df["Amount Range"] = (
    df["Transaction Amount"]
    .apply(amount_range)
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
# BALANCE ANALYSIS
# ============================================================

balance_df = df[
    df["Balance"].notna()
].copy()


if len(balance_df) > 0:

    balance_df = balance_df.sort_values(
        "Completion Time"
    )

    opening_balance = (
        balance_df.iloc[0]["Balance"]
    )

    closing_balance = (
        balance_df.iloc[-1]["Balance"]
    )

    highest_balance = (
        balance_df["Balance"].max()
    )

    lowest_balance = (
        balance_df["Balance"].min()
    )

else:

    opening_balance = 0
    closing_balance = 0
    highest_balance = 0
    lowest_balance = 0


# ============================================================
# SIDEBAR FILTERS
# ============================================================

st.sidebar.header("Filters")


transaction_types = sorted(
    df["Transaction Type"]
    .dropna()
    .unique()
)


selected_types = st.sidebar.multiselect(
    "Transaction Type",
    transaction_types,
    default=list(transaction_types)
)


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


# ============================================================
# MAIN HEADER
# ============================================================

st.header("Account Balance Overview")


# ============================================================
# BALANCE CARDS
# ============================================================

col1, col2, col3, col4 = st.columns(4)


with col1:

    st.metric(
        "Opening Balance",
        f"KSh {opening_balance:,.2f}"
    )


with col2:

    st.metric(
        "Closing Balance",
        f"KSh {closing_balance:,.2f}"
    )


with col3:

    st.metric(
        "Highest Balance",
        f"KSh {highest_balance:,.2f}"
    )


with col4:

    st.metric(
        "Lowest Balance",
        f"KSh {lowest_balance:,.2f}"
    )


# ============================================================
# MONEY OVERVIEW
# ============================================================

st.subheader("Money Overview")


total_money_in = filtered_df["Money In"].sum()

total_money_out = filtered_df["Money Out"].sum()

net_cash_flow = (
    total_money_in
    -
    total_money_out
)

transaction_count = len(filtered_df)


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
        f"KSh {net_cash_flow:,.2f}"
    )


with col4:

    st.metric(
        "Transactions",
        f"{transaction_count:,}"
    )


# ============================================================
# BALANCE TREND
# ============================================================

st.subheader("M-Pesa Balance Over Time")


if len(balance_df) > 0:

    fig = go.Figure()


    fig.add_trace(
        go.Scatter(
            x=balance_df["Completion Time"],
            y=balance_df["Balance"],
            mode="lines+markers",
            name="M-Pesa Balance",
            hovertemplate=(
                "<b>Date:</b> %{x}<br>"
                "<b>Balance:</b> KSh %{y:,.2f}"
                "<extra></extra>"
            )
        )
    )


    fig.update_layout(
        title="Running M-Pesa Account Balance",
        template="plotly_white",
        xaxis_title="Completion Time",
        yaxis_title="Balance (KSh)",
        hovermode="x unified"
    )


    st.plotly_chart(
        fig,
        use_container_width=True
    )


# ============================================================
# BALANCE TABLE
# ============================================================

st.subheader("Balance History")


balance_display = balance_df[
    [
        "Receipt No.",
        "Completion Time",
        "Details",
        "Paid In",
        "Withdrawn",
        "Balance"
    ]
].copy()


# Format money columns

for column in [
    "Paid In",
    "Withdrawn",
    "Balance"
]:

    balance_display[column] = (
        balance_display[column]
        .apply(
            lambda x:
            f"KSh {x:,.2f}"
            if pd.notna(x)
            else ""
        )
    )


st.dataframe(
    balance_display,
    use_container_width=True,
    height=500
)


# ============================================================
# BALANCE CHANGE ANALYSIS
# ============================================================

st.subheader("Balance Change Analysis")


balance_change_df = balance_df.copy()


balance_change_df["Previous Balance"] = (
    balance_change_df["Balance"]
    .shift(1)
)


balance_change_df["Balance Change"] = (
    balance_change_df["Balance"]
    -
    balance_change_df["Previous Balance"]
)


balance_change_df = balance_change_df[
    [
        "Completion Time",
        "Details",
        "Previous Balance",
        "Paid In",
        "Withdrawn",
        "Balance",
        "Balance Change"
    ]
]


st.dataframe(
    balance_change_df,
    use_container_width=True
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
    title="Money In vs Money Out"
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
# PIE CHART
# ============================================================

fig = px.pie(
    money_summary,
    names="Category",
    values="Amount",
    hole=0.4,
    title="Money Movement Distribution"
)


fig.update_layout(
    template="plotly_white"
)


st.plotly_chart(
    fig,
    use_container_width=True
)


# ============================================================
# TRANSACTION TYPES
# ============================================================

st.subheader("Transaction Type Analysis")


type_analysis = (
    filtered_df
    .groupby(
        "Transaction Type",
        as_index=False
    )
    .agg(
        Total_Amount=(
            "Transaction Amount",
            "sum"
        ),
        Transaction_Count=(
            "Transaction Type",
            "count"
        )
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
    title="Amount by Transaction Type"
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
# AMOUNT RANGE
# ============================================================

st.subheader("Amount Range Analysis")


range_analysis = (
    filtered_df
    .groupby(
        "Amount Range",
        as_index=False
    )
    .agg(
        Transaction_Count=(
            "Transaction Amount",
            "count"
        ),
        Total_Amount=(
            "Transaction Amount",
            "sum"
        )
    )
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


    daily = (
        daily_df
        .groupby(
            "Date",
            as_index=False
        )
        .agg(
            Money_In=("Money In", "sum"),
            Money_Out=("Money Out", "sum")
        )
    )


    fig = go.Figure()


    fig.add_trace(
        go.Scatter(
            x=daily["Date"],
            y=daily["Money_In"],
            mode="lines+markers",
            name="Money In"
        )
    )


    fig.add_trace(
        go.Scatter(
            x=daily["Date"],
            y=daily["Money_Out"],
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
# TRANSACTION AMOUNT HISTOGRAM
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
# TRANSACTION STATUS
# ============================================================

st.subheader("Transaction Status")


status = (
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
    status,
    names="Transaction Status",
    values="Transactions",
    hole=0.4,
    title="Transaction Status Distribution"
)


fig.update_layout(
    template="plotly_white"
)


st.plotly_chart(
    fig,
    use_container_width=True
)


# ============================================================
# TOP MONEY-IN SOURCES
# ============================================================

st.subheader("Top Money-In Sources")


money_in = filtered_df[
    filtered_df["Money In"] > 0
]


if len(money_in) > 0:

    sources = (
        money_in
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
        sources,
        x="Money In",
        y="Person / Business",
        orientation="h",
        text_auto=".2f",
        title="Top 10 Money-In Sources"
    )


    fig.update_layout(
        template="plotly_white",
        xaxis_title="Money Received (KSh)"
    )


    st.plotly_chart(
        fig,
        use_container_width=True
    )


# ============================================================
# TOP MONEY-OUT DESTINATIONS
# ============================================================

st.subheader("Top Money-Out Destinations")


money_out = filtered_df[
    filtered_df["Money Out"] > 0
]


if len(money_out) > 0:

    destinations = (
        money_out
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
        destinations,
        x="Money Out",
        y="Person / Business",
        orientation="h",
        text_auto=".2f",
        title="Top 10 Money-Out Destinations"
    )


    fig.update_layout(
        template="plotly_white",
        xaxis_title="Money Sent (KSh)"
    )


    st.plotly_chart(
        fig,
        use_container_width=True
    )


# ============================================================
# SAME PERSON ANALYSIS
# ============================================================

st.subheader("Same Person / Business Analysis")


received = (
    filtered_df[
        filtered_df["Money In"] > 0
    ]
    .groupby(
        "Person / Business"
    )["Money In"]
    .sum()
)


sent = (
    filtered_df[
        filtered_df["Money Out"] > 0
    ]
    .groupby(
        "Person / Business"
    )["Money Out"]
    .sum()
)


same_people = sorted(
    set(received.index)
    &
    set(sent.index)
)


if same_people:

    same_person = pd.DataFrame(
        {
            "Person / Business": same_people,
            "Received": [
                received.get(x, 0)
                for x in same_people
            ],
            "Sent": [
                sent.get(x, 0)
                for x in same_people
            ]
        }
    )


    same_person["Net"] = (
        same_person["Received"]
        -
        same_person["Sent"]
    )


    st.dataframe(
        same_person,
        use_container_width=True
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
    x
    for x in display_columns
    if x in filtered_df.columns
]


# Create a copy so we don't change the underlying data

display_df = filtered_df[
    display_columns
].copy()


# ============================================================
# FORMAT MONEY FOR DISPLAY
# ============================================================

money_columns = [
    "Paid In",
    "Withdrawn",
    "Balance",
    "Money In",
    "Money Out",
    "Transaction Amount"
]


for column in money_columns:

    if column in display_df.columns:

        display_df[column] = (
            display_df[column]
            .apply(
                lambda x:
                f"KSh {x:,.2f}"
                if pd.notna(x)
                else ""
            )
        )


st.dataframe(
    display_df,
    use_container_width=True,
    height=600
)


# ============================================================
# LARGEST TRANSACTIONS
# ============================================================

st.subheader("Largest Transactions")


largest = (
    filtered_df
    .sort_values(
        "Transaction Amount",
        ascending=False
    )
    .head(10)
)


largest_display = largest[
    display_columns
].copy()


for column in money_columns:

    if column in largest_display.columns:

        largest_display[column] = (
            largest_display[column]
            .apply(
                lambda x:
                f"KSh {x:,.2f}"
                if pd.notna(x)
                else ""
            )
        )


st.dataframe(
    largest_display,
    use_container_width=True
)


# ============================================================
# ORIGINAL STATEMENT SUMMARY
# ============================================================

if "Statement Summary" in excel.sheet_names:

    st.subheader("Statement Summary Sheet")

    try:

        statement_summary = pd.read_excel(
            uploaded_file,
            sheet_name="Statement Summary"
        )

        st.dataframe(
            statement_summary,
            use_container_width=True
        )

    except Exception as e:

        st.warning(
            f"Unable to read Statement Summary: {e}"
        )


# ============================================================
# OCR RAW ROWS
# ============================================================

if "OCR Raw Rows" in excel.sheet_names:

    with st.expander("View OCR Raw Rows"):

        try:

            ocr = pd.read_excel(
                uploaded_file,
                sheet_name="OCR Raw Rows"
            )

            st.dataframe(
                ocr,
                use_container_width=True
            )

        except Exception as e:

            st.warning(
                f"Unable to read OCR Raw Rows: {e}"
            )


# ============================================================
# DOWNLOAD FILTERED DATA
# ============================================================

st.subheader("Download")


csv = filtered_df.to_csv(
    index=False
).encode("utf-8")


st.download_button(
    "Download Filtered Transactions",
    data=csv,
    file_name="mpesa_filtered_transactions.csv",
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
    "For OCR-generated statements, verify important amounts "
    "and balances against the original statement."
)