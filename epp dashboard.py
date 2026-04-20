import streamlit as st
import pandas as pd
import plotly.express as px

# ==============================
# PAGE CONFIG
# ==============================
st.set_page_config(page_title="EPP HSE Dashboard", layout="wide")

st.title("📊 EPP HSE Performance Dashboard")

uploaded_file = st.file_uploader("Upload EPP HSE File", type=["xlsx"])

if uploaded_file:

    st.success(f"Loaded File: {uploaded_file.name}")

    # ==============================
    # READ FILE
    # ==============================
    df = pd.read_excel(uploaded_file, engine='openpyxl', header=1)

    # ==============================
    # CLEAN DATA
    # ==============================
    df.columns = df.columns.astype(str).str.strip()
    df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

    df = df[df["Date"].notna()]
    df = df[df["Date"] != "Annual Planned"]

    df["Date"] = pd.to_datetime(df["Date"], errors='coerce')
    df["Year"] = df["Date"].dt.year
    df = df.dropna(subset=["Year"])

    # ==============================
    # SMART COLUMN FINDER
    # ==============================
    def find_col(keyword):
        for col in df.columns:
            if keyword.lower() in col.lower():
                return col
        return None

    col_hours = find_col("EPP Total")
    col_training = find_col("Training")
    col_nearmiss = find_col("Near Miss")
    col_fatal = find_col("Fatal")
    col_ptw = find_col("PTW")
    col_lwdc = find_col("LWDC")
    col_mtc = find_col("MTC")
    col_fac = find_col("FAC")

    # ==============================
    # KPIs
    # ==============================
    st.subheader("📊 Key Metrics")

    k1, k2, k3, k4 = st.columns(4)
    k5, k6, k7, k8 = st.columns(4)

    k1.metric("Worked Hours", int(df[col_hours].sum()) if col_hours else 0)
    k2.metric("Training Hours", int(df[col_training].sum()) if col_training else 0)
    k3.metric("Near Miss", int(df[col_nearmiss].sum()) if col_nearmiss else 0)
    k4.metric("Fatal", int(df[col_fatal].sum()) if col_fatal else 0)

    k5.metric("PTW", int(df[col_ptw].sum()) if col_ptw else 0)
    k6.metric("LWDC", int(df[col_lwdc].sum()) if col_lwdc else 0)
    k7.metric("MTC", int(df[col_mtc].sum()) if col_mtc else 0)
    k8.metric("FAC", int(df[col_fac].sum()) if col_fac else 0)

    # ==============================
    # GROUP DATA
    # ==============================
    yearly = df.groupby("Year").sum(numeric_only=True).reset_index()

    # ==============================
    # CHARTS
    # ==============================
    st.subheader("📈 Trends & Analysis")

    # 🔹 Worked Hours (Area)
    if col_hours:
        fig1 = px.area(yearly, x="Year", y=col_hours,
                       title="Worked Man-Hours Trend")
        st.plotly_chart(fig1, use_container_width=True)

    # 🔹 Training Hours (Line like Near Miss)
    if col_training:
        fig2 = px.line(yearly, x="Year", y=col_training,
                       markers=True,
                       title="Training Hours Trend")
        st.plotly_chart(fig2, use_container_width=True)

    # 🔹 Near Miss
    if col_nearmiss:
        fig3 = px.line(yearly, x="Year", y=col_nearmiss,
                       markers=True,
                       title="Near Miss Trend")
        st.plotly_chart(fig3, use_container_width=True)

    # 🔹 PTW vs LWDC
    if col_ptw and col_lwdc:
        fig4 = px.bar(yearly, x="Year",
                      y=[col_ptw, col_lwdc],
                      barmode="group",
                      title="PTW vs LWDC")
        st.plotly_chart(fig4, use_container_width=True)

    # ==============================
    # SEPARATE INCIDENT CHARTS
    # ==============================
    st.subheader("📊 Incident Analysis")

    c1, c2, c3 = st.columns(3)

    if col_lwdc:
        fig_lwdc = px.line(yearly, x="Year", y=col_lwdc,
                           markers=True, title="LWDC Trend")
        c1.plotly_chart(fig_lwdc, use_container_width=True)

    if col_mtc:
        fig_mtc = px.line(yearly, x="Year", y=col_mtc,
                          markers=True, title="MTC Trend")
        c2.plotly_chart(fig_mtc, use_container_width=True)

    if col_fac:
        fig_fac = px.line(yearly, x="Year", y=col_fac,
                          markers=True, title="FAC Trend")
        c3.plotly_chart(fig_fac, use_container_width=True)

    # ==============================
    # PIE CHART
    # ==============================
    st.subheader("📊 Incident Distribution")

    incident_cols = [col_lwdc, col_mtc, col_fac]
    incident_cols = [c for c in incident_cols if c]

    if incident_cols:
        totals = [df[c].sum() for c in incident_cols]
        fig_pie = px.pie(names=incident_cols, values=totals,
                         title="Incident Distribution")
        st.plotly_chart(fig_pie, use_container_width=True)

    # ==============================
    # CORRELATION
    # ==============================
    st.subheader("📊 Correlation Matrix")

    corr = yearly.corr(numeric_only=True)
    fig_corr = px.imshow(corr, text_auto=True)
    st.plotly_chart(fig_corr, use_container_width=True)

    # ==============================
    # AI INSIGHTS
    # ==============================
    st.subheader("🤖 AI Insights")

    insights = []

    if col_hours and yearly[col_hours].iloc[-1] > yearly[col_hours].mean():
        insights.append("📈 Work activity increased.")

    if col_training and yearly[col_training].iloc[-1] < yearly[col_training].mean():
        insights.append("⚠️ Training needs improvement.")

    if col_nearmiss and yearly[col_nearmiss].iloc[-1] > yearly[col_nearmiss].mean():
        insights.append("🔍 Strong near miss reporting culture.")

    if col_lwdc and yearly[col_lwdc].sum() > 0:
        insights.append("🚨 Lost workday cases exist.")

    insights.append("✅ Maintain proactive safety culture.")
    insights.append("✅ Increase inspections and audits.")

    for i in insights:
        st.write(i)

    # ==============================
    # DATA TABLE
    # ==============================
    st.subheader("📄 Data Table")
    st.dataframe(df)

else:
    st.info("Upload your Excel file")
