import os
import re
from langchain_google_genai import ChatGoogleGenerativeAI
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from langchain_experimental.agents import create_pandas_dataframe_agent
from langchain_groq import ChatGroq

# Set page layout to wide
st.set_page_config(
    page_title="Real Estate & Workplace Dashboard", layout="wide"
)

# Fetch all credentials directly from .streamlit/secrets.toml
APP_USERNAME = st.secrets.get("APP_USERNAME", "admin")
APP_PASSWORD = st.secrets.get("APP_PASSWORD", "mll@2026")
GEMINI_API_KEY = st.secrets.get("GOOGLE_API_KEY", "")

# ==========================================
# AUTHENTICATION GATE
# ==========================================
if "authenticated" not in st.session_state:
  st.session_state.authenticated = False

if not st.session_state.authenticated:
  st.markdown(
      """
        <style>
            .login-container {
                max-width: 400px;
                margin: 100px auto;
                padding: 30px;
                background-color: #1e2530;
                border-radius: 12px;
                border: 1px solid #2d3748;
                box-shadow: 0 8px 16px rgba(0,0,0,0.4);
            }
        </style>
    """,
      unsafe_allow_html=True,
  )

  col1, col2, col3 = st.columns([1, 1.2, 1])
  with col2:
    st.markdown("## 🔐 Secure Login")
    st.markdown(
        "Please enter your credentials to access the Real Estate Dashboard."
    )

    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    if st.button("Login", use_container_width=True):
      if username == APP_USERNAME and password == APP_PASSWORD:
        st.session_state.authenticated = True
        st.rerun()
      else:
        st.error("Invalid username or password. Please try again.")
  st.stop()

# ==========================================
# CUSTOM STYLING
# ==========================================
st.markdown(
    """
    <style>
        .main { background-color: #0e1117; }
        div[data-testid="stMetric"] {
            background-color: #1e2530 !important;
            padding: 15px !important;
            border-radius: 10px !important;
            box-shadow: 0 4px 6px rgba(0,0,0,0.3) !important;
            border: 1px solid #2d3748 !important;
        }
        div[data-testid="stMetric"] label {
            color: #a0aec0 !important;
            font-size: 14px !important;
        }
        div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
            color: #ffffff !important;
            font-size: 26px !important;
            font-weight: 700 !important;
        }
        .stTextArea div[data-baseweb="base-input"], .stTextArea textarea {
            background-color: #1e2530 !important;
            color: #ffffff !important;
            border-color: #2d3748 !important;
            font-family: monospace;
            font-size: 14px;
        }
        .tooltip-container {
            position: relative;
            display: inline-block;
            cursor: pointer;
            width: 100%;
        }
        .tooltip-container .tooltip-text {
            visibility: hidden;
            width: 240px;
            background-color: #1a202c;
            color: #fff;
            text-align: left;
            border-radius: 6px;
            padding: 10px 14px;
            position: absolute;
            z-index: 9999;
            bottom: 100%;
            left: 50%;
            margin-left: -120px;
            margin-bottom: 10px;
            opacity: 0;
            transition: opacity 0.3s ease-in-out, visibility 0.3s ease-in-out;
            transition-delay: 0.2s;
            border: 1px solid #4a5568;
            font-size: 13px;
            box-shadow: 0px 4px 12px rgba(0,0,0,0.6);
            max-height: 200px;
            overflow-y: auto;
            pointer-events: auto;
        }
        .tooltip-container .tooltip-text::after {
            content: "";
            position: absolute;
            top: 100%;
            left: 0;
            width: 100%;
            height: 15px; 
            background: transparent;
        }
        .tooltip-container:hover .tooltip-text,
        .tooltip-text:hover {
            visibility: visible;
            opacity: 1;
            transition-delay: 0s;
        }
        .info-banner {
            background: linear-gradient(90deg, #b30000 0%, #ff1a1a 100%);
            color: white;
            padding: 15px;
            border-radius: 8px;
            text-align: center;
            font-family: 'Arial Black', sans-serif;
            text-transform: uppercase;
            letter-spacing: 2px;
            margin-bottom: 20px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.4);
        }
        .info-kpi-box {
            background-color: #1e2530;
            border-top: 4px solid #ff1a1a;
            padding: 15px;
            border-radius: 8px;
            text-align: center;
            box-shadow: 0 4px 6px rgba(0,0,0,0.3);
            margin-bottom: 15px;
        }
        .info-kpi-value {
            font-size: 28px;
            font-weight: bold;
            color: #ffffff;
            margin: 5px 0;
        }
        .info-kpi-label {
            font-size: 12px;
            color: #a0aec0;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        .info-section-header {
            background-color: #003366;
            color: white;
            padding: 8px 15px;
            border-radius: 5px;
            text-align: center;
            font-weight: bold;
            font-size: 14px;
            margin-bottom: 15px;
            text-transform: uppercase;
        }
    </style>
""",
    unsafe_allow_html=True,
)

# ==========================================
# SIDEBAR: FILE UPLOADER & LOGOUT
# ==========================================
st.sidebar.header("📁 Data Source")
uploaded_files = st.sidebar.file_uploader(
    "Upload your `prep.xlsx` workbook (Up to 10 files)", type=["xlsx", "xls"], accept_multiple_files=True
)

if st.sidebar.button("Log Out"):
  st.session_state.authenticated = False
  st.rerun()

if not uploaded_files:
  st.markdown(
      """
        <div style="text-align: center; margin-top: 100px;">
            <h2>👋 Welcome to the Dashboard</h2>
            <p style="color: #a0aec0; font-size: 16px;">Please upload your <b>prep.xlsx</b> workbook using the sidebar file uploader to load analytics.</p>
        </div>
    """,
      unsafe_allow_html=True,
  )
  st.stop()


# Load Data dynamically from uploaded files for original dashboard
@st.cache_data
def load_data(file):
  df_prep = pd.read_excel(file, sheet_name="prep")
  return df_prep

# Find the prep file to power the original tabs
prep_file = next((f for f in uploaded_files if "prep" in f.name.lower()), uploaded_files[0])

try:
    df_prep = load_data(prep_file)
except Exception:
    st.error("Could not find a 'prep' sheet. Please ensure one of the uploaded files is 'prep.xlsx'.")
    st.stop()

# Build the global dictionary of all sheets for Tab 5 (Chatbot)
dfs = {}
for file in uploaded_files:
    file.seek(0) # Reset file pointer after cache load
    excel_obj = pd.ExcelFile(file)
    for sheet in excel_obj.sheet_names:
        dfs[f"{file.name} [{sheet}]"] = pd.read_excel(file, sheet_name=sheet)

# ==========================================
# SIDEBAR FILTERS
# ==========================================
st.sidebar.markdown("---")
st.sidebar.header("🔍 Global Data Filters")

all_businesses = sorted(
    [str(x) for x in df_prep["Business"].dropna().unique() if str(x) != "nan"]
)
all_cities = sorted(
    [str(x) for x in df_prep["City"].dropna().unique() if str(x) != "nan"]
)
all_statuses = sorted(
    [str(x) for x in df_prep["Query Status"].dropna().unique() if str(x) != "nan"]
)

selected_businesses = st.sidebar.multiselect(
    "Filter Business Unit(s):", options=all_businesses
)
selected_cities = st.sidebar.multiselect("Filter City/Cities:", options=all_cities)
selected_statuses = st.sidebar.multiselect(
    "Filter Query Status(es):", options=all_statuses
)

# Apply filters
filtered_df = df_prep.copy()
if selected_businesses:
  filtered_df = filtered_df[filtered_df["Business"].isin(selected_businesses)]
if selected_cities:
  filtered_df = filtered_df[filtered_df["City"].isin(selected_cities)]
if selected_statuses:
  filtered_df = filtered_df[
      filtered_df["Query Status"].isin(selected_statuses)
  ]

# CREATE MULTI-TAB NAVIGATION
tab1, tab2, tab3, tab4, tab5 = st.tabs(
    [
        "🚀 Executive Overview & KPIs",
        "📈 Interactive Pivot Matrix",
        "✉️ SPOC Communication Center",
        "🏢 MLL Pipeline Infographic",
        "🤖 AI Chatbot Assistant"
    ]
)

# ==========================================
# TAB 1: EXECUTIVE OVERVIEW & CHARTS
# ==========================================
with tab1:
  st.subheader("⚡ Real-Time Operational Metrics & Visual Insights")
  st.markdown("High-level telemetry derived strictly from uploaded records.")

  m1, m2, m3, m4 = st.columns(4)
  total_enquiries = len(filtered_df)
  total_area = (
      filtered_df["Closed Area Reqd"].sum()
      if "Closed Area Reqd" in filtered_df.columns
      else 0
  )

  unique_clients_list = (
      sorted(
          [
              str(x)
              for x in filtered_df["Client"].dropna().unique()
              if str(x) != "nan"
          ]
      )
      if "Client" in filtered_df.columns
      else []
  )
  client_count = len(unique_clients_list)
  clients_hover_html = (
      "".join([f"• {c}<br>" for c in unique_clients_list])
      if unique_clients_list
      else "No clients available"
  )
  active_cities = (
      filtered_df["City"].nunique() if "City" in filtered_df.columns else 0
  )

  m1.metric("Total Enquiries", f"{total_enquiries:,}")
  m2.metric("Total Area Required (Sq. Ft.)", f"{total_area:,.2f}")

  with m3:
    st.markdown(
        f"""
        <div class="tooltip-container">
            <div style="background-color: #1e2530; padding: 15px; border-radius: 10px; border: 1px solid #2d3748; box-shadow: 0 4px 6px rgba(0,0,0,0.3);">
                <span style="color: #a0aec0; font-size: 14px;">Unique Clients</span><br>
                <span style="color: #ffffff; font-size: 26px; font-weight: 700;">{client_count}</span>
            </div>
            <div class="tooltip-text">
                <strong>Filtered Unique Clients:</strong><br><br>{clients_hover_html}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

  m4.metric("Active Operating Cities", f"{active_cities}")
  st.markdown("---")

  col_chart1, col_chart2 = st.columns(2)
  with col_chart1:
    st.markdown("### 🏢 Enquiries by Business Unit")
    if "Business" in filtered_df.columns and not filtered_df.empty:
      biz_counts = filtered_df["Business"].value_counts().reset_index()
      biz_counts.columns = ["Business", "Count"]
      st.bar_chart(biz_counts.set_index("Business"), color="#1f77b4", height=350)
    else:
      st.info("No data available for current filters.")

  with col_chart2:
    st.markdown("### 📌 Enquiries by Query Status")
    if "Query Status" in filtered_df.columns and not filtered_df.empty:
      status_counts = filtered_df["Query Status"].value_counts().reset_index()
      status_counts.columns = ["Query Status", "Count"]
      st.bar_chart(
          status_counts.set_index("Query Status"), color="#ff7f0e", height=350
      )
    else:
      st.info("No data available for current filters.")

  st.markdown("---")
  st.markdown("### 🏙️ Geographical Footprint (Area Required by City)")
  if (
      "City" in filtered_df.columns
      and "Closed Area Reqd" in filtered_df.columns
      and not filtered_df.empty
  ):
    city_area = filtered_df.groupby("City")["Closed Area Reqd"].sum().reset_index()
    st.bar_chart(city_area.set_index("City"), color="#2ca02c", height=300)
  else:
    st.info("Insufficient data for geographical breakdown.")

# ==========================================
# TAB 2: INTERACTIVE PIVOT MATRIX
# ==========================================
with tab2:
  st.subheader("📈 Dynamic Business vs. Query Status Pivot Matrix")
  st.markdown(
      "Cross-tabulation matrix mapping **Business Units** against **Query Statuses**."
  )

  if (
      "Business" in filtered_df.columns
      and "Query Status" in filtered_df.columns
      and "Enquiry No." in filtered_df.columns
  ):
    if not filtered_df.empty:
      pivot_df = pd.pivot_table(
          filtered_df,
          index="Business",
          columns="Query Status",
          values="Enquiry No.",
          aggfunc="count",
          fill_value=0,
      )
      pivot_df["Grand Total"] = pivot_df.sum(axis=1)
      summary_row = pivot_df.sum(numeric_only=True)
      pivot_df.loc["Grand Total"] = summary_row

      st.dataframe(pivot_df, use_container_width=True)
      csv_data = pivot_df.to_csv().encode("utf-8")
      st.download_button(
          label="📥 Download Pivot Matrix as CSV",
          data=csv_data,
          file_name="real_estate_pivot_matrix.csv",
          mime="text/csv",
      )
    else:
      st.warning("No records match the active filter criteria.")
  else:
    st.error("Required columns missing from dataset.")

  st.markdown("---")
  st.subheader("📋 Filtered Raw Data View")
  if not filtered_df.empty:
    col_list = list(filtered_df.columns)
    f_col1, f_col2, f_col3 = st.columns(3)
    with f_col1:
      selected_filter_col = st.selectbox(
          "Select Column to Filter/Search:",
          options=col_list,
          key="table_filter_col",
      )
    with f_col2:
      unique_col_vals = sorted([
          str(x)
          for x in filtered_df[selected_filter_col].dropna().unique()
          if str(x) != "nan"
      ])
      selected_col_vals = st.multiselect(
          f"Filter values for '{selected_filter_col}':",
          options=unique_col_vals,
          key="table_filter_vals",
      )
    with f_col3:
      sort_order = st.selectbox(
          "Sort Order:",
          options=["None", "Ascending", "Descending"],
          key="table_sort_order",
      )

    table_view_df = filtered_df.copy()
    if selected_col_vals:
      table_view_df = table_view_df[
          table_view_df[selected_filter_col].astype(str).isin(selected_col_vals)
      ]
    if sort_order == "Ascending":
      table_view_df = table_view_df.sort_values(
          by=selected_filter_col, ascending=True
      )
    elif sort_order == "Descending":
      table_view_df = table_view_df.sort_values(
          by=selected_filter_col, ascending=False
      )

    st.dataframe(table_view_df, use_container_width=True)
  else:
    st.dataframe(filtered_df, use_container_width=True)

# ==========================================
# TAB 3: SPOC COMMUNICATION CENTER
# ==========================================
with tab3:
  st.subheader("✉️ Automated Business SPOC Email & Update Generator")
  if not all_businesses:
    st.warning("No business units identified.")
  else:
    selected_spoc_biz = st.selectbox(
        "Select Business Unit for Operational Briefing:", all_businesses
    )
    if selected_spoc_biz:
      spoc_subset = df_prep[df_prep["Business"] == selected_spoc_biz]
      spoc_count = len(spoc_subset)
      spoc_sqft = (
          spoc_subset["Closed Area Reqd"].sum()
          if "Closed Area Reqd" in spoc_subset.columns
          else 0
      )
      spoc_status_breakdown = (
          spoc_subset["Query Status"].value_counts().to_dict()
          if "Query Status" in spoc_subset.columns
          else {}
      )
      status_lines = (
          "\n".join([
              f"- {status}: {count} queries"
              for status, count in spoc_status_breakdown.items()
          ])
          or "No status data available"
      )

      generated_msg = f"""Dear SPOC,

Here is the latest real estate and operational footprint update for your business unit ({selected_spoc_biz}):

- **Total Active Enquiries:** {spoc_count}
- **Total Required Space:** {spoc_sqft:,.2f} Sq. Ft.

**Status Breakdown:**
{status_lines}

Please let us know if your team requires additional site evaluations or lease structuring adjustments.

Best regards,  
Real Estate & Operations Team"""

      st.text_area(
          "Generated Message (Ready to copy & distribute):",
          value=generated_msg,
          height=240,
      )
      col_s1, col_s2 = st.columns(2)
      col_s1.metric(f"Enquiries for {selected_spoc_biz}", f"{spoc_count}")
      col_s2.metric(
          f"Total Space Requested ({selected_spoc_biz})",
          f"{spoc_sqft:,.2f} Sq. Ft.",
      )

# ==========================================
# TAB 4: MLL PIPELINE INFOGRAPHIC
# ==========================================
with tab4:
  st.markdown(
      """
        <div class="info-banner">
            <h1 style="margin:0; font-size: 32px;">MLL's Real Estate & Workplace Pipeline</h1>
        </div>
        """,
      unsafe_allow_html=True,
  )

  if filtered_df.empty:
    st.warning("No data matches the current filters to generate the infographic.")
  else:
    total_queries_info = len(filtered_df)
    total_sqft_info = (
        filtered_df["Closed Area Reqd"].sum()
        if "Closed Area Reqd" in filtered_df.columns
        else 0
    )
    total_cities_info = (
        filtered_df["City"].nunique() if "City" in filtered_df.columns else 0
    )
    total_biz_info = (
        filtered_df["Business"].nunique()
        if "Business" in filtered_df.columns
        else 0
    )

    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
      st.markdown(
          f'<div class="info-kpi-box"><div class="info-kpi-value">{total_queries_info}</div><div'
          ' class="info-kpi-label">Total Queries</div></div>',
          unsafe_allow_html=True,
      )
    with k2:
      st.markdown(
          f'<div class="info-kpi-box"><div'
          f' class="info-kpi-value">{total_sqft_info / 1_000_000:.2f}'
          ' Mn</div><div class="info-kpi-label">Sq. Ft. Reqd</div></div>',
          unsafe_allow_html=True,
      )
    with k3:
      st.markdown(
          f'<div class="info-kpi-box"><div'
          f' class="info-kpi-value">{total_biz_info}</div><div'
          ' class="info-kpi-label">Businesses</div></div>',
          unsafe_allow_html=True,
      )
    with k4:
      st.markdown(
          f'<div class="info-kpi-box"><div'
          f' class="info-kpi-value">{total_cities_info}</div><div'
          ' class="info-kpi-label">Locations</div></div>',
          unsafe_allow_html=True,
      )

    st.markdown("<br>", unsafe_allow_html=True)
    m_col1, m_col2 = st.columns([1, 1])

    with m_col1:
      st.markdown(
          '<div class="info-section-header">PORTFOLIO MIX (BY BUSINESS SQ.'
          " FT.)</div>",
          unsafe_allow_html=True,
      )
      if (
          "Business" in filtered_df.columns
          and "Closed Area Reqd" in filtered_df.columns
      ):
        biz_agg = (
            filtered_df.groupby("Business")
            .agg(
                Area=("Closed Area Reqd", "sum"),
                Queries=("Enquiry No.", "count"),
            )
            .reset_index()
        )
        fig = px.pie(
            biz_agg,
            names="Business",
            values="Area",
            hole=0.65,
            color_discrete_sequence=px.colors.qualitative.Set1,
            custom_data=["Queries"],
        )
        fig.update_traces(
            hovertemplate=(
                "<b>%{label}</b><br>Area: %{value:,.0f}"
                " Sq.Ft.<br>Queries: %{customdata[0]}<extra></extra>"
            )
        )
        fig.update_layout(
            showlegend=True,
            margin=dict(t=10, b=10, l=10, r=10),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color="white"),
        )
        st.plotly_chart(fig, use_container_width=True)

    with m_col2:
      st.markdown(
          '<div class="info-section-header">WORKFLOW / REGISTRATION'
          " STATUS</div>",
          unsafe_allow_html=True,
      )
      if "Query Status" in filtered_df.columns:
        status_summary = (
            filtered_df["Query Status"].value_counts().reset_index()
        )
        status_summary.columns = ["Status", "Count"]
        fig3 = px.pie(
            status_summary,
            names="Status",
            values="Count",
            hole=0.5,
            color_discrete_sequence=px.colors.qualitative.Pastel,
        )
        fig3.update_traces(textposition="inside", textinfo="percent+label")
        fig3.update_layout(
            showlegend=False,
            margin=dict(t=10, b=10, l=10, r=10),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color="white"),
        )
        st.plotly_chart(fig3, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        '<div class="info-section-header">PAN-INDIA PRESENCE & BUSINESS'
        " FOOTPRINT (AREA SQ. FT.)</div>",
        unsafe_allow_html=True,
    )
    if (
        "City" in filtered_df.columns
        and "Business" in filtered_df.columns
        and "Closed Area Reqd" in filtered_df.columns
    ):
      fig_tree = px.treemap(
          filtered_df,
          path=[px.Constant("India Operations"), "Business", "City"],
          values="Closed Area Reqd",
          color="Business",
          color_discrete_sequence=px.colors.qualitative.Set1,
      )
      fig_tree.update_traces(
          textinfo="label+value+percent parent",
          hovertemplate=(
              "<b>%{label}</b><br>Required Area: %{value:,.0f}"
              " Sq.Ft.<extra></extra>"
          ),
      )
      fig_tree.update_layout(
          margin=dict(t=20, l=10, r=10, b=10),
          paper_bgcolor="rgba(0,0,0,0)",
          plot_bgcolor="rgba(0,0,0,0)",
          font=dict(color="white", size=14),
      )
      st.plotly_chart(fig_tree, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)
    b_col1, b_col2 = st.columns([2, 1])

    with b_col1:
      st.markdown(
          '<div class="info-section-header">BUSINESS-WISE MIX (BY NUMBER OF'
          " QUERIES & AREA)</div>",
          unsafe_allow_html=True,
      )
      if (
          "Business" in filtered_df.columns
          and "Closed Area Reqd" in filtered_df.columns
      ):
        biz_mix_df = (
            filtered_df.groupby("Business")
            .agg(
                Queries=("Enquiry No.", "count"),
                Area=("Closed Area Reqd", "sum"),
            )
            .reset_index()
        )
        total_area_all = biz_mix_df["Area"].sum()
        biz_mix_df["% of Total Area"] = (
            (biz_mix_df["Area"] / total_area_all * 100)
            if total_area_all > 0
            else 0
        )
        biz_mix_df = biz_mix_df.sort_values(by="Area", ascending=False)
        st.dataframe(
            biz_mix_df,
            column_config={
                "Business": st.column_config.TextColumn(
                    "Business", width="medium"
                ),
                "Queries": st.column_config.NumberColumn("# of Queries"),
                "Area": st.column_config.NumberColumn(
                    "Area (Sq. Ft.)", format="%d"
                ),
                "% of Total Area": st.column_config.ProgressColumn(
                    "% of Total Area",
                    help="Percentage contribution to total required footprint",
                    format="%.1f%%",
                    min_value=0,
                    max_value=100,
                ),
            },
            hide_index=True,
            use_container_width=True,
        )

    with b_col2:
      st.markdown(
          '<div class="info-section-header">TOP 7 CUSTOMERS</div>',
          unsafe_allow_html=True,
      )
      if "Client" in filtered_df.columns:
        top_clients = (
            filtered_df["Client"].value_counts().head(7).reset_index()
        )
        top_clients.columns = ["Customer Name", "Active Queries"]
        st.dataframe(
            top_clients,
            column_config={
                "Customer Name": st.column_config.TextColumn("Customer"),
                "Active Queries": st.column_config.NumberColumn(
                    "Queries", width="small"
                ),
            },
            hide_index=True,
            use_container_width=True,
        )

# ==========================================
# TAB 5: AI CHATBOT
# ==========================================
with tab5: 
    st.subheader("🤖 Universal Excel Assistant") 
    st.markdown("Ask natural language questions across all uploaded files.")

    if "messages" not in st.session_state:
        st.session_state.messages = []

    for msg in st.session_state.messages:
        st.chat_message(msg["role"]).write(msg["content"])

    if prompt := st.chat_input("E.g., 'What is the total area for L&T?'"):
        st.session_state.messages.append({"role": "user", "content": prompt})
        st.chat_message("user").write(prompt)

        if not GEMINI_API_KEY:
            st.error("Please configure your GOOGLE_API_KEY in the `.streamlit/secrets.toml` file to use the AI Agent.")
        else:
            os.environ["GOOGLE_API_KEY"] = GEMINI_API_KEY
            with st.chat_message("assistant"):
                with st.spinner("Analyzing spreadsheets..."):
                    try:
                        llm = ChatGoogleGenerativeAI(model="gemini-3.5-flash-lite", temperature=0)
                        agent = create_pandas_dataframe_agent(
                            llm, list(dfs.values()), agent_type="tool-calling", allow_dangerous_code=True, max_iterations=13, verbose=True
                        )
                        
                        safe_prompt = f"{prompt}\n\nInstructions: Provide the final answer clearly in simple terms. Do not output raw python code."
                        response = agent.invoke({"input": safe_prompt})
                        
                        output_data = response["output"]
                        final_text = output_data[0].get("text", "") if isinstance(output_data, list) else str(output_data)
                        st.write(final_text)
                    except Exception as e:
                        st.error(f"Agent Error: {e}")
