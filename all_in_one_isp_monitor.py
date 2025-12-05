import subprocess
import platform
from datetime import datetime
import requests
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import re
import os
import pandas as pd
import streamlit as st
import plotly.express as px

# =========================
# CONFIG
# =========================
LATENCY_WARNING_THRESHOLD = 100
REPORT_FOLDER = "ping_reports"
IP_LIST_PATH = "ip_list.txt"

if not os.path.exists(REPORT_FOLDER):
    os.makedirs(REPORT_FOLDER)

# =========================
# PING + REPORT FUNCTIONS
# =========================
param = "-n" if platform.system().lower() == "windows" else "-c"

def get_isp(ip):
    try:
        url = f"http://ip-api.com/json/{ip}"
        response = requests.get(url, timeout=5).json()
        return response.get("isp", "Unknown")
    except:
        return "Lookup Failed"

def ping_ip(ip):
    try:
        command = ["ping", param, "2", ip]
        output = subprocess.check_output(command, universal_newlines=True)

        latency = "N/A"
        for line in output.splitlines():
            match = re.search(r"time[=<]([\d\.]+)", line)
            if match:
                latency = match.group(1)
                break

        return "PASS", latency
    except:
        return "FAIL", "N/A"

def generate_report():
    with open(IP_LIST_PATH, "r") as file:
        ip_list = [ip.strip() for ip in file if ip.strip()]

    wb = Workbook()
    ws = wb.active
    ws.title = "Ping Report"
    ws.append(["IP Address", "ISP Name", "Status", "Latency (ms)", "Timestamp"])

    red_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")

    pass_count = 0
    fail_count = 0
    high_latency_count = 0
    any_failure = False

    for ip in ip_list:
        status, latency = ping_ip(ip)
        isp = get_isp(ip)
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        ws.append([ip, isp, status, latency, timestamp])
        row = ws.max_row

        if status == "PASS":
            pass_count += 1
        else:
            fail_count += 1
            any_failure = True
            for cell in ws[row]:
                cell.fill = red_fill

        if latency != "N/A":
            try:
                if float(latency) > LATENCY_WARNING_THRESHOLD:
                    high_latency_count += 1
                    ws.cell(row=row, column=4).fill = yellow_fill
            except:
                pass

    for column_cells in ws.columns:
        max_length = max(len(str(cell.value)) for cell in column_cells)
        ws.column_dimensions[column_cells[0].column_letter].width = max_length + 3

    summary_ws = wb.create_sheet(title="Summary")
    summary_ws.append(["Metric", "Value"])
    summary_ws.append(["Total IPs Checked", len(ip_list)])
    summary_ws.append(["Total PASS", pass_count])
    summary_ws.append(["Total FAIL", fail_count])
    summary_ws.append(["High Latency (>100ms)", high_latency_count])
    summary_ws.append(["Report Generated At", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])

    for column_cells in summary_ws.columns:
        max_length = max(len(str(cell.value)) for cell in column_cells)
        summary_ws.column_dimensions[column_cells[0].column_letter].width = max_length + 3

    filename = os.path.join(
        REPORT_FOLDER,
        f"ping_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )
    wb.save(filename)

    return filename, any_failure

# =========================
# STREAMLIT DASHBOARD UI
# =========================
st.set_page_config(page_title="ISP Monitoring Dashboard", layout="wide")

# Auto-refresh every 60 seconds
st.experimental_set_query_params(auto_refresh=int(datetime.now().timestamp()))
st.markdown("<meta http-equiv='refresh' content='60'>", unsafe_allow_html=True)

st.title("🌐 All-in-One Multi-ISP Monitoring Dashboard")

if st.button("▶ Run Ping Test & Generate New Report"):
    file_created, any_failure = generate_report()
    st.success(f"New Report Created: {file_created}")
    if any_failure:
        st.warning("⚠️ Some IPs have FAILED the ping test! Check the report for details.")

# Load latest report
files = sorted(
    [os.path.join(REPORT_FOLDER, f) for f in os.listdir(REPORT_FOLDER) if f.endswith(".xlsx")],
    reverse=True
)

if not files:
    st.warning("No reports found yet. Click the button above to generate one.")
    st.stop()

latest_file = files[0]
latest_df = pd.read_excel(latest_file, sheet_name="Ping Report")
latest_summary = pd.read_excel(latest_file, sheet_name="Summary")

# =========================
# Summary Metrics
# =========================
col1, col2, col3, col4 = st.columns(4)
col1.metric("Total IPs", int(latest_summary.iloc[0,1]))
col2.metric("PASS", int(latest_summary.iloc[1,1]))
col3.metric("FAIL", int(latest_summary.iloc[2,1]))
col4.metric("High Latency", int(latest_summary.iloc[3,1]))

# =========================
# Filter by ISP
# =========================
st.subheader("Filter by ISP")
isp_list = ["All"] + sorted(latest_df["ISP Name"].astype(str).unique().tolist())
selected_isp = st.selectbox("Select ISP", isp_list)

df_live = latest_df.copy()
if selected_isp != "All":
    df_live = df_live[df_live["ISP Name"] == selected_isp]

st.subheader("Live Status Table")
st.dataframe(df_live, use_container_width=True)

st.subheader("Failed Links (Live)")
st.dataframe(df_live[df_live["Status"] == "FAIL"], use_container_width=True)

# =========================
# ISP-wise Live Health with PASS/FAIL/High Latency
# =========================
st.subheader("ISP-wise Live Health (PASS / FAIL / High Latency)")

df_live["High Latency"] = df_live["Latency (ms)"].apply(
    lambda x: 1 if x != "N/A" and float(x) > LATENCY_WARNING_THRESHOLD else 0
)

isp_summary = df_live.groupby("ISP Name").agg({
    "PASS": lambda x: sum(x=="PASS"),
    "FAIL": lambda x: sum(x=="FAIL"),
    "High Latency": "sum"
}).reset_index()

isp_summary_melt = isp_summary.melt(
    id_vars="ISP Name",
    value_vars=["PASS", "FAIL", "High Latency"],
    var_name="Status",
    value_name="Count"
)

color_map = {"PASS": "green", "FAIL": "red", "High Latency": "yellow"}

fig = px.bar(
    isp_summary_melt,
    x="ISP Name",
    y="Count",
    color="Status",
    color_discrete_map=color_map,
    barmode="stack",
    text="Count"
)
fig.update_layout(yaxis_title="Number of IPs", xaxis_title="ISP Name")
st.plotly_chart(fig, use_container_width=True)

# =========================
# Historical Trend
# =========================
st.subheader("📈 Historical PASS / FAIL Trend")
history_data = []
for file in files[::-1]:
    try:
        s = pd.read_excel(file, sheet_name="Summary")
        history_data.append({
            "Timestamp": s.iloc[4,1],
            "PASS": int(s.iloc[1,1]),
            "FAIL": int(s.iloc[2,1]),
            "High Latency": int(s.iloc[3,1])
        })
    except:
        pass

history_df = pd.DataFrame(history_data)
history_df["Timestamp"] = pd.to_datetime(history_df["Timestamp"])
history_df = history_df.sort_values("Timestamp")

st.line_chart(history_df.set_index("Timestamp")[["PASS", "FAIL"]])
st.subheader("🟡 High Latency Trend")
st.line_chart(history_df.set_index("Timestamp")[["High Latency"]])

st.caption(f"Last Updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
