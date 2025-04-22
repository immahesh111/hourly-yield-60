import streamlit as st
from pymongo import MongoClient
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime, timedelta
import pytz
import certifi

# MongoDB connection
MONGO_URI = st.secrets["mongo"]["uri"]  # Store in secrets.toml
client = MongoClient(MONGO_URI, tlsCAFile=certifi.where())
db = client["yield_dashboard"]
collection = db["yield_data"]

# Streamlit app configuration
st.set_page_config(layout="wide", page_title="Real-Time Yield Dashboard", initial_sidebar_state="expanded")

# Custom CSS for modern UI/UX
st.markdown("""
<style>
/* Main header styling */
.header {
    text-align: center;
    background: linear-gradient(90deg, #0062cc, #007bff);
    color: white;
    padding: 0.5rem;
    border-radius: 10px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    margin-bottom: 1rem;
}
.header h1 {
    margin: 0;
    font-size: 1.5rem;
    font-weight: 600;
}
.date-time {
    font-size: 1rem;
    margin-top: 0.3rem;
}
.date-time span {
    background-color: rgba(255,255,255,0.2);
    padding: 0.2rem 0.5rem;
    border-radius: 5px;
    margin: 0 0.5rem;
}

/* Card and subheader styling */
.card {
    background: white;
    border-radius: 10px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    padding: 1rem;
    margin-bottom: 1rem;
}
.card-header {
    text-align: center;
    font-size: 1.2rem;
    font-weight: bold;
    padding: 0.5rem;
    border-radius: 5px;
    margin-bottom: 0.5rem;
}
.card-header.hourly-yield {
    background-color: #cce5ff;
    color: #004085;
}
.card-header.overall-yield {
    background-color: #d4edda;
    color: #155724;
}
.card-header.yield-trend {
    background-color: #fff3cd;
    color: #856404;
}
.card-header.top-issues {
    background-color: #f8d7da;
    color: #721c24;
}
.card-header.error-table {
    background-color: #e2e3e5;
    color: #383d41;
}
.card-header.error-trend {
    background-color: #f5c6cb;
    color: #491217;
}

/* Sidebar styling */
.sidebar .sidebar-content {
    background-color: #f8f9fa;
}
</style>
""", unsafe_allow_html=True)

# Current date and time
tz = pytz.timezone('Asia/Kolkata')
today = datetime.now(tz)

# Centered header with date and time
st.markdown(f"""
<div class='header'>
    <h1>Real-Time Yield Dashboard</h1>
    <div class='date-time'>
        <span>{today.strftime('%Y-%m-%d')}</span>
        <span>{today.strftime('%H:%M:%S')}</span>
    </div>
</div>
""", unsafe_allow_html=True)

# Sidebar controls
with st.sidebar:
    st.header("Controls")
    selected_line = st.radio("Select Line", ["Line 01", "Line 03", "Line 08", "Line 09", "Line 10", "Line 12", "Line 15"])
    selected_date = st.date_input("Select Date", value=today.date())
# Main content
if selected_line:
    start_of_day = datetime.combine(selected_date, datetime.min.time()).replace(tzinfo=tz)
    end_of_day = start_of_day + timedelta(days=1)
    query = {
        "line": selected_line,
        "start_time": {"$gte": start_of_day, "$lt": end_of_day}
    }
    documents = list(collection.find(query).sort("start_time", 1))

    if documents:
        # Data preparation
        time_slots = [doc["time_slot"] for doc in documents]
        yields = [doc["yield"] * 100 for doc in documents]
        inputs = [doc["input"] for doc in documents]
        total_failures = [sum(rc["count"] for rc in doc["root_causes"]) + doc["other_failures"]["count"] for doc in documents]

        # Reason codes table
        reason_codes = set()
        for doc in documents:
            for rc in doc["root_causes"]:
                reason_codes.add(rc["root_cause"])
        reason_codes.add("Other Failures")
        reason_codes = sorted(reason_codes)

        cols = pd.MultiIndex.from_product([time_slots, ["Failure Count", "Failure Rate (%)"]], names=["Time Slot", "Metric"])
        df_reasons = pd.DataFrame(index=reason_codes, columns=cols)
        df_reasons.loc[:, (slice(None), "Failure Count")] = 0

        for doc in documents:
            ts = doc["time_slot"]
            for rc in doc["root_causes"]:
                df_reasons.loc[rc["root_cause"], (ts, "Failure Count")] = rc["count"]
                df_reasons.loc[rc["root_cause"], (ts, "Failure Rate (%)")] = f"{rc['rate']*100:.2f}%"
            df_reasons.loc["Other Failures", (ts, "Failure Count")] = doc["other_failures"]["count"]
            df_reasons.loc["Other Failures", (ts, "Failure Rate (%)")] = f"{doc['other_failures']['rate']*100:.2f}%"

        for ts in time_slots:
            df_reasons.loc[:, (ts, "Failure Rate (%)")] = df_reasons.loc[:, (ts, "Failure Rate (%)")].fillna('0.00%')

        def color_scale(series):
            vals = pd.to_numeric(series.str.replace('%',''), errors='coerce')
            colors = []
            for v in vals:
                if pd.isna(v): colors.append('background-color: white')
                elif v == 0: colors.append('background-color: #28a745')
                elif v > 1: colors.append('background-color: #dc3545')
                else: colors.append('background-color: #ffc107')
            return colors

        styled_df = df_reasons.style.apply(lambda col: ['']*len(col) if col.name[1] == 'Failure Count' else color_scale(col), axis=0)

        # Aggregate issues
        hourly_issues = {}
        for doc in documents:
            for rc in doc["root_causes"]:
                hourly_issues[rc["root_cause"]] = hourly_issues.get(rc["root_cause"], 0) + rc["count"]
            hourly_issues["Other Failures"] = hourly_issues.get("Other Failures", 0) + doc["other_failures"]["count"]
        top_issues = sorted(hourly_issues.items(), key=lambda x: x[1], reverse=True)[:3]

        # Gauges with input and failure text
        def create_gauge(val, title, input_val, failure_val):
            color = 'red' if val <= 95 else 'orange' if val <= 98 else 'green'
            fig = go.Figure(go.Indicator(
                mode="gauge+number",
                value=val,
                title={'text': title, 'font': {'size': 18}},
                gauge={
                    'axis': {'range': [0, 100]},
                    'bar': {'color': color},
                    'steps': [
                        {'range': [0, 95], 'color': 'rgba(255,0,0,0.2)'},
                        {'range': [95, 98], 'color': 'rgba(255,165,0,0.2)'},
                        {'range': [98, 100], 'color': 'rgba(0,128,0,0.2)'}
                    ]
                }
            ))
            fig.add_annotation(
                x=0.5, y=-0.1, xref="paper", yref="paper",
                text=f"Input: {input_val}<br>Failures: {failure_val}",
                showarrow=False, font=dict(size=12), align="center"
            )
            fig.update_layout(margin=dict(l=20, r=20, t=30, b=50), height=250, paper_bgcolor='white')
            return fig

        hourly_yield = yields[-1]
        overall_yield = np.mean(yields)
        hourly_input = inputs[-1]
        hourly_failures = total_failures[-1]
        overall_input = int(sum(inputs))  # Corrected to sum of inputs
        overall_failures = int(sum(total_failures))  # Corrected to sum of failures
        hourly_gauge = create_gauge(hourly_yield, "Hourly Yield", hourly_input, hourly_failures)
        overall_gauge = create_gauge(overall_yield, "Overall Yield", overall_input, overall_failures)

        # Trend charts with hover data
        fig_hourly = go.Figure(data=[
            go.Bar(
                x=time_slots,
                y=yields,
                marker_color=['#dc3545' if y <= 95 else '#ffc107' if y <= 98 else '#28a745' for y in yields],
                text=[f"{y:.2f}%" for y in yields],
                textposition='auto',
                customdata=np.stack((inputs, total_failures), axis=-1),
                hovertemplate='Time Slot: %{x}<br>Yield: %{y:.2f}%<br>Input: %{customdata[0]}<br>Failures: %{customdata[1]}'
            )
        ])
        fig_hourly.update_layout(
            margin=dict(l=20, r=20, t=20, b=40),
            xaxis_title="Time Slot",
            yaxis_title="Yield (%)",
            yaxis_range=[0, 100],
            height=300
        )

        daily = {}
        daily_inputs = {}
        daily_failures = {}
        for doc in documents:
            d = doc['start_time'].strftime('%Y-%m-%d')
            daily.setdefault(d, []).append(doc['yield'] * 100)
            daily_inputs.setdefault(d, []).append(doc['input'])
            daily_failures.setdefault(d, []).append(sum(rc['count'] for rc in doc['root_causes']) + doc['other_failures']['count'])
        daily_avg = {d: np.mean(vals) for d, vals in daily.items()}
        daily_avg_inputs = {d: int(sum(vals)) for d, vals in daily_inputs.items()}  # Sum for daily inputs
        daily_avg_failures = {d: int(sum(vals)) for d, vals in daily_failures.items()}  # Sum for daily failures
        fig_daily = go.Figure(data=[
            go.Scatter(
                x=list(daily_avg.keys()),
                y=list(daily_avg.values()),
                mode='lines+markers',
                customdata=np.stack((list(daily_avg_inputs.values()), list(daily_avg_failures.values())), axis=-1),
                hovertemplate='Date: %{x}<br>Yield: %{y:.2f}%<br>Input: %{customdata[0]}<br>Failures: %{customdata[1]}'
            )
        ])
        fig_daily.update_layout(
            margin=dict(l=20, r=20, t=20, b=40),
            xaxis_title='Date',
            yaxis_title='Avg Yield (%)',
            height=300
        )

        # Error trend charts
        fig_error = go.Figure(data=[
            go.Bar(
                x=time_slots,
                y=total_failures,
                marker_color='#dc3545',
                text=total_failures,
                textposition='auto'
            )
        ])
        fig_error.update_layout(
            margin=dict(l=20, r=20, t=20, b=40),
            xaxis_title='Time Slot',
            yaxis_title='Total Failures',
            height=300
        )

        daily_err = {}
        for doc in documents:
            d = doc['start_time'].strftime('%Y-%m-%d')
            daily_err[d] = daily_err.get(d, 0) + sum(rc['count'] for rc in doc['root_causes']) + doc['other_failures']['count']
        fig_error_daily = go.Figure(data=[
            go.Scatter(
                x=list(daily_err.keys()),
                y=list(daily_err.values()),
                mode='lines+markers'
            )
        ])
        fig_error_daily.update_layout(
            margin=dict(l=20, r=20, t=20, b=40),
            xaxis_title='Date',
            yaxis_title='Total Failures',
            height=300
        )

        # Layout: Top section
        with st.container():
            col1, col2, col3 = st.columns([1, 2, 1])
            with col1:
                st.markdown("<div class='card'><div class='card-header hourly-yield'>Hourly Yield</div></div>", unsafe_allow_html=True)
                st.plotly_chart(hourly_gauge, use_container_width=True, config={'displayModeBar': False})
                st.markdown("<div class='card'><div class='card-header overall-yield'>Overall Yield</div></div>", unsafe_allow_html=True)
                st.plotly_chart(overall_gauge, use_container_width=True, config={'displayModeBar': False})
            with col2:
                st.markdown("<div class='card'><div class='card-header yield-trend'>Yield Trend</div></div>", unsafe_allow_html=True)
                mode = st.radio("", ["Hourly", "Daily"], horizontal=True, key='yield_trend')
                st.plotly_chart(fig_hourly if mode == 'Hourly' else fig_daily, use_container_width=True, config={'displayModeBar': False})
            with col3:
                st.markdown("<div class='card'><div class='card-header top-issues'>Top 3 Issues</div></div>", unsafe_allow_html=True)
                if top_issues:
                    for i, (issue, cnt) in enumerate(top_issues):
                        st.markdown(f"<div style='text-align: center;'>{i+1}. {issue} — {cnt}</div>", unsafe_allow_html=True)
                    pie = go.Figure(data=[
                        go.Pie(labels=[i for i, _ in top_issues], values=[c for _, c in top_issues], hole=0.3)
                    ])
                    pie.update_layout(margin=dict(l=20, r=20, t=20, b=20), height=220)
                    st.plotly_chart(pie, use_container_width=True, config={'displayModeBar': False})
                else:
                    st.write("No issues found.")

        # Layout: Bottom section
        with st.container():
            b1, b2 = st.columns([3, 2])
            with b1:
                st.markdown("<div class='card'><div class='card-header error-table'>Hourly Error Counts and Percentages</div></div>", unsafe_allow_html=True)
                st.dataframe(styled_df, height=350)
            with b2:
                st.markdown("<div class='card'><div class='card-header error-trend'>Error Count Trend</div></div>", unsafe_allow_html=True)
                err_mode = st.radio("", ["Hourly", "Daily"], horizontal=True, key='error_trend')
                st.plotly_chart(fig_error if err_mode == 'Hourly' else fig_error_daily, use_container_width=True, config={'displayModeBar': False})
    else:
        st.error("No data available for the selected line and date.")
else:
    st.info("Please select a line to view the data.")

# Auto-refresh every 10 minutes
st.markdown("""
<script>
    setTimeout(function() {
        location.reload();
    }, 600000); // 10 minutes = 600,000 milliseconds
</script>
""", unsafe_allow_html=True)

# Refresh note
st.markdown("<div style='text-align: center; color: grey; font-size: 0.9rem;'>Data refreshes every 10 minutes</div>", unsafe_allow_html=True)