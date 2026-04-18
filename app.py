"""
PMI Auto-Dispatcher — Streamlit App
TXANT San Antonio Hub
"""

import io
import streamlit as st
import pandas as pd
from datetime import datetime

# Must be first Streamlit call
st.set_page_config(
    page_title="PMI Auto-Dispatcher",
    page_icon="🔧",
    layout="wide"
)

import dispatcher_core as dc

# ── Bundled file paths ────────────────────────────────────────────────────────
MASTER_PLAN_PATH = "PMI_12MoCal_master.xlsx"

# ── Styling ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1F4E79, #2E75B6);
        padding: 20px 30px;
        border-radius: 10px;
        margin-bottom: 25px;
        color: white;
    }
    .main-header h1 { color: white; margin: 0; font-size: 28px; }
    .main-header p  { color: #BDD7EE; margin: 5px 0 0 0; font-size: 14px; }
    .upload-card {
        background: #F8F9FA;
        border: 2px dashed #BDD7EE;
        border-radius: 8px;
        padding: 15px;
        margin-bottom: 10px;
    }
    .stat-box {
        background: #EBF3FB;
        border-left: 4px solid #2E75B6;
        padding: 12px 16px;
        border-radius: 4px;
        margin: 6px 0;
    }
    .flag-box {
        background: #FFF8E1;
        border-left: 4px solid #FFC107;
        padding: 12px 16px;
        border-radius: 4px;
        margin: 6px 0;
    }
    .success-box {
        background: #E8F5E9;
        border-left: 4px solid #4CAF50;
        padding: 12px 16px;
        border-radius: 4px;
    }
    .error-box {
        background: #FFEBEE;
        border-left: 4px solid #F44336;
        padding: 12px 16px;
        border-radius: 4px;
    }
    div[data-testid="stButton"] button {
        background: #1F4E79;
        color: white;
        border: none;
        padding: 10px 30px;
        font-size: 16px;
        font-weight: bold;
        border-radius: 6px;
        width: 100%;
    }
    div[data-testid="stButton"] button:hover {
        background: #2E75B6;
    }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>🔧 PMI Auto-Dispatcher</h1>
    <p>TXANT — San Antonio Hub | Automated PM Assignment Tool</p>
</div>
""", unsafe_allow_html=True)

# ── Layout ────────────────────────────────────────────────────────────────────
col_left, col_right = st.columns([1, 1.6])

with col_left:
    st.subheader("📋 Inputs")

    # Month selector
    month_names = ["January", "February", "March", "April", "May", "June",
                   "July", "August", "September", "October", "November", "December"]
    current_month = datetime.now().month
    current_year  = datetime.now().year

    col_m, col_y = st.columns(2)
    with col_m:
        selected_month_name = st.selectbox("Month", month_names, index=current_month - 1)
    with col_y:
        selected_year = st.number_input("Year", min_value=2024, max_value=2030,
                                        value=current_year, step=1)

    selected_month = month_names.index(selected_month_name) + 1
    quarter = dc.quarter_from_month(selected_month)
    st.caption(f"📅 Quarter: **Q{quarter}** — rotation auto-applied")

    st.markdown("---")

    # File uploads
    st.markdown("**① PM Compliance Export** *(monthly Maximo pull)*")
    compliance_file = st.file_uploader(
        "compliance", type=["xlsx"], label_visibility="collapsed",
        key="compliance"
    )
    if compliance_file:
        st.success(f"✓ {compliance_file.name}")

    st.markdown("**② Quarterly Rotation File**")
    st.caption("Update when quarter changes — shift data down by 1 row")
    rotation_file = st.file_uploader(
        "rotation", type=["xlsx"], label_visibility="collapsed",
        key="rotation"
    )
    if rotation_file:
        st.success(f"✓ {rotation_file.name}")

    st.markdown("**③ Mechanic Schedule File**")
    st.caption("Update after union bid — edit shift times and sort overlaps")
    schedule_file = st.file_uploader(
        "schedule", type=["xlsx"], label_visibility="collapsed",
        key="schedule"
    )
    if schedule_file:
        st.success(f"✓ {schedule_file.name}")

    st.markdown("---")

    all_uploaded = compliance_file and rotation_file and schedule_file
    generate_btn = st.button("⚡ Generate Dispatch", disabled=not all_uploaded)

    if not all_uploaded:
        missing = []
        if not compliance_file: missing.append("PM Compliance")
        if not rotation_file:   missing.append("Quarterly Rotation")
        if not schedule_file:   missing.append("Mechanic Schedule")
        if missing:
            st.caption(f"⚠️ Still needed: {', '.join(missing)}")

with col_right:
    st.subheader("📊 Results")

    if not all_uploaded and not generate_btn:
        st.markdown("""
        <div class="stat-box">
            <b>How to use this tool:</b><br><br>
            1. Select the <b>month and year</b> you are dispatching for<br>
            2. Upload the <b>PM Compliance</b> export from Maximo<br>
            3. Upload the current <b>Quarterly Rotation</b> file<br>
            4. Upload the <b>Mechanic Schedule</b> file<br>
            5. Click <b>Generate Dispatch</b><br>
            6. Download the finished Excel file<br><br>
            <small>⚠️ Pull compliance report between the 20th–25th of the month.<br>
            Any PMs with due dates outside the selected month are automatically filtered out.</small>
        </div>
        """, unsafe_allow_html=True)

    if generate_btn and all_uploaded:
        with st.spinner("Running dispatcher..."):
            try:
                # Validate rotation
                rotation_file.seek(0)
                rotation = dc.parse_rotation(rotation_file)
                rotation_errors = dc.validate_rotation(rotation)

                if rotation_errors:
                    st.markdown('<div class="error-box"><b>❌ Rotation file errors:</b><br>' +
                                '<br>'.join(rotation_errors) + '</div>', unsafe_allow_html=True)
                    st.stop()

                # Validate schedule
                schedule_file.seek(0)
                schedule = dc.parse_schedule(schedule_file)
                schedule_errors = dc.validate_schedule(schedule)

                if schedule_errors:
                    st.markdown('<div class="error-box"><b>❌ Schedule file errors:</b><br>' +
                                '<br>'.join(schedule_errors) + '</div>', unsafe_allow_html=True)
                    st.stop()

                # Load compliance
                compliance_file.seek(0)
                df_raw = dc.load_maximo_export(compliance_file)
                df_month = dc.filter_by_month(df_raw, int(selected_year), selected_month)

                if len(df_month) == 0:
                    st.markdown(f'<div class="error-box">❌ No PMs found with due date in '
                                f'{selected_month_name} {selected_year}. '
                                f'Check that the compliance file covers the right period.</div>',
                                unsafe_allow_html=True)
                    st.stop()

                # Run dispatch
                master_hours = dc.load_master_plan(MASTER_PLAN_PATH, selected_month - 1)

                rotation_file.seek(0)
                schedule_file.seek(0)
                blocks, hour_summary, flags = dc.dispatch(
                    df_month, master_hours, rotation, schedule
                )

                # Write output to bytes buffer
                output_buf = io.BytesIO()
                dc.write_excel(blocks, hour_summary, flags, "/tmp/dispatch_output.xlsx")
                with open("/tmp/dispatch_output.xlsx", "rb") as f:
                    output_bytes = f.read()

                # ── Results display ───────────────────────────────────────────
                assigned_count = sum(len(v) for k, v in blocks.items() if k != "Unassigned")
                unassigned_count = len(blocks.get("Unassigned", pd.DataFrame()))
                filtered_out = len(df_raw) - len(df_month)

                st.markdown(f"""
                <div class="success-box">
                    <b>✅ Dispatch complete — {selected_month_name} {selected_year} (Q{quarter})</b><br>
                    {len(df_month)} PMs processed &nbsp;|&nbsp;
                    {assigned_count} assigned &nbsp;|&nbsp;
                    {unassigned_count} unassigned &nbsp;|&nbsp;
                    {filtered_out} filtered (wrong month)
                </div>
                """, unsafe_allow_html=True)

                # Hour summary table
                st.markdown("**Hour Distribution:**")
                summary_data = []
                for mech in dc.MECHANIC_ORDER:
                    hrs = round(hour_summary.get(mech, 0), 1)
                    cnt = len(blocks.get(mech, pd.DataFrame()))
                    summary_data.append({"Mechanic": mech, "Hours": hrs, "PM Count": cnt})

                summary_df = pd.DataFrame(summary_data)
                max_hrs = summary_df["Hours"].max() if not summary_df.empty else 1

                def style_hours(val):
                    ratio = val / max_hrs if max_hrs > 0 else 0
                    if ratio > 0.9:
                        return "background-color: #FFCDD2"
                    elif ratio > 0.75:
                        return "background-color: #FFF9C4"
                    return ""

                styled = summary_df.style.applymap(style_hours, subset=["Hours"])
                st.dataframe(styled, use_container_width=True, hide_index=True)

                # Flags
                if flags:
                    st.markdown(f"""
                    <div class="flag-box">
                        <b>⚠️ {len(flags)} item(s) flagged for review</b> — see "Review Required"
                        tab in the downloaded Excel file
                    </div>
                    """, unsafe_allow_html=True)
                    with st.expander("View flagged items"):
                        flags_df = pd.DataFrame(flags)
                        st.dataframe(flags_df, use_container_width=True, hide_index=True)

                # Download button
                filename = f"PMI_Dispatch_{selected_month_name}_{selected_year}.xlsx"
                st.download_button(
                    label="📥 Download Dispatch Excel",
                    data=output_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except ValueError as e:
                st.markdown(f'<div class="error-box">❌ File error: {e}</div>',
                            unsafe_allow_html=True)
            except Exception as e:
                st.markdown(f'<div class="error-box">❌ Unexpected error: {e}<br>'
                            f'<small>Check that all files are in the correct format.</small></div>',
                            unsafe_allow_html=True)
                import traceback
                with st.expander("Technical details"):
                    st.code(traceback.format_exc())

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption("PMI Auto-Dispatcher | TXANT San Antonio Hub | Built by Andres | "
           "Quarterly rotation and mechanic schedule files managed by facility team")
