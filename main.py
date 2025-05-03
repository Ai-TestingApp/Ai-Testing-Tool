import streamlit as st
st.set_page_config(page_title="Testing Tool", layout="wide")
import pandas as pd
from utils import load_excel_data, save_screenshots_to_excel
from PIL import Image
import io
import os
import matplotlib.pyplot as plt
import seaborn as sns
import time
# Replace the GitHub import with this try-except block
try:
    from github import Github  # type: ignore # For GitHub integration
    GITHUB_ENABLED = True
except ImportError:
    GITHUB_ENABLED = False
    st.warning("PyGithub not installed - GitHub updates disabled", icon="⚠️")

# Page setup with custom theme
st.markdown("""
    <style>
    .main { background-color: #f5f7fa; }
    .block-container { padding: 2rem 3rem; }
    .css-1d391kg { background-color: #ffffff; border-radius: 8px; padding: 2rem; }
    .stButton>button { background-color: #0b6efd; color: white; font-weight: bold; border-radius: 6px; }
    .stTextInput>div>div>input { background-color: #eef2f7; border-radius: 5px; }
    .stSelectbox>div>div>div>div { background-color: #eef2f7; border-radius: 5px; }
    .scrollable-table { overflow-x: auto; }
    </style>
""", unsafe_allow_html=True)

# NEW: GitHub configuration
GITHUB_REPO = "Ai-TestingApp/Ai-Testing-Tool"  # Change this
GITHUB_FILE = "main_excel.xlsx"

MAIN_EXCEL_PATH = "main_excel.xlsx"
df_main, wb = load_excel_data(MAIN_EXCEL_PATH)

# Sidebar navigation
st.sidebar.title("🛍️ Navigation")
page = st.sidebar.radio("Go to", ["Testing App", "Excel Sheet", "Analytics"])

# Helper to normalize task IDs like 1.0 vs 1
normalize_id = lambda tid: str(int(float(tid))) if float(tid).is_integer() else str(tid)

# Graph plotting function
def plot_test_result_summary(df):
    result_counts = df['Test Result'].dropna().value_counts()
    result_counts = result_counts.reindex(['Pass', 'Fail', 'Hold']).fillna(0)

    if result_counts.sum() == 0:
        st.warning("No test results available to display.")
        return

    fig, ax = plt.subplots()
    colors = ['#28a745', '#dc3545', '#ffc107']
    ax.pie(result_counts, labels=result_counts.index, autopct='%1.1f%%', startangle=90, colors=colors)
    ax.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
    st.markdown("### 📊 Test Result Summary")
    st.pyplot(fig)


if page == "Testing App":
    st.title("🔪 Testing Documentation Tool")

    # Tester Selection
    tester_names = sorted(df_main["Tester Name"].dropna().unique())
    tester_name = st.selectbox("👤 Select Tester Name", tester_names)

    # Filter tasks for selected tester
    tester_tasks = df_main[df_main["Tester Name"] == tester_name].copy()
    tester_tasks["Task ID"] = tester_tasks["Task ID"].astype(str)
    tester_tasks.sort_values("Task ID", key=lambda col: col.map(lambda x: [int(i) if i.isdigit() else i for i in x.split('.')]), inplace=True)

    # Determine completed task IDs
    completed_ids = df_main[df_main["Test Result"].notna()]["Task ID"].astype(str).apply(normalize_id).tolist()

    # Determine which task IDs to show and enable
    available_task_ids = []
    disabled_task_ids = []
    sorted_task_ids = tester_tasks["Task ID"].tolist()

    for i, tid in enumerate(sorted_task_ids):
        norm_tid = normalize_id(tid)
        if norm_tid in completed_ids:
            disabled_task_ids.append(tid)
        elif i == 0 or normalize_id(sorted_task_ids[i - 1]) in completed_ids:
            available_task_ids.append(tid)
        else:
            break

    task_display_options = []
    for tid in sorted_task_ids:
        label = tid
        if tid in disabled_task_ids:
            label += " ✅ (Completed)"
        elif tid not in available_task_ids:
            label += " 🔒 (Locked)"
        task_display_options.append(label)

    if available_task_ids:
        task_id = st.selectbox("🆔 Select Task ID", options=available_task_ids)

        # Auto-fill metadata (check if selected task ID exists)
        selected_row = tester_tasks[tester_tasks["Task ID"] == task_id]

        if not selected_row.empty:
            selected_row = selected_row.iloc[0]
            task_heading = selected_row.get("Task Name", "")
            navigation = selected_row.get("Navigation", "")
            parameters = selected_row.get("Parameters", "")

            with st.expander("📋 Task Details", expanded=True):
                st.text_input("📝 Task Heading", task_heading, disabled=True)
                st.text_input("🛍️ Navigation", navigation, disabled=True)
                st.text_input("⚙️ Parameters", parameters, disabled=True)

            # User Inputs
            with st.expander("📝 Submit Test Result", expanded=True):
                test_result = st.selectbox("✅ Test Result", ["Pass", "Fail", "Hold"])
                
                # Add unique key for the comment box to reset on page rerun
                comment = st.text_area("💬 Comment", key=f"comment_{task_id}")
                
                # Add a unique key to reset the file uploader widget on page rerun
                screenshots = st.file_uploader(
                    "📌 Upload Screenshot(s)", 
                    type=["png", "jpg", "jpeg"], 
                    accept_multiple_files=True,
                    key=f"screenshots_{task_id}"  # Use task_id to reset uploader
                )

                if screenshots:
                    st.markdown("**📸 Preview Screenshots:**")
                    cols = st.columns(min(3, len(screenshots)))
                    for i, img_file in enumerate(screenshots):
                        img = Image.open(img_file)
                        with cols[i % 3]:
                            st.image(img, caption=img_file.name, use_column_width=True)

                # MODIFIED SUBMIT LOGIC:
                if st.button("✅ Submit Task"):
                    if not screenshots:
                        st.error("❗ Please upload at least one screenshot.")
                    else:
                        # Create in-memory Excel
                        output = io.BytesIO()
                        save_screenshots_to_excel(
                            excel_path=output,  # Now accepts BytesIO
                            df_main=df_main,
                            wb=wb,
                            task_id=task_id,
                            tester_name=tester_name,
                            test_result=test_result,
                            comment=comment,
                            screenshots=screenshots
                        )
                        excel_bytes = output.getvalue()
                        
                        # 1. Keep download option
                        st.download_button(
                            label="📥 Download Updated Excel",
                            data=excel_bytes,
                            file_name="updated_results.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                        # 2. GitHub update (if configured)
                        if 'github_token' in st.secrets:
                            try:
                                g = Github(st.secrets.github_token)
                                repo = g.get_repo(GITHUB_REPO)
                                contents = repo.get_contents(GITHUB_FILE)
                                repo.update_file(
                                    path=GITHUB_FILE,
                                    message=f"Update by {tester_name}",
                                    content=excel_bytes,
                                    sha=contents.sha
                                )
                                st.success("✅ Updated GitHub successfully!")
                            except Exception as e:
                                st.warning(f"⚠️ GitHub update failed: {str(e)}")
                        
                        st.balloons()
                        time.sleep(2)
                        st.rerun()
                else:
                    st.error(f"❌ No data found for Task ID {task_id}.")
        else:
            st.markdown("<span style='color: green; font-weight: bold;'>🎉 You have completed all your assigned tasks. No tasks are left.</span>", unsafe_allow_html=True)

elif page == "Excel Sheet":
    st.title("📄 Excel Sheet Viewer")

    # Add a tester name filter
    tester_filter = st.sidebar.selectbox("👤 Filter by Tester", ["All"] + sorted(df_main["Tester Name"].dropna().unique().tolist()))

    if tester_filter != "All":
        filtered_df = df_main[df_main["Tester Name"] == tester_filter]
        st.write(f"Showing tasks assigned to **{tester_filter}**:")
    else:
        filtered_df = df_main
        st.write("Showing all tasks:")

    with st.container():
        st.markdown("<div class='scrollable-table'>", unsafe_allow_html=True)
        st.dataframe(filtered_df, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)
    

    
elif page == "Analytics":
    st.title("📊 Analytics Dashboard")

    # Filter rows with either Test Result or Timestamp
    df_filtered = df_main[(df_main["Test Result"].notna()) | (df_main["Timestamp"].notna())].copy()

    # Create layout
    col1, col2 = st.columns(2)

    # -- Graph 1: Test Result Summary Pie (LEFT) --
    with col1:
        result_counts = df_filtered['Test Result'].dropna().value_counts()
        result_counts = result_counts.reindex(['Pass', 'Fail', 'Hold']).fillna(0)

        if result_counts.sum() > 0:
            fig, ax = plt.subplots()
            colors = ['#28a745', '#dc3545', '#ffc107']
            ax.pie(result_counts, labels=result_counts.index, autopct='%1.1f%%', startangle=90, colors=colors)
            ax.axis('equal')
            st.markdown("### 📊 Test Result Summary")
            st.pyplot(fig)
        else:
            st.warning("No test results available to display.")

    # -- Graph 2: Task Completion Over Time (RIGHT) --
        # -- Graph 2: Task Completion Over Time (RIGHT) --
    with col2:
        if "Timestamp" in df_filtered.columns:
            df_filtered["Date"] = pd.to_datetime(df_filtered["Timestamp"], errors='coerce').dt.date
            df_filtered = df_filtered.dropna(subset=["Date"])

            if not df_filtered.empty:
                # Add date range filter
                min_date = df_filtered["Date"].min()
                max_date = df_filtered["Date"].max()
                start_date, end_date = st.date_input(
                    "📅 Select Date Range",
                    value=(min_date, max_date),
                    min_value=min_date,
                    max_value=max_date,
                )

                filtered_dates = df_filtered[
                    (df_filtered["Date"] >= start_date) & (df_filtered["Date"] <= end_date)
                ]
                date_summary = filtered_dates.groupby("Date").size().reset_index(name="Tasks Completed")

                st.markdown("### 📈 Task Completion Over Time")
                fig, ax = plt.subplots(figsize=(6, 4))
                sns.lineplot(data=date_summary, x="Date", y="Tasks Completed", marker="o", ax=ax)
                ax.set_ylabel("Tasks")
                ax.set_xlabel("Date")
                ax.set_title("Tasks Completed Per Day")

                # Format the x-axis
                ax.set_xticks(date_summary["Date"])
                ax.set_xticklabels(date_summary["Date"].astype(str), rotation=45, ha='right')

                st.pyplot(fig)
            else:
                st.info("No valid timestamp data found.")
    # -- Graph 3: Tasks Completed Per Tester (BOTTOM) --
    st.markdown("### 🧑‍💻 Tasks Completed Per Tester")

    tester_filter = st.selectbox("👤 Filter by Tester", ["All"] + sorted(df_filtered["Tester Name"].dropna().unique()))

    if tester_filter != "All":
        tester_data = df_filtered[df_filtered["Tester Name"] == tester_filter]
    else:
        tester_data = df_filtered

    tester_summary = tester_data["Tester Name"].value_counts().reset_index()
    tester_summary.columns = ["Tester Name", "Tasks Completed"]

    if not tester_summary.empty:
        fig, ax = plt.subplots(figsize=(6, 3))  # Smaller size
        sns.barplot(data=tester_summary, x="Tester Name", y="Tasks Completed", palette="Blues_d", ax=ax)
        
        ax.set_xlabel("Tester", fontsize=10)
        ax.set_ylabel("Task Count", fontsize=10)
        ax.set_title("Tasks Completed Per Tester", fontsize=12)
        ax.tick_params(axis='x', labelrotation=0, labelsize=8)
        ax.tick_params(axis='y', labelsize=8)
        fig.tight_layout()  # Adjust spacing to avoid cut-offs

        st.pyplot(fig)
    else:
        st.info("No tester task completion data available.")
        # -- Completion Progress Bar --
    total_tasks = df_main.shape[0]
    completed_tasks = df_main["Test Result"].notna().sum()
    completion_percent = int((completed_tasks / total_tasks) * 100) if total_tasks > 0 else 0

    st.markdown("### 📈 Overall Task Completion")

    # Display as: "4 / 9 tasks completed (44%)"
    progress_text = f"{completed_tasks} / {total_tasks} tasks completed (Overall Completion - {completion_percent}%)"
    st.markdown(f"<h5 style='text-align: left;'>{progress_text}</h5>", unsafe_allow_html=True)

    # Render progress bar
    st.progress(completion_percent)




       
