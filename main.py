import streamlit as st
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
    from github import Github  # For GitHub integration
    GITHUB_ENABLED = True
except ImportError:
    GITHUB_ENABLED = False
    st.warning("PyGithub not installed - GitHub updates disabled", icon="⚠️")

# Page setup with custom theme (MUST BE FIRST STREAMLIT COMMAND)
st.set_page_config(page_title="Testing Tool", layout="wide")
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
GITHUB_REPO = "Ai-TestingApp/Ai-Testing-Tool"
GITHUB_FILE = "main_excel.xlsx"

# Load Excel data
MAIN_EXCEL_PATH = "main_excel.xlsx"
df_main, wb = load_excel_data(MAIN_EXCEL_PATH)

# Clean and normalize Task IDs
df_main["Task ID"] = df_main["Task ID"].astype(str).str.strip()

# Sidebar navigation
st.sidebar.title("🛍️ Navigation")
page = st.sidebar.radio("Go to", ["Testing App", "Excel Sheet", "Analytics"])

# Improved ID normalization function
def normalize_id(task_id):
    """Convert all task IDs to consistent string format"""
    try:
        # Handle string inputs
        if isinstance(task_id, str):
            if '.' in task_id:
                return task_id  # Keep decimal IDs as-is (2.0, 2.1 etc.)
            return str(int(float(task_id)))  # Convert "2" to "2"
        
        # Handle numeric inputs
        if float(task_id).is_integer():
            return str(int(task_id))  # Convert 2.0 to "2"
        return str(task_id)  # Keep 2.1 as "2.1"
    except:
        return str(task_id)  # Fallback for any format

# Graph plotting function (unchanged)
def plot_test_result_summary(df):
    result_counts = df['Test Result'].dropna().value_counts()
    result_counts = result_counts.reindex(['Pass', 'Fail', 'Hold']).fillna(0)

    if result_counts.sum() == 0:
        st.warning("No test results available to display.")
        return

    fig, ax = plt.subplots()
    colors = ['#28a745', '#dc3545', '#ffc107']
    ax.pie(result_counts, labels=result_counts.index, autopct='%1.1f%%', startangle=90, colors=colors)
    ax.axis('equal')
    st.markdown("### 📊 Test Result Summary")
    st.pyplot(fig)

if page == "Testing App":
    st.title("🔪 Testing Documentation Tool")

    # Tester Selection
    tester_names = sorted(df_main["Tester Name"].dropna().unique())
    tester_name = st.selectbox("👤 Select Tester Name", tester_names)

    # Filter and prepare tasks
    tester_tasks = df_main[df_main["Tester Name"] == tester_name].copy()
    tester_tasks["Normalized_ID"] = tester_tasks["Task ID"].apply(normalize_id)
    
    # Debug output (can remove later)
    st.sidebar.write("Debug - Task IDs:", tester_tasks["Task ID"].unique())
    st.sidebar.write("Normalized IDs:", tester_tasks["Normalized_ID"].unique())

    # Determine completed tasks
    completed_ids = df_main[df_main["Test Result"].notna()]["Task ID"].apply(normalize_id).tolist()

    # Prepare task availability
    available_task_ids = []
    disabled_task_ids = []
    sorted_task_ids = tester_tasks.sort_values("Task ID")["Task ID"].unique()

    for i, tid in enumerate(sorted_task_ids):
        norm_tid = normalize_id(tid)
        if norm_tid in completed_ids:
            disabled_task_ids.append(tid)
        elif i == 0 or normalize_id(sorted_task_ids[i-1]) in completed_ids:
            available_task_ids.append(tid)

    # Display task options
    task_display_options = [
        f"{tid} ✅ (Completed)" if tid in disabled_task_ids else
        f"{tid} 🔒 (Locked)" if tid not in available_task_ids else
        tid
        for tid in sorted_task_ids
    ]

    if available_task_ids:
        task_id = st.selectbox("🆔 Select Task ID", 
                             options=available_task_ids,
                             format_func=lambda x: task_display_options[sorted_task_ids.tolist().index(x)])
        
        # Find matching task
        search_id = normalize_id(task_id)
        selected_row = tester_tasks[tester_tasks["Normalized_ID"] == search_id]

        if not selected_row.empty:
            selected_row = selected_row.iloc[0]
            with st.expander("📋 Task Details", expanded=True):
                st.text_input("📝 Task Heading", selected_row.get("Task Name", ""), disabled=True)
                st.text_input("🛍️ Navigation", selected_row.get("Navigation", ""), disabled=True)
                st.text_input("⚙️ Parameters", selected_row.get("Parameters", ""), disabled=True)

            # Test submission
            with st.expander("📝 Submit Test Result", expanded=True):
                test_result = st.selectbox("✅ Test Result", ["Pass", "Fail", "Hold"])
                comment = st.text_area("💬 Comment", key=f"comment_{task_id}")
                
                screenshots = st.file_uploader(
                    "📌 Upload Screenshot(s)", 
                    type=["png", "jpg", "jpeg"], 
                    accept_multiple_files=True,
                    key=f"screenshots_{task_id}"
                )

                if screenshots:
                    cols = st.columns(min(3, len(screenshots)))
                    for i, img_file in enumerate(screenshots):
                        with cols[i % 3]:
                            st.image(Image.open(img_file), caption=img_file.name, use_column_width=True)

                if st.button("✅ Submit Task"):
                    output = io.BytesIO()
                    screenshots = screenshots if screenshots else []

                    # Save results to in-memory Excel file
                    save_screenshots_to_excel(
                        excel_path=output,
                        df_main=df_main,
                        wb=wb,
                        task_id=task_id,
                        tester_name=tester_name,
                        test_result=test_result,
                        comment=comment,
                        screenshots=screenshots
                    )

                    # Get raw bytes of the Excel file
                    excel_bytes = output.getvalue()

                    # Push to GitHub (before UI feedback)
                    if GITHUB_ENABLED and 'GITHUB_TOKEN' in st.secrets:
                        try:
                            st.write("🔄 Pushing to GitHub...")
                            g = Github(st.secrets.GITHUB_TOKEN)
                            repo = g.get_repo(GITHUB_REPO)
                            contents = repo.get_contents(GITHUB_FILE)
                            repo.update_file(
                                path=GITHUB_FILE,
                                message=f"Update by {tester_name} on Task {task_id}",
                                content=excel_bytes,
                                sha=contents.sha
                            )
                            st.success("✅ Updated GitHub successfully!")
                        except Exception as e:
                            st.warning(f"⚠️ GitHub update failed: {str(e)}")

                    # Offer file for download
                    st.download_button(
                        label="📥 Download Updated Excel",
                        data=excel_bytes,
                        file_name="updated_results.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    st.balloons()
                    time.sleep(1.5)
                    # Optional: rerun the app to refresh task list
                    # st.rerun()


        else:
            st.error(f"❌ No data found for Task ID {task_id} (searched as {search_id})")
    else:
        st.success("🎉 All tasks completed!")

# [Rest of your code for Excel Sheet and Analytics pages remains exactly the same...])
        

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




       
