import streamlit as st
from jira import JIRA
import pandas as pd
from datetime import datetime

def jiraAPIfront():
    # Jira server URL and credentials
    JIRA_SERVER = ''
    JIRA_API_TOKEN = ''

    # Function to connect to Jira
    def connect_to_jira(email):
        return JIRA(server=JIRA_SERVER, basic_auth=(email, JIRA_API_TOKEN))

    # Function to fetch issues based on user inputs
    def fetch_issues(jira, project_key, start_date, end_date):
        jql_query = f'project = {project_key} AND created >= "{start_date}" AND created <= "{end_date}"'
        return jira.search_issues(jql_query, maxResults=100000)

    # Function to extract data from issues
    def extract_data_from_issues(issues):
        custom_field_ids = {
            'projid': 'customfield_10034',
            'priority': 'customfield_10035',
            'locate': 'customfield_10036'
        }
        
        data = []
        for issue in issues:
            comments = [comment.body for comment in issue.fields.comment.comments] if issue.fields.comment.comments else []
            comments_concatenated = "\n".join(comments)
            
            issue_data = {
                'Key': issue.key,
                'Summary': issue.fields.summary,
                'Status': issue.fields.status.name,
                'Assignee': issue.fields.assignee.displayName if issue.fields.assignee else None,
                'Created': issue.fields.created,
                'projid': getattr(issue.fields, custom_field_ids['projid'], None),
                'priority': getattr(issue.fields, custom_field_ids['priority'], None),
                'locate': getattr(issue.fields, custom_field_ids['locate'], None),
                'Comment': comments_concatenated,
                'ruleName': None,
                'ruleCate': None
            }
            data.append(issue_data)
        return pd.DataFrame(data)

    # Streamlit UI
    st.title("JIRA Data Export Tool")

    st.write("Build to export data form JIRA API with specified/required rows&columns.")
    st.warning("Unauthorized do not run the script")

    # User inputs
    email = st.text_input("Enter your Jira email", "")
    project_key = st.selectbox("Select Project Key", ["meow_meow", "peepeepoopoo", "test"])

    col1, col2 = st.columns([2, 2])
    with col1:
        start_date = st.date_input("Start Date", datetime.now() - pd.DateOffset(days=30))
    with col2:
        end_date = st.date_input("End Date", datetime.now())

    if st.button("Export Data"):
        if email and start_date and end_date:
            jira = connect_to_jira(email)
            issues = fetch_issues(jira, project_key, start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d'))
            df = extract_data_from_issues(issues)
            
            excel_file = 'output.xlsx'
            df.to_excel(excel_file, index=False)
            st.success(f'Data exported to {excel_file}')
        else:
            st.error("Please fill in all the fields.")
    
    st.write("")
    st.write("")
    st.write("")

    # Back button
    if st.button("Back to Index"):
        st.session_state.page = 'index'

    
    # Label
    st.markdown(
        """
        <style>
        .footer {
            position: fixed;
            bottom: 0;
            left: 0;
            width: 100%;
            text-align: center;
            padding: 10px;
            background-color: #0E1117;
            color: #fff;
            font-size: 14px;
        }
        </style>
        <div class="footer">
            ᚱᛟᛟᛏ@ᛞᛖᛖ:~#
        </div>
        """,
        unsafe_allow_html=True
    )