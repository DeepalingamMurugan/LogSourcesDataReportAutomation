import streamlit as st

# Initialize session state for page tracking if not already set
if 'page' not in st.session_state:
    st.session_state.page = 'index' 

# Navigation using selectbox
def navigate_to(page):
    st.session_state.page = page

# Page selection
if st.session_state.page == 'index':
    st.title("Beta Build ~ under test")
    st.markdown("""
                # select a script to navigate & run
                double click to redirect""")

    st.write("")
    st.write("")

    # Navigation buttons
    if st.button("JIRA Data Exporter"):
        navigate_to('jiraApiTestFE')
    st.write("")

    if st.button("Rule Category Mapper"):
        navigate_to('FRCFE')
    st.write("")
    
    if st.button("QRadar Integration Processesor"):
        navigate_to('QRadIntRep')
    st.write("")

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

# Navigation assigning
elif st.session_state.page == 'jiraApiTestFE':
    import jiraAPItestfrontend
    jiraAPItestfrontend.jiraAPIfront()
    
elif st.session_state.page == 'FRCFE':
    import JRCfrontend
    JRCfrontend.jrcFront()

elif st.session_state.page == 'QRadIntRep':
    import qradint
    qradint.qradintfun()