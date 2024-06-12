import streamlit as st

# Initialize session state for login
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False

# Define hardcoded credentials
CORRECT_USERNAME = "soj"
CORRECT_PASSWORD = "soj"

# Function to check login credentials
def authenticate(username, password):
    return username == CORRECT_USERNAME and password == CORRECT_PASSWORD

# Main function for the login page
def login():
    st.set_page_config(page_title="Student Attendance Report", layout="wide")
    
    st.markdown("""
        <style>
            .reportview-container {margin-top: -2em;}
            .st-emotion-cache-1jicfl2 {padding: 2rem 3rem 10rem;}
            h1#student-attendance-report {text-align: center;}
            header #MainMenu {visibility: hidden; display: none;}
            .stActionButton {visibility: hidden; display: none;}
            # .stDeployButton {display:none;}
            footer {visibility: hidden;}
            stDecoration {display:none;}
            .stTabs button {margin-right: 50px;}
            .viewerBadge_container__r5tak {display: none;}
            p.credits {user-select: none; filter: opacity(0);}
        </style>
    """, unsafe_allow_html=True)

    
    st.markdown("<p class='credits'>Made by <a href='https://github.com/sojith29034'>Sojith Sunny</a></p>", unsafe_allow_html=True)
    st.title("Login to Your App")
    
    # Get user input
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")
    
    # Check if login button is clicked
    if st.button("Login"):
        if authenticate(username, password):
            st.session_state.logged_in = True
            st.experimental_rerun()
        else:
            st.error("Invalid username or password. Please try again.")

def main():
    if st.session_state.logged_in:
        import index
        index.run_main_app()
    else:
        login()

if __name__ == "__main__":
    main()