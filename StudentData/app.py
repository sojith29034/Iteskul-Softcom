import os
import streamlit as st
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Initialize session state for login
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False

# Function to check login credentials
def authenticate(username, password):
    CORRECT_USERNAME = os.getenv('USERNAME')
    CORRECT_PASSWORD = os.getenv('PASSWORD')
    
    # Logging to check if environment variables are loaded correctly
    st.write(f"Loaded USERNAME from .env: {CORRECT_USERNAME}")
    st.write(f"Loaded PASSWORD from .env: {CORRECT_PASSWORD}")

    return username == CORRECT_USERNAME and password == CORRECT_PASSWORD

# Main function for the login page
def login():
    # Custom CSS to hide Streamlit elements
    st.markdown("""
        <style>
            .reportview-container {margin-top: -2em;}
            .st-emotion-cache-1jicfl2 {padding: 2rem 3rem 10rem;}
            h1#login-to-your-app, h1#student-atendance-report {text-align: center;}
            header #MainMenu {visibility: hidden; display: none;}
            .stActionButton {visibility: hidden; display: none;}
            footer {visibility: hidden;}
            .stDecoration {display:none;}
            .stTabs button {margin-right: 50px;}
            .st-emotion-cache-15ecox0, .viewerBadge_container__r5tak, .styles_viewerBadge__CvC9N {display: none;}
            p.credits {user-select: none; filter: opacity(0);}
        </style>
    """, unsafe_allow_html=True)

    # Credit footer
    st.markdown("<p class='credits'>Made by <a href='https://github.com/sojith29034'>Sojith Sunny</a></p>", unsafe_allow_html=True)
    
    # Title and instructions
    st.title("Login to Your App")
    st.markdown("<p>Username: qwerty <br> Password: qwerty</p>", unsafe_allow_html=True)
    
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

# Main function to run the app
def main():
    if st.session_state.logged_in:
        st.write("You are logged in!")  # Logging to confirm the login state
        # Ensure you have an `index.py` with a `run_main_app` function
        try:
            import index
            index.run_main_app()
        except ModuleNotFoundError:
            st.error("The index module was not found. Ensure it exists and is correctly named.")
        except AttributeError:
            st.error("The index module does not contain a `run_main_app` function.")
    else:
        login()

if __name__ == "__main__":
    main()
