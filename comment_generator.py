import ollama
from openpyxl import Workbook, load_workbook
import streamlit as st
import os
import base64
from io import BytesIO
from datetime import datetime
import json
import re
import requests
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
file_path = f"student_comments.xlsx"
usage_file = "usage_tracking.json"

st.set_page_config(page_title='Student Comment Generator', page_icon="📝", layout="wide")

# Admin email list - add admin emails here
ADMIN_EMAILS = [
    "jiyamaryjerin04@gmail.com",
    "admin@example.com"
]


def get_base64(bin_file):
    with open(bin_file, 'rb') as f:
        data = f.read()
    return base64.b64encode(data).decode()


def set_background(png_file):
    if os.path.exists(png_file):
        bin_str = get_base64(png_file)
        page_bg_img = '''
        <style>
        .stApp {
        background-color : #B0D8F3;
        background-image: url("data:image/png;base64,%s");
        background-size: cover;
        }
        </style>
        ''' % bin_str
        st.markdown(page_bg_img, unsafe_allow_html=True)


def is_valid_email(email):
    """Validate email format"""
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email) is not None


def is_admin(email):
    """Check if email is in admin list"""
    return email.lower() in [admin.lower() for admin in ADMIN_EMAILS]


def load_usage_data():
    """Load or initialize usage tracking"""
    if os.path.exists(usage_file):
        try:
            with open(usage_file, 'r') as f:
                return json.load(f)
        except:
            return {}
    return {}


def save_usage_data(data):
    """Save usage data to JSON file"""
    with open(usage_file, 'w') as f:
        json.dump(data, f, indent=2)


def update_user_stats(email, username, action="login"):
    """Update user statistics"""
    usage_data = load_usage_data()
    
    if email not in usage_data:
        usage_data[email] = {
            "username": username,
            "login_count": 0,
            "comments_generated": 0,
            "first_login": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "last_login": ""
        }
    
    if action == "login":
        usage_data[email]["login_count"] += 1
        usage_data[email]["last_login"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    elif action == "comment":
        usage_data[email]["comments_generated"] += 1
    
    save_usage_data(usage_data)
    return usage_data[email]


def query_mistral(prompt):
    """Query the Mistral API"""
    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": "mistralai/mistral-7b-instruct",
        "messages": [{"role": "user", "content": prompt}]
    }
    
    try:
        response = requests.post(OPENROUTER_API_URL, json=payload, headers=headers)
        data = response.json()
        
        print("Full API Response:", data)  
        
        if "error" in data:
            return f"API Error: {data['error']}"
        if "choices" not in data or not isinstance(data["choices"], list) or len(data["choices"]) == 0:
            return f"Error: 'choices' key missing or empty. Response: {data}"
        first_choice = data["choices"][0]
        if "message" not in first_choice or "content" not in first_choice["message"]:
            return f"Error: 'message' or 'content' key missing. Response: {data}"

        return first_choice["message"]["content"]
    
    except Exception as e:
        return f"Error processing response: {str(e)}"


# Set background and load stylesheet
set_background('./back6.png')

if os.path.exists('./stylesheet.css'):
    with open('./stylesheet.css') as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

# Initialize session state for login
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
    st.session_state["username"] = ""
    st.session_state["email"] = ""

if "workbook" not in st.session_state:
    st.session_state["workbook"] = Workbook()
    st.session_state["file_path"] = f"student_comments_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

# Login Section
if not st.session_state["logged_in"]:
    st.markdown('<h1 style="text-align: center;">Student Comment Generator</h1>', unsafe_allow_html=True)
    st.markdown('<h3 style="text-align: center;">Please login to continue</h3>', unsafe_allow_html=True)
    
    col_login1, col_login2, col_login3 = st.columns([1, 2, 1])
    with col_login2:
        username_input = st.text_input("Enter your name:", key="username_input")
        email_input = st.text_input("Enter your email:", key="email_input")
        login_button = st.button("Login", use_container_width=True)
        
        if login_button:
            if not username_input.strip():
                st.error("❌ Please enter your name")
            elif not email_input.strip():
                st.error("❌ Please enter your email")
            elif not is_valid_email(email_input.strip()):
                st.error("❌ Please enter a valid email address (e.g., user@example.com)")
            else:
                st.session_state["logged_in"] = True
                st.session_state["username"] = username_input.strip()
                st.session_state["email"] = email_input.strip().lower()
                update_user_stats(st.session_state["email"], st.session_state["username"], action="login")
                st.success("✅ Login successful!")
                st.rerun()
    
    st.stop()

# Main Application (after login)
user_is_admin = is_admin(st.session_state["email"])

# Display sidebar (common to both admin and regular users)
if user_is_admin:
    with st.sidebar:
        st.markdown(f"### Welcome, {st.session_state['username']}! 👋")
        st.markdown(f"**Email:** {st.session_state['email']}")
        
        if user_is_admin:
            st.success("🔑 Admin Access")
        
        user_stats = load_usage_data().get(st.session_state["email"], {})
        
        st.markdown("---")
        st.markdown("### 📊 Your Statistics")
        st.metric("Total Logins", user_stats.get("login_count", 0))
        st.metric("Comments Generated", user_stats.get("comments_generated", 0))
        
        #if user_stats.get("first_login"):
            #st.info(f"**First Login:** {user_stats['first_login']}")
        if user_stats.get("last_login"):
            st.info(f"**Last Login:** {user_stats['last_login']}")
        
        st.markdown("---")
        if st.button("Logout", use_container_width=True):
            st.session_state["logged_in"] = False
            st.session_state["username"] = ""
            st.session_state["email"] = ""
            st.rerun()

# ==================== ADMIN DASHBOARD ====================
if user_is_admin:
    st.markdown('<h1 style="text-align: center;">🔐 Admin Dashboard</h1>', unsafe_allow_html=True)
    st.markdown('<h3 style="text-align: center;">User Statistics Overview</h3>', unsafe_allow_html=True)
    
    all_stats = load_usage_data()
    
    if not all_stats:
        st.info("No user data available yet.")
        st.stop()
    
    total_users = len(all_stats)
    total_logins = sum(stats.get('login_count', 0) for stats in all_stats.values())
    total_comments = sum(stats.get('comments_generated', 0) for stats in all_stats.values())
    
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.metric("Total Users", total_users)
    with col_b:
        st.metric("Total Logins", total_logins)
    with col_c:
        st.metric("Total Comments Generated", total_comments)
    
    if total_users > 0:
        st.metric("Average Comments per User", f"{total_comments / total_users:.1f}")
    
    st.markdown("---")
    st.markdown("### 👥 Detailed User Statistics")
    
    # Sort users by most comments generated
    sorted_users = sorted(all_stats.items(), key=lambda x: x[1].get('comments_generated', 0), reverse=True)
    
    for email, stats in sorted_users:
        with st.expander(f"{stats.get('username', 'Unknown User')} ({email}) — {stats.get('comments_generated', 0)} comments"):
            col1, col2 = st.columns(2)
            with col1:
                st.write(f"**Login Count:** {stats.get('login_count', 0)}")
                st.write(f"**Comments Generated:** {stats.get('comments_generated', 0)}")
            with col2:
                st.write(f"**First Login:** {stats.get('first_login', 'N/A')}")
                st.write(f"**Last Login:** {stats.get('last_login', 'N/A')}")
    
    # Footer for admin dashboard
    st.markdown(
        """
        <style>
        .footer {
            position: fixed;
            bottom: 0;
            left: 0;
            width: 100%;
            background-color: rgba(255, 255, 255, 0.6)!important;
            color: black;
            text-align: center;
            padding: 10px;
            font-size: 14px;
        }
        </style>
        <div class="footer">
            Designed and developed by Jiya Mary Jerin | For assistance/support reach out to jiyamaryjerin04@gmail.com
        </div>
        """,
        unsafe_allow_html=True
    )
    
    st.stop()  # Stop execution here for admins - they won't see the regular user interface

# ==================== REGULAR USER INTERFACE ====================
# (Only runs for non-admins)

workbook = st.session_state["workbook"]
sheet = workbook.active

if sheet.max_row == 1:
    sheet["A1"] = "Student Comments"
    sheet["A3"] = "Name"
    sheet["B3"] = "Comment"

OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY")

if not OPENROUTER_API_KEY:
    st.error("⚠️ API key not found! Please check your .env file.")
    st.info("Create a .env file in your project directory with: OPENROUTER_API_KEY=your_api_key_here")
    st.stop()

OPENROUTER_API_URL = "https://openrouter.ai/api/v1/chat/completions"

col1, spacer, col2 = st.columns([1.3, 0.2, 1])

with col1:
    with st.container(key="main"):
        with st.form("form1"):
            st.markdown('<p class="title">Student Comment Generator</p>', unsafe_allow_html=True)
            grade = st.text_input("Enter grade : ")
            name = st.text_input("Enter name : ")
            gender = st.text_input("Enter gender : ")
            strength = st.text_input("Enter strengths : ")
            weakness = st.text_input("Enter weakness : ")
            col3, col4 = st.columns([1, 1])
            with col3:
                style = st.radio("Select Comment Style :", ["Simple", "Funny", "Formal"])
            with col4:
                size = st.radio("Select Length :", ["50 words", "100 words", "150 words"])
                
            submit = st.form_submit_button("Generate Comment")

with col2:
    if submit:
        prompt = (
            f"Give a {style} progress card comment in {size} for a student named {name} "
            f"The student is a {gender} "
            f"whose strengths include {strength}. Weaknesses include {weakness}. "
            "Make it sound positive. Write in third person. Do not add any emojis."
        )
        
        with st.spinner("Generating comment..."):
            comment_text = query_mistral(prompt)
        
        update_user_stats(st.session_state["email"], st.session_state["username"], action="comment")
        
        st.write(comment_text)
        st.write()
        st.markdown("""
    <style>
    .disclaimer {
        color: black;
        text-align: center;
        font-size: 10px;
    }
    </style>
    <p class="disclaimer">
        <i>**This is a machine generated response using AI. Please review before use**</i>
    </p>
    """, unsafe_allow_html=True)
        
        next_row = sheet.max_row + 1
        sheet.cell(row=next_row, column=1, value=name)
        sheet.cell(row=next_row, column=2, value=comment_text)

        workbook.save(filename=file_path)
        st.balloons()

        excel_buffer = BytesIO()
        workbook.save(excel_buffer)
        excel_buffer.seek(0)

        st.download_button(
            label="Download Excel File",
            data=excel_buffer,
            file_name="student_comments.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# Footer for regular user interface
st.markdown(
    """
    <style>
    .footer {
        position: fixed;
        bottom: 0;
        left: 0;
        width: 100%;
        background-color: rgba(255, 255, 255, 0.6)!important;
        color: black;
        text-align: center;
        padding: 10px;
        font-size: 14px;
    }
    </style>
    <div class="footer">
        Designed and developed by Jiya Mary Jerin | For assistance/support reach out to jiyamaryjerin04@gmail.com
    </div>
    """,
    unsafe_allow_html=True
)
