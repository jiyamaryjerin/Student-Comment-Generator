import ollama
from openpyxl import Workbook, load_workbook
import streamlit as st
import os
import base64
from io import BytesIO
from datetime import datetime


timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
file_path = f"student_comments.xlsx"

st.set_page_config(layout="wide")


def get_base64(bin_file):
    with open(bin_file, 'rb') as f:
        data = f.read()
    return base64.b64encode(data).decode()

def set_background(png_file):
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


set_background('./back6.png')


with open('./stylesheet.css') as f:
    st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)


if "workbook" not in st.session_state:
    st.session_state["workbook"] = Workbook()
    st.session_state["file_path"] = f"student_comments_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

# ✅ Get the active workbook & sheet
workbook = st.session_state["workbook"]
sheet = workbook.active

if sheet.max_row == 1:
    sheet["A1"] = "Student Comments"
    sheet["A3"] = "Name"
    sheet["B3"] = "Comment"

import requests

OPENROUTER_API_KEY = "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"
OPENROUTER_API_URL = "https://openrouter.ai/api/v1/chat/completions"

def query_mistral(prompt):
    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": "mistralai/mistral-7b-instruct",
        "messages": [{"role": "user", "content": prompt}]
    }
    response = requests.post(OPENROUTER_API_URL, json=payload, headers=headers)
    
    try:
        data = response.json()  # Convert response to dictionary
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



col1, spacer, col2 = st.columns([1.3, 0.2, 1])

with col1:
    with st.container(key="main"):
        with st.form("form1"):
            st.markdown('<p class="title">Student Comment Generator</p>', unsafe_allow_html=True)
            grade = st.text_input("Enter grade : ")
            name = st.text_input("Enter name : ")
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
            f"whose strengths include {strength}. Weaknesses include {weakness}. "
            "Make it sound positive. Write in third person. Do not add any emojis."
        )
        comment_text = query_mistral(prompt)
        #st.write(comment_text)
        st.write(comment_text)

        
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
