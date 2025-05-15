import streamlit as st
import requests
import base64
import json
import os
import openpyxl
from pathlib import Path
import hashlib
import io
import zipfile
from PIL import Image

API_KEY = st.secrets["API_KEY"]
HEADERS = {"Content-Type": "application/json"}

if 'processed_files' not in st.session_state:
    st.session_state.processed_files = set()

# No longer needed for naming the Excel file consistently
# if 'uploaded_file_names' not in st.session_state:
#     st.session_state.uploaded_file_names = []

def get_file_hash(file_content):
    return hashlib.md5(file_content).hexdigest()

def encode_image(image_path):
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode("utf-8")

def extract_info_from_image(image_path):
    url = f"https://generativelanguage.googleapis.com/v1/models/gemini-1.5-flash:generateContent?key={API_KEY}"
    prompt = {
        "contents": [
            {
                "parts": [
                    {
                        "text": "Extract the Company Name, Person's Name, Designation, Phone, Email, Website, and Address from this business card. Respond only with JSON format. Use exactly these field names: 'Company Name', 'Person Name', 'Designation', 'Phone', 'Email', 'Website', 'Address'. Do NOT use triple backticks or markdown."
                    },
                    {
                        "inlineData": {
                            "mimeType": "image/jpeg",
                            "data": encode_image(image_path)
                        }
                    }
                ]
            }
        ]
    }
    response = requests.post(url, headers=HEADERS, data=json.dumps(prompt))
    if response.status_code == 200:
        try:
            data = response.json()
            extracted_text = data['candidates'][0]['content']['parts'][0]['text']
            for marker in ["```json", "```"]:
                extracted_text = extracted_text.replace(marker, "").strip()
            return json.loads(extracted_text)
        except Exception as e:
            st.error(f"Error processing image: {e}")
            return None
    else:
        st.error(f"API Error: {response.status_code} - {response.text}")
        return None

def normalize_fields(info_dict):
    field_mapping = {
        "company name": "Company Name",
        "company's name": "Company Name",
        "company": "Company Name",
        "person's name": "Person Name",
        "person name": "Person Name",
        "name": "Person Name",
        "person": "Person Name",
        "full name": "Person Name",
        "designation": "Designation",
        "title": "Designation",
        "job title": "Designation",
        "position": "Designation",
        "role": "Designation",
        "phone": "Phone",
        "phone number": "Phone",
        "mobile": "Phone",
        "telephone": "Phone",
        "tel": "Phone",
        "contact": "Phone",
        "email": "Email",
        "mail": "Email",
        "e-mail": "Email",
        "email address": "Email",
        "website": "Website",
        "web": "Website",
        "url": "Website",
        "site": "Website",
        "web address": "Website",
        "address": "Address",
        "location": "Address",
        "office": "Address",
        "office address": "Address",
    }
    normalized_data = {}
    for field, value in info_dict.items():
        normalized_field = field.lower().replace("'", "").strip()
        standard_field = next((std_field for key, std_field in field_mapping.items() if normalized_field == key or key in normalized_field), field)
        normalized_data[standard_field] = value
    return normalized_data

# Consistent Excel file name
EXCEL_FILE_NAME = "extracted_contacts.xlsx"

def save_to_excel(info_dict, file_base_name):
    documents_path = Path.cwd() / "documents"
    documents_path.mkdir(exist_ok=True)
    full_path = documents_path / EXCEL_FILE_NAME # Use the consistent name
    headers = ["Company Name", "Person Name", "Designation", "Phone", "Email", "Website", "Address"]
    normalized_data = normalize_fields(info_dict)
    new_row = [normalized_data.get(header, "") for header in headers]

    try:
        if full_path.is_file():
            wb = openpyxl.load_workbook(full_path)
            ws = wb.active
            existing_rows = [
                tuple(str(cell) for cell in row)
                for row in ws.iter_rows(min_row=2, values_only=True)
            ]
            new_row_tuple = tuple(str(cell) for cell in new_row)
            if new_row_tuple not in existing_rows:
                next_row = ws.max_row + 1
                for col_idx, value in enumerate(new_row, start=1):
                    ws.cell(row=next_row, column=col_idx).value = value
                wb.save(full_path)
                return str(full_path)
            else:
                st.info("This card's information is already in the Excel sheet.")
                return None
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Business Cards"
            ws.append(headers)
            ws.append(new_row)
            wb.save(full_path)
            return str(full_path)
    except Exception as e:
        st.error(f"Failed to save data: {e}")
        return None

def process_file(file_content, file_name):
    try:
        file_hash = get_file_hash(file_content)
        if file_hash in st.session_state.processed_files:
            st.info(f"Skipping already processed file: {file_name}")
            return True

        img = Image.open(io.BytesIO(file_content))
        temp_path = os.path.join("uploaded_cards", file_name)
        os.makedirs("uploaded_cards", exist_ok=True)
        img.save(temp_path)

        file_base_name = os.path.splitext(file_name)[0]

        with st.spinner(f"Processing {file_name}..."):
            info = extract_info_from_image(temp_path)
            if info:
                saved_path = save_to_excel(info, file_base_name)
                if saved_path:
                    st.success(f"Data from '{file_name}' added to Excel!")
                    # No longer appending individual file names for Excel naming
                    # st.session_state.uploaded_file_names.append(file_name)
                st.session_state.processed_files.add(file_hash)

        os.remove(temp_path)
        return True
    except Exception as e:
        st.error(f"Error processing {file_name}: {e}")
        return False

def main():
    st.set_page_config(page_title="CardSnap", layout="centered")
    st.title("Welcome to cardSnap")
    st.write("Upload one or more business card images (PNG, JPG, JPEG) or a ZIP file of images. We'll extract the contact info and save it to Excel!")

    if st.button("Clear processed files history"):
        st.session_state.processed_files = set()
        # st.session_state.uploaded_file_names = [] # No longer strictly needed
        st.success("Processing history cleared! All uploaded files will be processed again.")

    uploaded_files = st.file_uploader("Upload Business Card Images or ZIP file", type=["png", "jpg", "jpeg", "zip"], accept_multiple_files=True)

    if uploaded_files:
        new_files_processed = False
        for uploaded_file in uploaded_files:
            file_name = uploaded_file.name
            file_content = uploaded_file.read()

            if file_name.endswith(".zip"):
                try:
                    zip_hash = get_file_hash(file_content)
                    if zip_hash in st.session_state.processed_files:
                        st.info(f"Skipping already processed ZIP file: {file_name}")
                        continue

                    with zipfile.ZipFile(io.BytesIO(file_content), 'r') as zip_ref:
                        for member in zip_ref.namelist():
                            if member.lower().endswith(('.png', '.jpg', '.jpeg')):
                                with zip_ref.open(member) as image_file:
                                    image_content = image_file.read()
                                    if process_file(image_content, os.path.basename(member)):
                                        new_files_processed = True
                    st.session_state.processed_files.add(zip_hash)
                except zipfile.BadZipFile:
                    st.error(f"Error: '{file_name}' is not a valid ZIP file.")
                except Exception as e:
                    st.error(f"Error processing '{file_name}': {e}")
            else:
                if process_file(file_content, file_name):
                    new_files_processed = True

        if st.session_state.processed_files: # Check if any files were processed
            excel_path = Path.cwd() / "documents" / EXCEL_FILE_NAME
            if excel_path.exists():
                with open(excel_path, "rb") as f:
                    st.download_button(
                        label="Download Combined Excel File",
                        data=f.read(),
                        file_name=EXCEL_FILE_NAME,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        if new_files_processed:
            st.write(f"Currently processed {len(st.session_state.processed_files)} unique files.")

if __name__ == "__main__":
    main()