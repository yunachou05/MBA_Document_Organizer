#!/usr/bin/env python
# coding: utf-8

# In[1]:


import os
import pandas as pd


# In[2]:


import sys
get_ipython().system('{sys.executable} -m pip install openpyxl')


# In[3]:


import time


# In[4]:


# 👉 Change this to your MBA folder path
folder_path = "MBA studying record"

file_data = []

modified_time = time.ctime(os.path.getmtime(folder_path))

# Walk through all folders and files
for root, dirs, files in os.walk(folder_path):
    for file in files:
        file_path = os.path.join(root, file)
        file_type = file.split('.')[-1] if '.' in file else 'Unknown'

        file_data.append({
            "File Name": file,
            "Folder Path": root,
            "Full Path": file_path,
            "File Type": file_type, "Last Modified": modified_time
        })

# Create DataFrame
df = pd.DataFrame(file_data)

# Save to Excel
output_file = "MBA_File_Index.xlsx"
df.to_excel(output_file, index=False)



# In[5]:


print("✅ Excel file created:", output_file)




############ Next Step #########      + AI function 




# In[7]:

import sys
get_ipython().system('{sys.executable} -m pip install pdfplumber python-docx')


# In[8]:

import pdfplumber
from docx import Document


# In[9]:
# Read PDF
def extract_text_from_pdf(file_path):
    text = ""
    try:
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                text += page.extract_text() or ""
    except:
        pass
    return text


# In[10]:
# Read Word file
def extract_text_from_docx(file_path):
    text = ""
    try:
        doc = Document(file_path)
        for para in doc.paragraphs:
            text += para.text + "\n"
    except:
        pass
    return text


# In[11]:
import re


# In[12]:
def extract_email(text):
    match = re.search(r'[\w\.-]+@[\w\.-]+', text)
    return match.group(0) if match else ""


# In[13]:
def extract_professor(text):
    lines = text.split("\n")
    for line in lines:
        if "Professor" in line or "Prof." in line:
            return line.strip()
    return ""


# In[15]:
pip install tqdm


# In[16]:

from tqdm import tqdm

file_list = []

# First collect all files
for root, dirs, files in os.walk(folder_path):
    for file in files:
        file_list.append((root, file))

# Loop with progress bar
for root, file in tqdm(file_list):
    file_path = os.path.join(root, file)


# In[17]:

total_files = 0

for root, dirs, files in os.walk(folder_path):
    total_files += len(files)

print(f"Total files: {total_files}")


# In[19]:
file_data = []

# First collect all files
for root, dirs, files in os.walk(folder_path):
    for file in files:
        file_list.append((root, file))

# Loop with progress bar
for root, file in tqdm(file_list):
        
    file_path = os.path.join(root, file)
    file_type = file.split('.')[-1] if '.' in file else 'Unknown'

    print(f"Processing: {file}")

    text = ""
        
    # Read file content
    if file_type == "pdf":
        text = extract_text_from_pdf(file_path)
    elif file_type == "docx":
        text = extract_text_from_docx(file_path)

        # Extract info
        email = extract_email(text) 
        professor = extract_professor(text)

        file_data.append({
            "File Name": file,
            "Folder Path": root,
            "File Type": file_type,
            "Professor": professor,
            "Email": email
        })


# In[20]:
df = pd.DataFrame(file_data)
df.to_excel("File Organizer_Index.xlsx", index=False)


# In[21]:
print(f"Processing: {file_path}")


# In[ ]:




