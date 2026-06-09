import streamlit as st

import sys
sys.path.append('/content/drive/MyDrive/Python/Packages/')

import pandas as pd
from io import BytesIO
import requests
from datetime import datetime
from docxtpl import DocxTemplate
from docx2pdf import convert
import jinja2
from IPython.display import Javascript
from docx import Document # Added import
import os # Added import
from docx import Document
from docxcompose.composer import Composer

# Replace 'your_file_name.docx' with the actual path to your file in Google Drive
template_file = '/content/drive/MyDrive/Python/ST26-template.docx'
output_file = 'combined_output_2026.docx'

# Spreadsheet data Surat Tugas
url = "https://docs.google.com/spreadsheets/d/1SORCi_jXxEN-HSXWBOjX19FY8a6FzegPSAPbuh5k1sI/export?format=xlsx"
Kegiatan = pd.read_excel(url, sheet_name='Kegiatan', header=2)

# Ensure 'TanggalAwal' is datetime for filtering and handle errors
Kegiatan['TanggalAwal'] = pd.to_datetime(Kegiatan['TanggalAwal'], format='%d/%m/%Y', errors='coerce')
# Ensure 'TanggalAkhir' is datetime for filtering and handle errors
Kegiatan['TanggalAkhir'] = pd.to_datetime(Kegiatan['TanggalAkhir'], format='%d/%m/%Y', errors='coerce')

st.title('Surat Tugas Kanwil X')
col1, col2 = st.columns([0.8, 0.2]) # Adjust ratios as needed
with col1:
    st.write('Displaying the DataFrame:')
with col2:
    st.markdown(f"[Go to Spreadsheet](https://docs.google.com/spreadsheets/d/1SORCi_jXxEN-HSXWBOjX19FY8a6FzegPSAPbuh5k1sI/)", unsafe_allow_html=True)

# --- Sidebar for filters ---
st.sidebar.header('Filter Options')

# Filter Status Surat Tugas
options = ["All","Selesai","Belum Selesai"]
Status_surat = st.sidebar.segmented_control(
    "Status Surat Tugas", options, selection_mode="single", default="All"
)

# Month filter
months = {
    'All': 0,
    'January': 1, 'February': 2, 'March': 3, 'April': 4, 'May': 5, 'June': 6,
    'July': 7, 'August': 8, 'September': 9, 'October': 10, 'November': 11, 'December': 12
}

# Get current month name
current_month_name = datetime.now().strftime('%B')
# Find the index of the current month in the months dictionary keys
try:
    default_month_index = list(months.keys()).index(current_month_name)
except ValueError:
    default_month_index = months[current_month_name] # Default to 'All' if current month not found

selected_month_name = st.sidebar.selectbox('Select Month', list(months.keys()), index=default_month_index)
selected_month_num = months[selected_month_name]

# Name filter
search_name = st.sidebar.text_input('Search by Employee Name', '')

# Apply filters
filtered_Kegiatan = Kegiatan.copy()

if Status_surat != "All":
    filtered_Kegiatan = filtered_Kegiatan[filtered_Kegiatan['Keterangan'] == Status_surat]

# Filter by month, excluding NaT values
if selected_month_num != 0:
    filtered_Kegiatan = filtered_Kegiatan[filtered_Kegiatan['TanggalAwal'].dt.month == selected_month_num]

if search_name:
    filtered_Kegiatan = filtered_Kegiatan[filtered_Kegiatan['Karyawan'].str.contains(search_name, case=False, na=False)]

# Drop 'TanggalAwal' and 'TanggalAkhir' columns before displaying
filtered_Kegiatan = filtered_Kegiatan.drop(columns=['TanggalAwal', 'TanggalAkhir'], errors='ignore')

st.dataframe(filtered_Kegiatan)

def download_ST(df):
  Karyawan = pd.read_excel(url, sheet_name='Karyawan_ST')
  doc = DocxTemplate(template_file)
  ListST = list(df["NoSurat"].unique())
  list_dicts = []

  for i in ListST:
    KegiatanPerST = Kegiatan[Kegiatan['NoSurat'] == int(i)]
    KaryawanPerST = Karyawan[Karyawan['NoSurat'] == int(i)]
    nama_karyawan = KaryawanPerST[["Nama","NIK","UnitKerja"]]
    kegiatan = KegiatanPerST[["JangkaWaktu","Kegiatan","Lokasi"]]

    nama_dict = nama_karyawan.to_dict(orient='records')
    kegiatan_dict = kegiatan.to_dict(orient='records')

    context = {
        "NoSurat": f"{i}",
        "table1": nama_dict,
        "table2": kegiatan_dict,
            }
    list_dicts.append(context)

  temp_files = []
  for i, entry in enumerate(list_dicts):
      tpl = DocxTemplate(template_file)
      tpl.render(entry)
      temp_name = f'temp_{i}.docx'
      tpl.save(temp_name)
      temp_files.append(temp_name)

  # 3. MERGE LOGIC: Use the FIRST rendered file as the master
  if list_dicts: # Only proceed if there are documents to merge
      master_doc = Document(temp_files[0])

      # Helper function to append one docx to another with a page break
      def append_docx(master, sub_doc_path):
          sub_doc = Document(sub_doc_path)
          for element in sub_doc.element.body:
              master.element.body.append(element)

      # 4. Append the rest (starting from the second file)
      for temp_file in temp_files[1:]:          append_docx(master_doc, temp_file)

      # 5. Save and Download
      master_doc.save(output_file)

      # Download the file
      with open(output_file, "rb") as file:
          btn = st.download_button(
                  label="Download Surat Tugas",
                  data=file,
                  file_name=output_file,
                  mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
              )

      # Cleanup temporary files
      for f in temp_files:
          os.remove(f)
  else:
      st.warning("No documents generated for merging. Check your filters.")

# Add a button to trigger the download
if st.button('Generate & Download Surat Tugas'):
    download_ST(filtered_Kegiatan)
