
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
from docx import Document
import os
from docxcompose.composer import Composer

# --- Configuration --- #
template_file = 'ST26-template.docx'
output_file = 'combined_output_2026.docx'
url = "https://docs.google.com/spreadsheets/d/1SORCi_jXxEN-HSXWBOjX19FY8a6FzegPSAPbuh5k1sI/export?format=xlsx"

# --- Data Loading and Preprocessing --- #
def load_data():
    kegiatan_df = pd.read_excel(url, sheet_name='Kegiatan', header=2)
    karyawan_df = pd.read_excel(url, sheet_name='Karyawan_ST')
    kegiatan_df['TanggalAwal'] = pd.to_datetime(kegiatan_df['TanggalAwal'], format='%d/%m/%Y', errors='coerce')
    kegiatan_df['TanggalAkhir'] = pd.to_datetime(kegiatan_df['TanggalAkhir'], format='%d/%m/%Y', errors='coerce')
    return kegiatan_df, karyawan_df


Kegiatan, Karyawan = load_data() # Load data (will use cache unless cleared)

# --- Streamlit UI Components --- #
st.title('Surat Tugas Kanwil X')
col1, col2 = st.columns([0.8, 0.2]) # Add a third column for the refresh button
with col1:
  st.write('Displaying the DataFrame:')
with col2:
  st.markdown(f"[Go to Spreadsheet](https://docs.google.com/spreadsheets/d/1SORCi_jXxEN-HSXWBOjX19FY8a6FzegPSAPbuh5k1sI/)", unsafe_allow_html=True)
# with col3:
#   if st.button('Refresh', type="tertiary"):
#     # Invalidate cache if it was used, then reload
#     st.cache_data.clear()
#     Kegiatan, Karyawan = load_data()
#   else:
#     Kegiatan, Karyawan = load_data()

st.sidebar.header('Filter Options')

options = ["All", "Selesai", "Belum Selesai"]
Status_surat = st.sidebar.segmented_control(
    "Status Surat Tugas", options, selection_mode="single", default="All"
)

months = {
    'All': 0,
    'January': 1, 'February': 2, 'March': 3, 'April': 4, 'May': 5, 'June': 6,
    'July': 7, 'August': 8, 'September': 9, 'October': 10, 'November': 11, 'December': 12
}

current_month_name = datetime.now().strftime('%B')
try:
    default_month_index = list(months.keys()).index(current_month_name)
except ValueError:
    default_month_index = months[current_month_name]

selected_month_name = st.sidebar.selectbox('Select Month', list(months.keys()), index=default_month_index)
selected_month_num = months[selected_month_name]

search_name = st.sidebar.text_input('Search by Employee Name', '')

filtered_Kegiatan = Kegiatan.copy()

if Status_surat != "All":
    filtered_Kegiatan = filtered_Kegiatan[filtered_Kegiatan['Keterangan'] == Status_surat]

if selected_month_num != 0:
    filtered_Kegiatan = filtered_Kegiatan[filtered_Kegiatan['TanggalAwal'].dt.month == selected_month_num]

if search_name:
    filtered_Kegiatan = filtered_Kegiatan[filtered_Kegiatan['Karyawan'].str.contains(search_name, case=False, na=False)]

filtered_Kegiatan_display = filtered_Kegiatan.drop(columns=['TanggalAwal', 'TanggalAkhir'], errors='ignore')
st.dataframe(filtered_Kegiatan_display, use_container_width=True)

# --- Document Generation Logic --- #
def merge_docx_files(master_path, files_to_append, output_path):
    master = Document(master_path)
    master.add_page_break()
    composer = Composer(master)

    for file_path in files_to_append:
        doc_to_append = Document(file_path)
        doc_to_append.add_page_break()
        composer.append(doc_to_append)

    composer.save(output_path)

def download_ST(df, Kegiatan_all, Karyawan_all):
    ListST = list(df["NoSurat"].unique())
    list_dicts = []

    for i in ListST:
        KegiatanPerST = Kegiatan_all[Kegiatan_all['NoSurat'] == int(i)]
        KaryawanPerST = Karyawan_all[Karyawan_all['NoSurat'] == int(i)]
        nama_karyawan = KaryawanPerST[["Nama", "NIK", "UnitKerja"]]
        kegiatan = KegiatanPerST[["JangkaWaktu", "Kegiatan", "Lokasi"]]

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

    if list_dicts:
        merge_docx_files(temp_files[0], temp_files[1:], output_file)
        with open(output_file, "rb") as file:
            st.download_button(
                label="Download Surat Tugas",
                data=file,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.warning("No documents generated for merging. Check your filters.")

    for f in temp_files:
        os.remove(f)


if st.button('Generate Surat Tugas'):
    download_ST(filtered_Kegiatan, Kegiatan, Karyawan)
st.markdown("*Untuk tanda tangan, pastikan hanya surat tugas yang belum selesai yang tampil di dataframe*")
