from io import BytesIO
import os
import time
from typing import List
import warnings

import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor
from google import genai
from google.genai import types
from pydantic import BaseModel, Field
import streamlit as st

# Sembunyikan warning
warnings.filterwarnings("ignore")

# ---------------------------------------------------------
# 1. KONFIGURASI HALAMAN STREAMLIT
# ---------------------------------------------------------
st.set_page_config(
    page_title="Perangkat Ajar The Rufidz Indonesia",
    page_icon="📖",
    layout="wide",
)

# Custom CSS Tampilan Pesantren
st.markdown(
    """
    <style>
    .main-title {
        color: #0A5C36;
        text-align: center;
        font-weight: bold;
        margin-bottom: 0px;
    }
    .sub-title {
        text-align: center;
        color: #555555;
        font-style: italic;
        margin-bottom: 30px;
    }
    .stButton>button {
        background-color: #0A5C36;
        color: white;
        font-weight: bold;
        border-radius: 8px;
        height: 3em;
        width: 100%;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# API KEY GEMINI & INISIALISASI CLIENT
API_KEY = st.secrets.get(
    "GEMINI_API_KEY",
    os.getenv(
        "GEMINI_API_KEY",
        "AQ.Ab8RN6KtsISk3rZk-7-3eTNfxlMRtXGmCunOnjU6VZgMbMbdUg",
    ),
)

client = genai.Client(api_key=API_KEY)


# ---------------------------------------------------------
# 2. SCHEMA PYDANTIC DATA
# ---------------------------------------------------------
class SilabusItem(BaseModel):
  pertemuan_ke: int
  bab_fasal: str
  perkiraan_halaman: str
  metode: str
  capaian_indikator: str


class RPPDetail(BaseModel):
  tujuan_pembelajaran: List[str]
  target_bacaan: str
  langkah_pembukaan: str
  langkah_inti: str
  langkah_penutup: str
  asesmen_penilaian: List[str]


class PerangkatAjarSalaf(BaseModel):
  nama_kitab: str
  fan_ilmu: str
  tingkat_rekomendasi: str
  total_pertemuan: int
  silabus: List[SilabusItem]
  sample_rpp_pertemuan_1: RPPDetail


# ---------------------------------------------------------
# 3. FUNGSI GENERATE AI (GEMINI)
# ---------------------------------------------------------
def generate_ai(nama_kitab: str) -> PerangkatAjarSalaf:
  system_prompt = (
      "Kamu adalah pakar kurikulum pesantren salaf tradisional Indonesia. "
      "Tugasmu menyusun Perangkat Ajar (Silabus & RPP) berbasis nama kitab"
      " salaf. "
      "Gunakan istilah khas pesantren salaf seperti Bandongan, Sorogan, Lalaran,"
      " Mudzakarah, dan Makna Gandul. "
      "Pastikan output STRICT mengikuti skema JSON."
  )

  max_retries = 5
  for attempt in range(1, max_retries + 1):
    try:
      response = client.models.generate_content(
          model="gemini-3.6-flash",
          contents=f"Buatkan silabus dan RPP lengkap untuk kitab: {nama_kitab}",
          config=types.GenerateContentConfig(
              system_instruction=system_prompt,
              response_mime_type="application/json",
              response_schema=PerangkatAjarSalaf,
              temperature=0.2,
          ),
      )
      return PerangkatAjarSalaf.model_validate_json(response.text)
    except Exception as e:
      if "503" in str(e) or "UNAVAILABLE" in str(e):
        time.sleep(4)
      else:
        raise e
  raise Exception("Gagal terhubung ke server AI. Silakan coba kembali.")


# ---------------------------------------------------------
# 4. FUNGSI EKSPOR KE WORD (BYTES)
# ---------------------------------------------------------
def create_docx_bytes(data: PerangkatAjarSalaf) -> BytesIO:
  doc = docx.Document()
  PRIMARY_RGB = RGBColor(10, 92, 54)

  # Kop Header
  p_kop = doc.add_paragraph()
  p_kop.alignment = WD_ALIGN_PARAGRAPH.CENTER
  p_kop.paragraph_format.space_after = Pt(2)

  r_brand = p_kop.add_run("PERANGKAT AJAR THE RUFIDZ INDONESIA\n")
  r_brand.font.size = Pt(16)
  r_brand.font.bold = True
  r_brand.font.color.rgb = PRIMARY_RGB

  r_sub = p_kop.add_run(
      "Sistem Kurikulum & Administrasi Pembelajaran Pesantren Salaf /"
      " Diniyah\n"
  )
  r_sub.font.size = Pt(10)
  r_sub.font.italic = True

  r_title = p_kop.add_run("SILABUS & RENCANA PELAKSANAAN PEMBELAJARAN (RPP)")
  r_title.font.size = Pt(12)
  r_title.font.bold = True

  # Identitas
  doc.add_heading("I. IDENTITAS PERANGKAT AJAR", level=2)
  p_id = doc.add_paragraph()
  p_id.add_run(f"Nama Kitab: {data.nama_kitab}\n").bold = True
  p_id.add_run(f"Fan Ilmu: {data.fan_ilmu}\n")
  p_id.add_run(f"Tingkat: {data.tingkat_rekomendasi}\n")
  p_id.add_run(f"Total Pertemuan: {data.total_pertemuan} Pertemuan")

  # Tabel Silabus
  doc.add_heading("II. SILABUS PEMBELAJARAN", level=2)
  table = doc.add_table(rows=1, cols=5)
  table.alignment = WD_TABLE_ALIGNMENT.CENTER

  hdr_cells = table.rows[0].cells
  headers = ["Ptm", "Bab / Fasal", "Halaman", "Metode", "Capaian Santri"]
  for i, h in enumerate(headers):
    hdr_cells[i].text = h
    p = hdr_cells[i].paragraphs[0]
    p.runs[0].font.bold = True

  for item in data.silabus:
    row_cells = table.add_row().cells
    row_cells[0].text = str(item.pertemuan_ke)
    row_cells[1].text = item.bab_fasal
    row_cells[2].text = item.perkiraan_halaman
    row_cells[3].text = item.metode
    row_cells[4].text = item.capaian_indikator

  # RPP
  doc.add_heading(
      "III. RENCANA PELAKSANAAN PEMBELAJARAN (RPP) SAMPLE", level=2
  )
  rpp = data.sample_rpp_pertemuan_1
  doc.add_paragraph(f"Target Bacaan: {rpp.target_bacaan}")
  doc.add_paragraph(f"Langkah Pembukaan: {rpp.langkah_pembukaan}")
  doc.add_paragraph(f"Langkah Inti: {rpp.langkah_inti}")
  doc.add_paragraph(f"Langkah Penutup: {rpp.langkah_penutup}")

  file_stream = BytesIO()
  doc.save(file_stream)
  file_stream.seek(0)
  return file_stream


# ---------------------------------------------------------
# 5. TAMPILAN UTAMA APLIKASI WEB
# ---------------------------------------------------------
st.markdown(
    "<h1 class='main-title'>PERANGKAT AJAR THE RUFIDZ INDONESIA</h1>",
    unsafe_allow_html=True,
)
st.markdown(
    "<p class='sub-title'>Generator Silabus & RPP Otomatis untuk Pesantren"
    " Salaf & Madrasah Diniyah</p>",
    unsafe_allow_html=True,
)

# Sidebar
st.sidebar.header("📌 Referensi Cepat Kitab")
preset_kitab = st.sidebar.selectbox(
    "Pilih Kitab Populer:",
    [
        "-- Pilih atau Ketik Manual --",
        "Safinatun Najah",
        "Matan Al-Jurumiyah",
        "Ta'lim Muta'allim",
        "Fathul Qarib",
        "Aqidatul Awam",
        "Taisirul Khollaq",
    ],
)

# Input Nama Kitab
default_val = (
    "" if preset_kitab == "-- Pilih atau Ketik Manual --" else preset_kitab
)
nama_kitab_input = st.text_input(
    "Nama Kitab Salaf:",
    value=default_val,
    placeholder="Contoh: Safinatun Najah / Jurumiyah",
)

# Tombol Generate
if st.button("🚀 Generate Silabus & RPP Otomatis"):
  if not nama_kitab_input.strip():
    st.warning("Silakan masukkan atau pilih nama kitab terlebih dahulu!")
  else:
    with st.spinner(
        f"Sedang mengunci bab & menyusun kurikulum untuk"
        f" '{nama_kitab_input}'..."
    ):
      try:
        # 1. AI Generate
        result = generate_ai(nama_kitab_input)
        st.session_state["data_result"] = result
        st.success("Berhasil di-generate oleh AI!")
      except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")

# Tampilkan Hasil Jika Ada
if "data_result" in st.session_state:
  data = st.session_state["data_result"]

  st.divider()
  st.subheader(f"📖 Identitas Kurikulum: {data.nama_kitab}")

  col1, col2, col3 = st.columns(3)
  col1.metric("Fan Ilmu", data.fan_ilmu)
  col2.metric("Tingkat Santri", data.tingkat_rekomendasi)
  col3.metric("Total Pertemuan", f"{data.total_pertemuan} Pertemuan")

  tab1, tab2 = st.tabs(["📋 Silabus Pembelajaran", "📝 Sample RPP Pertemuan 1"])

  with tab1:
    # Tampilkan Tabel Silabus
    table_data = []
    for s in data.silabus:
      table_data.append({
          "Pertemuan": s.pertemuan_ke,
          "Bab / Fasal": s.bab_fasal,
          "Perkiraan Hal.": s.perkiraan_halaman,
          "Metode Salaf": s.metode,
          "Capaian Santri": s.capaian_indikator,
      })
    st.dataframe(table_data, use_container_width=True)

  with tab2:
    rpp = data.sample_rpp_pertemuan_1
    st.markdown(f"**Target Bacaan:** {rpp.target_bacaan}")
    st.markdown(f"**Pembukaan:** {rpp.langkah_pembukaan}")
    st.markdown(f"**Langkah Inti (Bandongan/Pepeg):** {rpp.langkah_inti}")
    st.markdown(f"**Penutup & Sorogan:** {rpp.langkah_penutup}")

  st.divider()

  # Tombol Download Word File
  docx_file = create_docx_bytes(data)
  clean_name = "".join(
      c for c in data.nama_kitab if c.isalnum() or c in (" ", "_")
  ).rstrip()

  st.download_button(
      label="📥 Download File Word (.docx)",
      data=docx_file,
      file_name=f"Perangkat_Ajar_{clean_name.replace(' ', '_')}.docx",
      mime=(
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      ),
  )
