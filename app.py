import streamlit as st
import google.generativeai as genai
import docx
import fitz  # PyMuPDF
import io
import re
from docx.enum.text import WD_COLOR_INDEX
import pandas as pd
from docx.shared import Pt
import zipfile

# --- Konfigurasi Awal Halaman ---
st.set_page_config(
    page_title="Proofreader Bahasa Indonesia",
    layout="wide"
)

# --- Header Aplikasi ---
try:
    # Ganti "Logo_IFG-removebg-preview.png" dengan nama file logo Anda
    st.image("Logo_IFG-removebg-preview.png", width=250)
except Exception:
    st.warning("Logo tidak ditemukan. Pastikan file logo ada di direktori yang sama.")

st.markdown("<h1 style='text-align: center;'>Website Proofreader COE Divisi SKAI IFG</h1>", unsafe_allow_html=True)

# --- BAGIAN 1: PROOFREAD ---
st.divider() # Garis pemisah
st.header("1. Proofread Dokumen")

# --- Inisialisasi Session State ---
if 'analysis_results' not in st.session_state:
    st.session_state.analysis_results = None

# --- Konfigurasi API Key Google ---
try:
    api_key = st.secrets["GOOGLE_API_KEY"]
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-1.5-flash-latest')
except KeyError:
    st.error("Google API Key belum diatur. Harap atur di Streamlit Secrets.")
    st.stop()
except Exception as e:
    st.error(f"Terjadi masalah saat mengkonfigurasi Google AI: {e}")
    st.stop()

# --- FUNGSI-FUNGSI UTAMA ---

def extract_text_with_pages(uploaded_file):
    """Mengekstrak teks dari file PDF atau DOCX."""
    pages_content = []
    file_extension = uploaded_file.name.split('.')[-1].lower()
    file_bytes = uploaded_file.getvalue()

    if file_extension == 'pdf':
        try:
            pdf_document = fitz.open(stream=file_bytes, filetype="pdf")
            for page_num, page in enumerate(pdf_document):
                pages_content.append({"halaman": page_num + 1, "teks": page.get_text()})
            pdf_document.close()
        except Exception as e:
            st.error(f"Gagal membaca file PDF: {e}")
            return None
    elif file_extension == 'docx':
        try:
            doc = docx.Document(io.BytesIO(file_bytes))
            full_text = "\n".join([para.text for para in doc.paragraphs])
            pages_content.append({"halaman": 1, "teks": full_text})
        except Exception as e:
            st.error(f"Gagal membaca file DOCX: {e}")
            return None
    else:
        st.error("Format file tidak didukung. Harap unggah .pdf atau .docx")
        return None
    return pages_content

def proofread_with_gemini(text_to_check):
    """Mengirim teks ke Gemini untuk proofreading dan mem-parsing hasilnya."""
    if not text_to_check or text_to_check.isspace():
        return []
    
    prompt = f"""
    Anda adalah seorang auditor dan ahli bahasa Indonesia yang sangat teliti.
    Tugas Anda adalah melakukan proofread pada teks berikut. Fokus pada:
    1. Memperbaiki kesalahan ketik (typo).
    2. Memastikan semua kata sesuai dengan Kamus Besar Bahasa Indonesia (KBBI).
    3. Memperbaiki kesalahan tata bahasa sederhana dan ejaan agar sesuai dengan Pedoman Umum Ejaan Bahasa Indonesia (PUEBI).
    4. Jika ada yang bahasa inggris, tolong di italic
    5. Nama-nama yang diberi ini pastikan benar juga "Yullyan, I Made Suandi Putra, Laila Fajriani, Hari Sundoro, Bakhas Nasrani Diso, Rizky Ananda Putra, Wirawan Arief Nugroho, Lelya Novita Kusumawati, Ryani Ariesti Syafitri, Darmo Saputro Wibowo, Lucky Parwitasari, Handarudigdaya Jalanidhi Kuncaratrah, Fajar Setianto, Jaka Tirtana Hanafiah,  Muhammad Rosyid Ridho Muttaqien, Octovian Abrianto, Deny Sjahbani, Jihan Abigail, Winda Anggraini, Fadian Dwiantara, Aliya Anindhita Rachman"
    6. Fontnya arial dan jangan diganti. Khusus untuk judul paling atas, itu font sizenya 12 dan bodynya selalu 11
    7. Khusus "Indonesia Financial Group (IFG)", meskipun bahasa inggris, tidak perlu di italic
    8. Bila ada bahasa yang lebih bagus, tolong berikan saran dan diberi warna highlight yang berbeda selain kuning
    9. Pada bagian penutup, yang mulai dari "Jakarta, 4 September 2025" hingga "Kepala Divisi Satuan Kerja Audit Internal", tidak perlu dicek
    10. Pada bagian nomor surat, itu juga tidak perlu di cek
    11. Kalau ada kata-kata yang tidak sesuai KBBI dan PUEBI, tolong jangan highlight semua kalimatnya, tapi cukup highlight kata-kata yang salah serta perbaiki kata-kata itu aja, jangan perbaiki semua kalimatnya
    12. Ketika Anda perbaiki, fontnya pastikan Arial dengan ukuran 11 juga
    13. Kalau ada kata-kata dalam bahasa inggris, tolong jangan Anda sarankan untuk ganti ke bahasa indonesia, cukup Anda sarankan untuk italic
    14. Pada kalimat "Indonesia Financial Group", jika terdapat kata typo "Finansial", tolong Anda sarankan untuk ganti ke "Financial"
    15. Yang benar adalah "Satuan Kerja Audit Internal", bukan "Satuan Pengendali Internal Audit"
    16. Jika terdapat kata "reviu", biarkan itu sebagai benar
    17. Kalau ada kata "IM", "ST", "SKAI", "IFG", "TV (Angka Romawi)", "RKAT", dan "RKAP", itu tidak perlu ditandai sebagai salah dan tidak perlu disarankan untuk italic / bold / underline
    18. Kalau ada kata "email", biarkan itu sebagai benar tapi disarankan italic saja

    PENTING: Berikan hasil dalam format yang SANGAT KETAT. Untuk setiap kesalahan, gunakan format:
    [SALAH] kata atau frasa yang salah -> [BENAR] kata atau frasa perbaikan -> [KALIMAT] kalimat lengkap asli tempat kesalahan ditemukan

    Contoh:
    [SALAH] dikarenakan -> [BENAR] karena -> [KALIMAT] Hal itu terjadi dikarenakan kelalaian petugas.

    Jika tidak ada kesalahan sama sekali, kembalikan teks: "TIDAK ADA KESALAHAN"

    Berikut adalah teks yang harus Anda periksa:
    ---
    {text_to_check}
    """
    try:
        response = model.generate_content(prompt)
        pattern = re.compile(r"\[SALAH\]\s*(.*?)\s*->\s*\[BENAR\]\s*(.*?)\s*->\s*\[KALIMAT\]\s*(.*?)\s*(\n|$)", re.IGNORECASE)
        found_errors = pattern.findall(response.text)
        return [{"salah": salah.strip(), "benar": benar.strip(), "kalimat": kalimat.strip()} for salah, benar, kalimat, _ in found_errors]
    except Exception as e:
        st.error(f"Terjadi kesalahan saat menghubungi AI: {e}")
        return []

def generate_revised_docx(file_bytes, errors):
    """
    Membuat dokumen .docx dengan semua kesalahan yang sudah diperbaiki
    SAMBIL MEMPERTAHANKAN FONT ASLI (Arial 11 untuk body dan Arial 12 untuk judul).
    """
    doc = docx.Document(io.BytesIO(file_bytes))

    for error in reversed(errors):
        salah = error["Kata/Frasa Salah"]
        benar = error["Perbaikan Sesuai KBBI"]
        for para in doc.paragraphs:
            if salah in para.text:
                original_font = None
                if para.runs:
                    original_font = para.runs[0].font
                
                current_text = para.text
                para.text = current_text.replace(salah, benar, 1)

                if original_font:
                    for run in para.runs:
                        font = run.font
                        font.name = original_font.name
                        font.size = original_font.size
                        font.bold = original_font.bold
                        font.italic = original_font.italic
                        font.underline = original_font.underline
                        font.color.rgb = original_font.color.rgb

    output_buffer = io.BytesIO()
    doc.save(output_buffer)
    return output_buffer.getvalue()

def generate_highlighted_docx(file_bytes, errors):
    """Membuat dokumen .docx dengan semua kesalahan yang di-highlight."""
    doc = docx.Document(io.BytesIO(file_bytes))
    unique_salah = set(error["Kata/Frasa Salah"] for error in errors)
    for para in doc.paragraphs:
        for term in unique_salah:
            if term in para.text:
                for run in para.runs:
                    if term in run.text:
                        run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    output_buffer = io.BytesIO()
    doc.save(output_buffer)
    return output_buffer.getvalue()

def create_zip_archive(revised_data, highlighted_data, original_filename):
    """Menggabungkan dua file DOCX ke dalam satu arsip ZIP di memori."""
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.writestr(f"revisi_{original_filename}", revised_data)
        zip_file.writestr(f"highlight_{original_filename}", highlighted_data)
    return zip_buffer.getvalue()

# --- ANTARMUKA STREAMLIT UNTUK BAGIAN 1 ---
uploaded_file = st.file_uploader(
    "Unggah dokumen (PDF/DOCX)",
    type=['pdf', 'docx'],
    help="File yang diunggah akan dianalisis untuk menemukan kesalahan ejaan dan ketik."
)

if uploaded_file is not None:
    st.info(f"File yang diunggah: **{uploaded_file.name}**")

    if st.button("Mulai Analisis", type="primary", use_container_width=True):
        with st.spinner("Membaca dan menganalisis dokumen..."):
            document_pages = extract_text_with_pages(uploaded_file)
            
            if document_pages:
                all_errors = []
                progress_bar = st.progress(0, text="Menganalisis teks dengan AI...")
                
                for i, page in enumerate(document_pages):
                    progress_text = f"Menganalisis Bagian {i + 1}/{len(document_pages)}..."
                    progress_bar.progress((i + 1) / len(document_pages), text=progress_text)
                    found_errors_on_page = proofread_with_gemini(page['teks'])
                    
                    for error in found_errors_on_page:
                        all_errors.append({
                            "Kata/Frasa Salah": error['salah'],
                            "Perbaikan Sesuai KBBI": error['benar'],
                            "Pada Kalimat": error['kalimat'],
                            "Ditemukan di Halaman": page['halaman']
                        })
                progress_bar.empty()
                st.session_state.analysis_results = all_errors

if st.session_state.analysis_results is not None:
    all_errors = st.session_state.analysis_results

    if not all_errors:
        st.success("Tidak ada kesalahan ejaan atau ketik yang ditemukan dalam dokumen.")
    else:
        st.warning(f"Ditemukan **{len(all_errors)}** potensi kesalahan dalam dokumen.")
        
        df_errors = pd.DataFrame(all_errors)
        st.dataframe(df_errors, use_container_width=True)
        
        st.subheader("Unduh Hasil")
        
        if uploaded_file and uploaded_file.name.endswith('.docx'):
            with st.spinner("Mempersiapkan semua file unduhan..."):
                revised_docx_data = generate_revised_docx(uploaded_file.getvalue(), all_errors)
                highlighted_docx_data = generate_highlighted_docx(uploaded_file.getvalue(), all_errors)

            col1, col2, col3 = st.columns(3)

            with col1:
                st.download_button(
                    label="Unduh Direvisi (.docx)",
                    data=revised_docx_data,
                    file_name=f"revisi_{uploaded_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            
            with col2:
                st.download_button(
                    label="Unduh Highlight (.docx)",
                    data=highlighted_docx_data,
                    file_name=f"highlight_{uploaded_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            
            with col3:
                zip_data = create_zip_archive(revised_docx_data, highlighted_docx_data, uploaded_file.name)
                st.download_button(
                    label="Unduh Semua (.zip)",
                    data=zip_data,
                    file_name=f"hasil_proofread_{uploaded_file.name.split('.')[0]}.zip",
                    mime="application/zip",
                    use_container_width=True
                )
    st.warning("Hasilnya masih bisa salah, tolong dicek ulang lagi.")

# --- BAGIAN 2: BANDINGKAN DOKUMEN ---
st.divider()
st.header("2. Bandingkan Dokumen")

# Tambahan import yang mungkin diperlukan untuk bagian ini
import difflib
from docx import Document
from docx.shared import Pt, Inches

# --- Fungsi-fungsi Helper untuk Bagian 2 ---

def extract_paragraphs(docx_file):
    """Membaca file DOCX yang diunggah dan mengembalikan isinya sebagai daftar paragraf."""
    try:
        source_stream = io.BytesIO(docx_file.getvalue())
        doc = docx.Document(source_stream)
        return [p.text for p in doc.paragraphs if p.text.strip() != ""]
    except Exception as e:
        st.error(f"Gagal membaca file {docx_file.name}: {e}")
        return []

def get_revision_confidence(original_sentence, revised_sentence):
    """Meminta model AI untuk memberikan skor keyakinan bahwa revisi lebih baik."""
    if original_sentence == revised_sentence:
        return 100
    prompt = f"""
    Bandingkan dua kalimat ini berdasarkan PUEBI dan KBBI (Pastikan sesuai dengan S-P-O-K pada PUEBI).
    Kalimat Asli: "{original_sentence}"
    Kalimat Revisi: "{revised_sentence}"

    Beri skor keyakinan (angka 0-100) bahwa revisi tersebut adalah perbaikan yang benar dan diperlukan.
    PENTING: Jawab HANYA dengan ANGKA.
    """
    try:
        response = model.generate_content(prompt)
        confidence_str = ''.join(filter(str.isdigit, response.text))
        return int(confidence_str) if confidence_str else "N/A"
    except Exception:
        return "N/A"

def find_word_diff(original_para, revised_para):
    """Menemukan kata-kata yang berbeda dan mengembalikannya sebagai list."""
    matcher = difflib.SequenceMatcher(None, original_para.split(), revised_para.split())
    diffs = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'replace' or tag == 'insert':
            # Mengambil kata/frasa yang direvisi dari kalimat baru
            revised_text = " ".join(revised_para.split()[j1:j2])
            # Mengambil kata/frasa asli
            original_text = " ".join(original_para.split()[i1:i2])
            # Format: "Teks Asli -> Teks Revisi" untuk kejelasan
            diffs.append(f'"{original_text}" → "{revised_text}"')
            
    return diffs if diffs else ["Perubahan Minor (tanda baca, spasi, dll.)"]

def create_comparison_docx(df):
    """Membuat file DOCX dari DataFrame hasil perbandingan dengan format font yang diinginkan."""
    doc = Document()
    
    # Menambahkan Judul Utama Dokumen
    title = doc.add_heading('Hasil Perbandingan Dokumen', level=1)
    for run in title.runs:
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        run.bold = True
    doc.add_paragraph() # Spasi

    # Menambahkan Tabel
    table = doc.add_table(rows=1, cols=len(df.columns))
    table.style = 'Table Grid'
    table.autofit = True

    # Menambahkan Header Tabel dengan Font Arial Bold
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(df.columns):
        p = hdr_cells[i].paragraphs[0]
        p.text = col_name
        for run in p.runs:
            run.font.name = 'Arial'
            run.font.bold = True
            run.font.size = Pt(11)

    # Menambahkan Isi Tabel dengan Font Arial 11
    for index, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, cell_value in enumerate(row):
            p = row_cells[i].paragraphs[0]
            p.text = str(cell_value)
            for run in p.runs:
                run.font.name = 'Arial'
                run.font.size = Pt(11)

    output_buffer = io.BytesIO()
    doc.save(output_buffer)
    return output_buffer.getvalue()

# --- Antarmuka Streamlit untuk Bagian 2 ---

col1, col2 = st.columns(2)
with col1:
    original_file = st.file_uploader("Unggah Dokumen Asli", type=['docx'], key="original_doc")
with col2:
    proofread_file = st.file_uploader("Unggah Dokumen Hasil Proofread", type=['docx'], key="proofread_doc")

if original_file is not None and proofread_file is not None:
    if st.button("Bandingkan Dokumen", use_container_width=True, type="primary"):
        with st.spinner("Mengekstrak teks dan membandingkan dokumen..."):
            original_paras = extract_paragraphs(original_file)
            revised_paras = extract_paragraphs(proofread_file)
            comparison_results = []
            matcher = difflib.SequenceMatcher(None, original_paras, revised_paras)

            for tag, i1, i2, j1, j2 in matcher.get_opcodes():
                if tag == 'replace':
                    # Ambil pasangan paragraf yang berubah
                    original_para = " ".join(original_paras[i1:i2])
                    revised_para = " ".join(revised_paras[j1:j2])

                    # Dapatkan skor keyakinan untuk keseluruhan paragraf (cukup 1x per paragraf)
                    confidence = get_statistical_confidence(original_para, revised_para, tool)
                    
                    # Dapatkan LIST kata-kata yang berbeda
                    word_diff_list = find_word_diff(original_para, revised_para)
                    
                    # Buat satu baris baru UNTUK SETIAP kata yang berbeda
                    for diff in word_diff_list:
                        comparison_results.append({
                            "Kalimat Awal": original_para,
                            "Kalimat Revisi": revised_para,
                            "Detail Perubahan": diff, # <-- Kolom ini sekarang berisi satu perubahan spesifik
                            "Skor Perbaikan (%)": confidence # <-- Skor tetap sama untuk semua perubahan dalam paragraf ini
                        })
            
            # Ganti nama kolom di DataFrame agar sesuai
            df = pd.DataFrame(comparison_results)
            if not df.empty:
                df = df.rename(columns={"Kata yang Direvisi": "Detail Perubahan"})

            st.session_state.comparison_results = df

# Menampilkan hasil jika ada di session state
if 'comparison_results' in st.session_state and not st.session_state.comparison_results.empty:
    df_comparison = st.session_state.comparison_results
    st.success(f"Perbandingan selesai. Ditemukan {len(df_comparison)} paragraf yang direvisi.")
    st.dataframe(df_comparison, use_container_width=True)

    # Menambahkan tombol download untuk hasil perbandingan
    docx_data = create_comparison_docx(df_comparison)
    st.download_button(
        label="Unduh Hasil Perbandingan (.docx)",
        data=docx_data,
        file_name=f"perbandingan_{original_file.name}",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
    st.info("Catatan: Skor 'Keyakinan Model' adalah estimasi dari AI dan bukan metrik statistik absolut.")

elif 'comparison_results' in st.session_state:
     st.info("Tidak ditemukan perbedaan signifikan antar paragraf di kedua dokumen.")

