import streamlit as st
import google.generativeai as genai
import docx
import fitz  # PyMuPDF
import io
import re
from docx.enum.text import WD_COLOR_INDEX
import pandas as pd # <-- Import pandas untuk st.dataframe

# --- Konfigurasi Awal ---
st.set_page_config(
    page_title="Proofreader Bahasa Indonesia",
    layout="wide"
)

# --- HEADER APLIKASI ---
try:
    # Logo diperbesar dengan mengubah width dari 150 menjadi 250
    st.image("Logo_IFG-removebg-preview.png", width=250)
except Exception as e:
    st.warning("Logo tidak ditemukan. Pastikan file 'logo.png' ada di direktori yang sama.")

st.markdown("<h1 style='text-align: center;'>Website Proofreader Departemen COE Divisi SKAI IFG</h1>", unsafe_allow_html=True)
st.caption("Unggah dokumen (PDF/DOCX) untuk mendeteksi kesalahan ketik, ejaan, dan tata bahasa.")

# --- Konfigurasi API Key (LEBIH AMAN) ---
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
    """Mengekstrak teks dari file DOCX."""
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

    # --- PROMPT DIMODIFIKASI ---
    # Menambahkan instruksi untuk menyertakan kalimat asli
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
    13. Jika pada  "Indonesia Financial Group" tertulis "Financial", tolong perbaiki menjadi "Financial" dan jangan italic
    14. Pastikan yang benar adalah "Satuan Kerja Audit Internal"

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
        # --- REGEX DIMODIFIKASI ---
        # Menambahkan grup ketiga untuk menangkap [KALIMAT]
        pattern = re.compile(r"\[SALAH\]\s*(.*?)\s*->\s*\[BENAR\]\s*(.*?)\s*->\s*\[KALIMAT\]\s*(.*?)\s*(\n|$)", re.IGNORECASE)
        found_errors = pattern.findall(response.text)
        
        # --- RETURN DIMODIFIKASI ---
        # Menambahkan 'kalimat' ke dalam dictionary
        return [{"salah": salah.strip(), "benar": benar.strip(), "kalimat": kalimat.strip()} for salah, benar, kalimat, _ in found_errors]
    except Exception as e:
        st.error(f"Terjadi kesalahan saat menghubungi AI: {e}")
        return []

# --- FUNGSI UNTUK MEMBUAT OUTPUT DOKUMEN ---

def generate_revised_docx(file_bytes, errors):
    """Membuat dokumen .docx dengan semua kesalahan yang sudah diperbaiki."""
    doc = docx.Document(io.BytesIO(file_bytes))
    # Iterasi terbalik agar penggantian tidak mengganggu indeks kata lain
    for error in reversed(errors):
        salah = error["Kata/Frasa Salah"]
        benar = error["Perbaikan Sesuai KBBI"]
        for para in doc.paragraphs:
            if salah in para.text:
                # Mengganti hanya satu kali untuk menghindari penggantian berlebih jika kata sama
                para.text = para.text.replace(salah, benar, 1)
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

# --- ANTARMUKA STREAMLIT ---
uploaded_file = st.file_uploader(
    "Pilih file PDF atau DOCX",
    type=['pdf', 'docx'],
    help="File yang diunggah akan dianalisis untuk menemukan kesalahan ejaan dan ketik."
)

if uploaded_file is not None:
    st.info(f"File yang diunggah: **{uploaded_file.name}**")

    if st.button("Mulai Analisis", type="primary", use_container_width=True):
        with st.spinner("Membaca dan mengekstrak teks dari dokumen..."):
            document_pages = extract_text_with_pages(uploaded_file)
        
        if document_pages:
            st.success(f"Ekstraksi teks berhasil. Memproses {len(document_pages)} bagian teks.")
            
            all_errors = []
            progress_bar = st.progress(0, text="Menganalisis teks dengan AI...")
            
            for i, page in enumerate(document_pages):
                progress_text = f"Menganalisis Bagian {i + 1}/{len(document_pages)}..."
                progress_bar.progress((i + 1) / len(document_pages), text=progress_text)
                found_errors_on_page = proofread_with_gemini(page['teks'])
                
                # --- PROSES HASIL DIMODIFIKASI ---
                # Menambahkan 'kalimat' ke dalam list all_errors
                for error in found_errors_on_page:
                    all_errors.append({
                        "Kata/Frasa Salah": error['salah'],
                        "Perbaikan Sesuai KBBI": error['benar'],
                        "Pada Kalimat": error['kalimat'], # <-- KOLOM BARU DITAMBAHKAN DI SINI
                        "Ditemukan di Halaman": page['halaman']
                    })

            progress_bar.empty()

            if not all_errors:
                st.success("Tidak ada kesalahan ejaan atau ketik yang ditemukan dalam dokumen.")
            else:
                st.warning(f"Ditemukan **{len(all_errors)}** potensi kesalahan dalam dokumen.")
                
                # Menggunakan st.dataframe untuk tampilan yang lebih baik
                df_errors = pd.DataFrame(all_errors)
                st.dataframe(df_errors, use_container_width=True)
                
                # --- BAGIAN DOWNLOAD ---
                st.subheader("Download Hasil")
                col1, col2 = st.columns(2)

                with col1:
                    if uploaded_file.name.endswith('.docx'):
                        with st.spinner("Membuat file revisi..."):
                            revised_docx_data = generate_revised_docx(uploaded_file.getvalue(), all_errors)
                            st.download_button(
                                label="Unduh DOCX (Sudah Direvisi)",
                                data=revised_docx_data,
                                file_name=f"revisi_{uploaded_file.name}",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )
                
                with col2:
                    if uploaded_file.name.endswith('.docx'):
                        with st.spinner("Membuat file highlight..."):
                            highlighted_docx_data = generate_highlighted_docx(uploaded_file.getvalue(), all_errors)
                            st.download_button(
                                label="Unduh DOCX (Highlight Kesalahan)",
                                data=highlighted_docx_data,
                                file_name=f"highlight_{uploaded_file.name}",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )

