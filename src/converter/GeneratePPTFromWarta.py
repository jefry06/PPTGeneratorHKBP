from pptx import Presentation
import camelot
import re
import pdfplumber

from colorama import Fore, Style, init
import os

from src.converter.GeneratePPTFromTxt import create_caption, create_song_caption, create_sins_caption, create_epistel_caption, create_creed_caption, create_sermon_caption

from src.converter.GeneratePPTFromTxt import generate_song_lyric

def read_file(file_path, prs, language):
    opening_song = True
    tables = camelot.read_pdf(file_path, pages="1")
    # print(f"Jumlah tabel yang ditemukan: {len(tables)}")
    # tables[0].to_csv("f:/table.csv")
    # print(tables[0].df)

    df = tables[0].df
    # print(df)
    # print("-------------" + str(len(df)))
    # print(df.loc[0, 1])
    sequences = []
    pattern = r'^\d+'
    for index, row in df.iterrows():
        match = re.search(pattern, row[0])
        if match:
            split_row = row[1].split("\n")[0]
            sequence = split_row if split_row else row[1]
            match sequence:
                case sequence if "BE." in sequence.upper():
                    sequence = sequence.replace("BE.", "NH.") if language == "2" else sequence
                    create_song_caption(prs, sequence.strip(), opening_song, language)
                    opening_song = False
                    sequences.append({"type": "song", "value": sequence.strip()})
                case sequence if "VOTUM" in sequence.upper():
                    create_caption(prs, sequence.strip())
                    sequences.append({"type": "votum", "value": sequence.strip()})
                case sequence if "PATIK" in sequence.upper():
                    create_caption(prs, sequence.strip())
                    sequences.append({"type": "patik", "value": sequence.strip()})
                case sequence if "DOSA" in sequence.upper():
                    create_sins_caption(prs, sequence.strip())
                    sequences.append({"type": "dosa", "value": sequence.strip()})
                case sequence if "EPISTEL" in sequence.upper():
                    create_epistel_caption(prs, sequence.strip())
                    sequences.append({"type": "epistel", "value": sequence.strip()})
                case sequence if "HAPORSEAON" in sequence.upper():
                    create_creed_caption(prs, sequence.strip())
                    sequences.append({"type": "epistel", "value": sequence.strip()})
                case sequence if "JAMITA" in sequence.upper():
                    create_sermon_caption(prs, sequence.strip())
                    sequences.append({"type": "epistel", "value": sequence.strip()})
                case _:
                    create_caption(prs, sequence.strip())
                    sequences.append({"type": "common", "value": sequence.strip()})


def extract_text_from_pdf(prs, pdf_path, language):
    caption, title, topik = "", "", ""
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            # print(f"--- Halaman {i} ---")
            lines = page.extract_text().splitlines()
            for line in lines:
                if language == "1":
                    if "Minggunta Margoar" in line:
                        title = line.split(":")[1].strip()
                    if "Topik Minggunta" in line:
                        topik = line.split(":")[1].strip()
                else:
                    if "Minggu kita hari ini" in line:
                        title = line.split(":")[1].strip()
                    if "Topik Minggu" in line:
                        topik = line.split(":")[1].strip()
    caption = f"{title}\n\n{topik}"
    create_caption(prs, caption)


def generate_ppt(pdf_path, pptx_path, language):
    prs = Presentation()

    extract_text_from_pdf(prs, pdf_path, language)
    # read_file(pdf_path ,prs, language)

    prs.save(pptx_path)



# Inisialisasi colorama
init(autoreset=True)

def print_description():
    """Menampilkan deskripsi aplikasi dengan warna dan gaya."""
    print(Fore.CYAN + """
+-----------------------------------------------------------+
|           Tentang PPT Generator HKBP Depok II             |
+-----------------------------------------------------------+
""" + Fore.YELLOW + """
| Alat ini membantu membuat slide PowerPoint dari Warta     |
| Ibadah yang diberikan Sekretariat Gereja.                 |
|                                                           |
| Namun, hasilnya perlu dicek ulang, terutama pada:         |
| 1. Koor       - Periksa urutan Koor.                      |
| 2. Tingting   - Tingting belum ada pada PPT ini.          |
| 3. Ayat Khotbah                                           |
|                                                           |
| Gunakan alat ini untuk menghemat waktu, tapi pastikan     |
| pengecekan manual tetap dilakukan.                        |
""" + Fore.CYAN + """
+-----------------------------------------------------------+
""")


def main():
    """Fungsi utama untuk menjalankan aplikasi."""
    print_description()

    # Input file PDF
    pdf_path = input(Fore.GREEN + "Masukkan nama file PDF dan path warta dari Gereja: ").strip()
    if not pdf_path.lower().endswith('.pdf'):
        print(Fore.RED + "Error: File harus dalam format .pdf!")
        return

    # Cek apakah file ada
    if not os.path.exists(pdf_path):
        print(Fore.RED + f"Error: File '{pdf_path}' tidak ditemukan!")
        return

    # Pilihan bahasa
    print(Fore.BLUE + "\nPilih bahasa:")
    print(Fore.YELLOW + "1. BATAK")
    print(Fore.YELLOW + "2. INDONESIA")
    language = input(Fore.GREEN + "Masukkan pilihan (1 atau 2): ").strip()
    if language not in ['1', '2']:
        print(Fore.RED + "Error: Pilihan bahasa tidak valid! Pilih 1 atau 2.")
        return

    # Output file PPTX
    pptx_path = pdf_path.replace('.pdf', '.pptx')
    bahasa = "BATAK" if language == "1" else "INDONESIA"
    print(Fore.CYAN + "\nMulai membuat PPT dalam bahasa " + bahasa )

    try:
        generate_ppt(pdf_path, pptx_path,language)  # Asumsi fungsi ini ada
        print(Fore.GREEN + f"PPT selesai! Silahkan cek di: {pptx_path}")
    except Exception as e:
        print(Fore.RED + f"Terjadi kesalahan saat membuat PPT: {e}")

    input(Fore.YELLOW + "\nTekan Enter untuk keluar...")


# Jalankan aplikasi
if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print(Fore.RED + "\nAplikasi dihentikan oleh pengguna.")
