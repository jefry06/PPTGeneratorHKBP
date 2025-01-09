from converter.GeneratePPTFromWarta import generate_ppt

def print_description():
    print("""
+-----------------------------------------------------------+
|           Tentang PPT Generator HKBP Depok II             |
+-----------------------------------------------------------+
| Alat ini membantu membuat slide PowerPoint dari Warta     |
| Ibadah yang diberikan Sekretariat Gereja.                 |
|                                                           |
| Namun, hasilnya perlu dicek ulang, terutama pada:         |
| 1. Koor       - Check urutan Koor.                        |
| 2. Tingting   - Tingting belum ada pada PPT ini.          |
| 3. Ayat Khotbah                                           |
|                                                           |
| Gunakan alat ini untuk menghemat waktu, tapi pastikan     |
| pengecekan manual tetap dilakukan.                        |
+-----------------------------------------------------------+
    """)



def main():
    print_description()
    pdf_path = input("Masukkan nama file PDF dan path warta dari Gereja: ")
    if not pdf_path.lower().endswith('.pdf'):
        print("Error: File harus dalam format pdf!")
        return

    # Check if file exists
    import os
    if not os.path.exists(pdf_path):
        print(f"Error: File {pdf_path} tidak ditemukan!")
        return

    language = input("Pilih bahasa (1 untuk BATAK, 2 untuk INDONESIA): ")
    if language not in ['1', '2']:
        print("Error: Pilihan bahasa tidak valid!")
        return

    pptx_path = pdf_path.replace('.pdf', '.pptx')

    print("Mulai membuat PPT...")
    generate_ppt()
    print(f"PPT Selesai! Silahkan cek di {pptx_path}")
    input("\nTekan Enter untuk keluar...")

if __name__ == "__main__":
    main()