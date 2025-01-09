import json
import pywhatkit as kit
import time
import os

def create_invitations_and_send(data):
    # Template undangan
    template = u"""*Syalom {Salutation} {Nama},*

Selamat Tahun Baru 2025! Semoga {Salutation} dan keluarga selalu dalam keadaan sehat.  
Kami ingin memberitahukan bahwa hasil rapat ruas sektor 7 merekomendasikan {Salutation} sebagai *{seksi}* dalam kepanitiaan Pesta Gotilon 2025 HKBP Depok II.

Mohon konfirmasi apakah {Salutation} bersedia menerima rekomendasi tersebut.

Kami juga mengundang {Salutation} untuk hadir pada *Rapat Perdana Calon Panitia* yang akan diadakan:  
📅 *Rabu, 8 Januari 2025*  
🕒 *19.30 WIB*  
📍 *via Zoom (online) https://bit.ly/gotilon-8-januari*

Kehadiran {Salutation} sangat diharapkan untuk membahas rencana program Pesta Gotilon 2025.

Pada *Jumat, 10 Januari 2025, pukul 19.30 di Gereja HKBP Depok II*, kami mengundang kembali {Salutation} untuk acara Ramah Tamah Panitia dan finalisasi materi Pesta Gotilon yang akan dibawa ke *Rapat Huria*. Mohon konfirmasi apakah {Salutation} bisa hadir.

Terima kasih atas perhatian {Salutation}.

*Horas,*  
Jefry Pasaribu  
(Calon Panitia)"""

    # Folder untuk menyimpan hasil undangan
    output_folder = "f:/HKBP/Gotilon 2025/invitation/result/"
    os.makedirs(output_folder, exist_ok=True)

    # Loop untuk membuat file dan mengirim pesan
    for person in data:
        # Buat konten pesan berdasarkan template
        content = template.format(Salutation=person['salutation'], Nama=person['nama'], seksi=person['seksi'])

        # Bersihkan nomor HP dari "+62"
        hp_cleaned = person['hp'].replace('+62', '')

        # Simpan undangan dalam file txt
        filename = os.path.join(output_folder, f"{hp_cleaned}.txt")
        with open(filename, 'w', encoding='utf-8') as file:
            file.write(content)

        # Kirim pesan melalui pywhatkit
        try:
            print(f"Mengirim pesan ke {person['nama']} ({person['hp']})...")
            current_hour = int(time.strftime("%H"))
            current_minute = int(time.strftime("%M")) + 1  # Waktu pengiriman 1 menit dari sekarang

            kit.sendwhatmsg(person['hp'], content, current_hour, current_minute)
            time.sleep(30)  # Tunggu 1 menit untuk menghindari spam
        except Exception as e:
            print(f"Gagal mengirim pesan ke {person['hp']} ({person['nama']}): {e}")

# Membaca file JSON
with open('f:/HKBP/Gotilon 2025/invitation/list.json', 'r', encoding='utf-8-sig') as json_file:  # Handle BOM by using utf-8-sig
    data = json.load(json_file)

# Panggil fungsi untuk membuat undangan dan mengirim pesan
create_invitations_and_send(data)
