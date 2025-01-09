from encodings.utf_7 import encode

import requests
from bs4 import BeautifulSoup
import time

def get_pericope_from_url(path, start_verse, end_verse):
    url = 'https://alkitab.me/bbc-toba/' + path

    session = requests.Session()
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    response = session.get(url, headers=headers, timeout=15)

    if response.status_code == 200:
        soup = BeautifulSoup(response.content, 'html.parser')
        verses = []
        for verse_number in range(int(start_verse), int(end_verse) + 1):
            verse = soup.find('div', {'data-verse': str(verse_number)})
            if verse:
                verse_number_text = verse.find('span', class_='vn').text.strip()
                verse_text = verse.find('p', class_='leading0').text.strip()
                # verses.append((verse_number_text, verse_text))
                verses.append(f"({verse_number}) {verse_text}")
                # print(f"Nomor Ayat: {verse_number_text}")
                # print(f"Isi Ayat: {verse_text}")
                # print('-' * 50)
            else:
                print(f"Ayat {verse_number} tidak ditemukan.")

        return verses
    else:
        print(url)
        print(f"Failed to retrieve content. HTTP Status code: {response.status_code}")
        return []


def get_lyrics_from_url(song_type, song_number, start_verse, last_verse):
    # URL to fetch content from
    if song_type == "BUKU ENDE":
        url = "https://alkitab.app/BE/" + str(song_number)
    else:
        url = "https://alkitab.app/BN/" + str(song_number)

    try:
        # Send a GET request to the URL with timeout
        session = requests.Session()
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        response = session.get(url, headers=headers, timeout=15)

        # Check if the request was successful (status code 200)
        if response.status_code == 200:
            # Parse the content with BeautifulSoup
            soup = BeautifulSoup(response.content, 'html.parser')

            # Add a small delay to ensure content is loaded
            time.sleep(2)

            # Extract the title using the specified CSS selector
            title = soup.select_one('body > div.container > div > div.col-md-6.float > div > div.judul')
            title_text = title.text.strip() if title else 'Title not found'

            # Extract the lyrics using the specified CSS selector
            lyrics_container = soup.select_one('body > div.container > div > div.col-md-6.float > div > div.lirik')

            # Initialize a list to store the lyrics
            lyrics = []

            if lyrics_container:
                # Extract the reff (if it exists)
                reff_element = lyrics_container.find('div', class_='bait reff')
                reff = ""
                if reff_element:
                    reff_lines = [line.text.strip() for line in reff_element.select('.baris')]
                    reff = "\n".join(reff_lines)

                # Extract verses and append the reff to each verse
                verses = lyrics_container.find_all('div', class_='bait')
                for verse in verses:
                    verse_no_element = verse.select_one('.bait-no')
                    verse_number = int(verse_no_element.text.strip()) if verse_no_element else 0
                    lines = [line.text.strip() for line in verse.select('.bait-text .baris')]
                    main_lyric = "\n".join(lines)

                    # Append reff to each verse
                    if reff:
                        full_lyric = f"({verse_number}) {main_lyric}\n{reff}"
                    else:
                        full_lyric = f"({verse_number}) {main_lyric}"

                    lyrics.append((verse_number, full_lyric))

            # print(lyrics)
            return lyrics
        else:
            print(url)
            print(f"Failed to retrieve content. HTTP Status code: {response.status_code}")
            return []

    except requests.Timeout:
        print("Request timed out. Please try again later.")
        return []
    except requests.RequestException as e:
        print(f"An error occurred: {str(e)}")
        return []
