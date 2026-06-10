import requests
from bs4 import BeautifulSoup
import pandas as pd
import time

data_penulis = []
unej = 406
unmuh = 2060
poltek = 447

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36"
}

# --- LANGKAH 1: Ambil Halaman Pertama untuk Hitung Total Halaman & Nama Kampus ---
url_awal = f"https://sinta.kemdiktisaintek.go.id/affiliations/authors/{poltek}?page=1"
response = requests.get(url_awal, headers=headers)

if response.status_code == 200:
    soup = BeautifulSoup(response.content, 'html.parser')
    
    # Ambil nama universitas sekali saja di awal halaman
    univ_tag = soup.find('div', class_='univ-name')
    if univ_tag and univ_tag.find('h3') and univ_tag.find('h3').find('a'):
        universitas = univ_tag.find('h3').find('a').text.strip()
    else:
        universitas = f"Univ_{poltek}"
    
    # Ambil total halaman secara dinamis
    pages = soup.find_all('div', class_='text-center pagination-text')
    c = 1  # Default jika tidak ditemukan
    for page in pages:
        small_tag = page.find('small')
        if small_tag:
            pages_sum = small_tag.text.strip()  # <-- SUDAH DI-FIX
            bagian_kiri = pages_sum.split(" | ")[0] 
            c = int(bagian_kiri.split(" ")[-1])
            
    print(f"Universitas Terdeteksi: {universitas}")
    print(f"Total Halaman Ditemukan: {c} halaman\nS memulai proses scraping...\n")

    # --- LANGKAH 2: Looping Semua Halaman (menggunakan c + 1) ---
    for page in range(1, c + 1):
        url = f"https://sinta.kemdiktisaintek.go.id/affiliations/authors/{poltek}?page={page}"
        
        try:
            # PENTING: Lakukan request dan update soup BARU di dalam loop
            page_response = requests.get(url, headers=headers)
            
            if page_response.status_code == 200:
                page_soup = BeautifulSoup(page_response.content, 'html.parser')
                authors = page_soup.find_all('div', class_='au-item mt-3 mb-3 pb-5 pt-3')
                
                for author in authors:
                    name_div = author.find('div', class_='profile-name')
                    name = name_div.find('a').text.strip() if name_div and name_div.find('a') else "Tidak ditemukan"
        
                    dept_div = author.find('div', class_='profile-dept')
                    dept = dept_div.find('a').text.strip() if dept_div and dept_div.find('a') else "Tidak ditemukan"
       
                    data_penulis.append({
                        "Nama": name,
                        "Program Studi": dept
                    })
                
                print(f"Berhasil mengekstrak data dari halaman {page} dari {c}")
            else:
                print(f"Gagal memuat halaman {page}. Status code: {page_response.status_code}")
                
        except Exception as e:
            print(f"Terjadi error pada halaman {page}: {e}")

        # Jeda rapi agar tidak membebani server
        time.sleep(3)

    # --- LANGKAH 3: Penyimpanan Data ---
    df = pd.DataFrame(data_penulis)
    # Membersihkan nama file dari spasi agar lebih aman
    nama_file = f"data_penulis_{universitas.replace(' ', '_')}.csv"
    df.to_csv(nama_file, index=False, encoding='utf-8')

    print(f"\nSelesai! Total data yang diambil: {len(data_penulis)} penulis.")
    print(f"Data sukses disimpan dalam file '{nama_file}'")

else:
    print(f"Gagal melakukan koneksi awal ke SINTA. Status code: {response.status_code}")