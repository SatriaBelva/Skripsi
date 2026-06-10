# tes = [1,2,3,4,5,6,7,8,123213123]
# a = len(tes)
# print(tes[a-1])

tes = [1,2,3,4,5,6,7,8,123213123]
a = 12
tes.append(a)
print(tes)
b = len(tes)
print(tes[b-1])

d = "Page 2 of 175 | Total Records 1.750"

# Langkah 1: Pisahkan teks berdasarkan karakter " | " (spasi, garis lurus, spasi)
# a.split(" | ") akan menghasilkan list: ['Page 2 of 175', 'Total Records 1.750']
# Kita ambil elemen pertama (indeks 0)
bagian_kiri = d.split(" | ")[0] 

# Langkah 2: Pisahkan lagi 'bagian_kiri' berdasarkan spasi (" ")
# bagian_kiri.split(" ") menghasilkan list: ['Page', '2', 'of', '175']
# Kita ambil elemen paling akhir (indeks -1)
c = bagian_kiri.split(" ")[-1]

# Langkah 3: (Opsional tapi Penting) Ubah teks menjadi angka integer
c = int(c)

print(c)
# Output: 175

unej = 406
unmuh = 2060
unej = 447