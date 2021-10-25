'Nama fungsi: Pengeja Bilangan
' Parameter Input: angka - tipe Long
' output fungsi: tipe String
Function PengejaBilangan(angka As Long) As String
 
    'DEKLARASI VARIABEL-------------------
    'Variabel "Temp" untuk menampung hasil
    'Ini penamaan umum, temp berasal dari kata temporary
    'yang artinya sementara
    'dengan kata lain, variabel ini hanya variabel sementara saja
    'yang hanya digunakan sesaat
    Temp = ""
    
    ' "numbers" adalah array untuk menampung angka ejaan
    'Array adalah tipe data seperti kontainer
    'Misalkan kontainer apel, maka di dalamnya pasti apel semua
    'Dalam kasus ini, numbers berisi tipe data String
    'anggota / elemen pada array dapat di akses menggunakan syntax:
    '       nama_array(index)
    '
    'dimana index adalah angka yang telah didefinisikan sebelumnya
    'pada saat deklarasi array
    'Contohnya, dalam kasus ini, index nya adalah 0 sampai 9
    ' - Apabila kita mengakses elemen array diluar index, maka akan ada error
    '       ex:
    '           numbers(20)     ERROR
    '           numbers(0)      OK
    '           numbers(4)      OK
    '           numbers(9)      OK
    '
    'deklarasi array:
    Dim numbers(0 To 9) As String
    
    'Mengisi array "numbers" dengan elemen yang sesuai
    numbers(0) = ""
    numbers(1) = "SATU "
    numbers(2) = "DUA "
    numbers(3) = "TIGA "
    numbers(4) = "EMPAT "
    numbers(5) = "LIMA "
    numbers(6) = "ENAM "
    numbers(7) = "TUJUH "
    numbers(8) = "DELAPAN "
    numbers(9) = "SEMBILAN "
    
    'array angka_input untuk menampung angka yang dimasukkan ke fungsi ini
    'saat memanggil fungsi ini di excel.
    'elemen dari array angka_input memiliki tipe data Double (bilangan desimal)
    'Karena mata uang umumnya dipecah menjadi kelompok 3 bilangan,
    'Maka array dibuat menjadi memiliki 3 elemen, yaitu 1 s/d 3
    'Dengan bilangan diurutkan dari kiri ke kanan sbb:
    '   pertama-kedua-ketiga
    Dim angka_input(1 To 3) As Double
    
    
    'PROSEDUR MENDAPATKAN PANJANG ANGKA
    '   misal, input = 100
    '   maka, panjang angka = 3
    '
    '   misal#2, input = 14
    '   maka, panjang angka = 2
    'pertama, kita memaksa agar input yang awalnya bertipe Long, menjadi String
    angka_string = Str(angka)
    
    'Selanjutnya, kita hilangkan spasi di sekitar string yang sudah dibentuk
    'dengan menggunakan fungsi Trim()
    angka_trimmed = Trim(angka_string)
    
    'Terakhir, setelah di hilangkan spasi di sekitar angkanya, bisa didapatkan
    ' "panjang" dari angka input, dengan menggunakan fungsi Len()
    ' dari kata "length" (panjang)
    ' fungsi ini hanya bisa digunakan untuk String, dan juga spasi dihitung
    ' sebagai 1 karakter. Karena itulah 2 langkah diatas harus dilakukan
    ' agar tidak terjadi kesalahan / error
    panjang_angka = Len(angka_trimmed)
    
    
    'PROSEDUR EKSTRAKSI ANGKA
    nilai = Right("000", 3 - panjang) + Trim(Str(angka))
    For y = 3 To 1 Step -1
        angka_input(y) = Mid(nilai, y, 1)
    Next y
    
    
    PengejaBilangan = "Panjang angka = " & panjang_angka
    
End Function
