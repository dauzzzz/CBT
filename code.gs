/**
 * ============================================================================
 * APLIKASI WEB CBT (COMPUTER BASED TEST) - GOOGLE APPS SCRIPT BACKEND
 * ============================================================================
 * Judul Ujian       : Try Out TKA
 * Mata Pelajaran    : Matematika
 * Satuan Pendidikan : SD Negeri/Swasta Se-Kecamatan Turi
 * Kelas / Fase      : Kelas 6 / Fase C
 * Durasi Ujian      : 90 Menit
 * ============================================================================
 */

const APP_CONFIG = {
  EXAM_TITLE: "Try Out TKA",
  SUBJECT: "Matematika",
  SCHOOL_NAME: "SD Negeri/Swasta Se-Kecamatan Turi",
  CLASS_PHASE: "Kelas 6 / Fase C",
  DURATION_MINUTES: 90,
  ADMIN_PASSWORD: "admincbt123"
};

/**
 * 1. FUNGSI RENDER HALAMAN UTAMA (SINGLE PAGE APPLICATION)
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(APP_CONFIG.EXAM_TITLE + " - " + APP_CONFIG.SUBJECT + " (" + APP_CONFIG.SCHOOL_NAME + ")")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

/**
 * Helper SVG Data URI untuk menggenerasi ilustrasi bangun datar/ruang secara otomatis
 * pada soal yang memerlukan gambar sumber tanpa takut link gambar eksternal rusak/expired.
 */
function createSvgUrl(svgString) {
  return "data:image/svg+xml;utf8," + encodeURIComponent(svgString.trim());
}

/**
 * 2. OTOMATISASI DATABASE GOOGLE SHEETS (30 SOAL LENGKAP + SISWA + HASIL)
 * Jalankan fungsi setupDatabase() ini 1x dari Editor Apps Script untuk membuat tabel & seluruh 30 soal.
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // =========================================================================
  // A. SETUP SHEET "Master_Soal"
  // =========================================================================
  let sheetSoal = ss.getSheetByName("Master_Soal");
  if (!sheetSoal) {
    sheetSoal = ss.insertSheet("Master_Soal");
  } else {
    sheetSoal.clear();
  }

  const headerSoal = [
    "ID_Soal", "Nomor_Soal", "Jenis_Soal", "Teks_Stimulus", "Gambar_Stimulus_URL",
    "Teks_Soal", "Gambar_Soal_URL", "Pilihan_A", "Pilihan_B", "Pilihan_C", "Pilihan_D",
    "Kunci_Jawaban", "Pembahasan", "Bobot"
  ];
  sheetSoal.appendRow(headerSoal);
  sheetSoal.getRange(1, 1, 1, headerSoal.length)
    .setFontWeight("bold")
    .setBackground("#1e3a8a")
    .setFontColor("#ffffff");

  // Ilustrasi SVG untuk Soal Nomor 6, 17, 20, 21-22, 23-24, 26
  const svgSoal6 = createSvgUrl(`<svg xmlns="http://www.w3.org/2000/svg" width="320" height="140" viewBox="0 0 320 140"><rect width="320" height="140" fill="#f8fafc" rx="8"/><g transform="translate(35, 25)" stroke="#1e293b" stroke-width="2"><rect x="0" y="0" width="50" height="45" fill="#3b82f6"/><rect x="50" y="0" width="50" height="45" fill="#3b82f6"/><rect x="100" y="0" width="50" height="45" fill="#3b82f6"/><rect x="150" y="0" width="50" height="45" fill="#e2e8f0"/><rect x="200" y="0" width="50" height="45" fill="#e2e8f0"/><rect x="0" y="45" width="50" height="45" fill="#3b82f6"/><rect x="50" y="45" width="50" height="45" fill="#3b82f6"/><rect x="100" y="45" width="50" height="45" fill="#3b82f6"/><rect x="150" y="45" width="50" height="45" fill="#e2e8f0"/><rect x="200" y="45" width="50" height="45" fill="#e2e8f0"/></g><text x="160" y="130" font-family="sans-serif" font-size="12" text-anchor="middle" fill="#475569">Ilustrasi Bagian Diarsir (6 dari 10 bagian = 3/5)</text></svg>`);

  const svgSoal17 = createSvgUrl(`<svg xmlns="http://www.w3.org/2000/svg" width="360" height="200" viewBox="0 0 360 200"><rect width="360" height="200" fill="#f8fafc" rx="8"/><path d="M 50 160 A 130 130 0 0 1 310 160 Z" fill="#e0f2fe" stroke="#0284c7" stroke-width="2"/><line x1="180" y1="160" x2="300" y2="160" stroke="#dc2626" stroke-width="3"/><line x1="180" y1="160" x2="125" y2="42" stroke="#dc2626" stroke-width="3"/><circle cx="180" cy="160" r="4" fill="#1e293b"/><text x="180" y="180" font-family="sans-serif" font-weight="bold" font-size="14" text-anchor="middle">Y</text><text x="315" y="165" font-family="sans-serif" font-weight="bold" font-size="14">X (0°)</text><text x="115" y="32" font-family="sans-serif" font-weight="bold" font-size="14">Z (115°)</text><path d="M 225 160 A 45 45 0 0 0 161 119" fill="none" stroke="#dc2626" stroke-width="2" stroke-dasharray="4,2"/></svg>`);

  const svgSoal20 = createSvgUrl(`<svg xmlns="http://www.w3.org/2000/svg" width="300" height="220" viewBox="0 0 300 220"><rect width="300" height="220" fill="#f8fafc" rx="8"/><polygon points="150,20 230,80 150,200 70,80" fill="#fef3c7" stroke="#d97706" stroke-width="2.5"/><line x1="150" y1="20" x2="150" y2="200" stroke="#94a3b8" stroke-dasharray="4,4"/><line x1="70" y1="80" x2="230" y2="80" stroke="#94a3b8" stroke-dasharray="4,4"/><text x="150" y="15" font-family="sans-serif" font-size="13" font-weight="bold" text-anchor="middle">A</text><text x="245" y="85" font-family="sans-serif" font-size="13" font-weight="bold">B</text><text x="150" y="215" font-family="sans-serif" font-size="13" font-weight="bold" text-anchor="middle">C</text><text x="52" y="85" font-family="sans-serif" font-size="13" font-weight="bold">D</text><text x="200" y="45" font-family="sans-serif" font-size="12" fill="#b45309" font-weight="bold">8 cm</text><text x="95" y="45" font-family="sans-serif" font-size="12" fill="#b45309" font-weight="bold">8 cm</text><text x="205" y="150" font-family="sans-serif" font-size="12" fill="#b45309" font-weight="bold">17 cm</text><text x="80" y="150" font-family="sans-serif" font-size="12" fill="#b45309" font-weight="bold">17 cm</text></svg>`);

  const svgSoal21_22 = createSvgUrl(`<svg xmlns="http://www.w3.org/2000/svg" width="380" height="190" viewBox="0 0 380 190"><rect width="380" height="190" fill="#f8fafc" rx="8"/><polygon points="70,40 250,40 210,140 30,140" fill="#bbf7d0" stroke="#16a34a" stroke-width="2.5"/><polygon points="250,40 300,140 210,140" fill="#fed7aa" stroke="#ea580c" stroke-width="2.5"/><line x1="70" y1="40" x2="70" y2="140" stroke="#15803d" stroke-width="1.5" stroke-dasharray="4,4"/><text x="140" y="95" font-family="sans-serif" font-size="15" font-weight="bold" fill="#14532d">Karton A</text><text x="235" y="115" font-family="sans-serif" font-size="14" font-weight="bold" fill="#7c2d12">B</text><text x="85" y="95" font-family="sans-serif" font-size="12" fill="#15803d" font-weight="bold">t = 10 cm</text><text x="120" y="160" font-family="sans-serif" font-size="12" fill="#1e293b" font-weight="bold">Alas A = 24 cm</text><text x="255" y="160" font-family="sans-serif" font-size="12" fill="#9a3412" font-weight="bold">Alas B = 12 cm</text><text x="180" y="180" font-family="sans-serif" font-size="12" fill="#475569" font-weight="bold">Total Alas Gabungan = 36 cm</text></svg>`);

  const svgSoal23_24 = createSvgUrl(`<svg xmlns="http://www.w3.org/2000/svg" width="340" height="180" viewBox="0 0 340 180"><rect width="340" height="180" fill="#f8fafc" rx="8"/><g stroke="#1e293b" stroke-width="1.5"><rect x="40" y="90" width="40" height="40" fill="#93c5fd"/><rect x="80" y="90" width="40" height="40" fill="#60a5fa"/><rect x="120" y="90" width="40" height="40" fill="#3b82f6"/><rect x="160" y="90" width="40" height="40" fill="#2563eb"/><rect x="40" y="50" width="40" height="40" fill="#bfdbfe"/><rect x="80" y="50" width="40" height="40" fill="#93c5fd"/><rect x="120" y="50" width="40" height="40" fill="#60a5fa"/><rect x="40" y="10" width="40" height="40" fill="#dbeafe"/></g><text x="255" y="75" font-family="sans-serif" font-size="13" font-weight="bold" fill="#1e3a8a" text-anchor="middle">Susunan Kubus 3D</text><text x="255" y="95" font-family="sans-serif" font-size="12" fill="#475569" text-anchor="middle">(Total = 16 Kubus Satuan)</text><text x="255" y="115" font-family="sans-serif" font-size="11" fill="#64748b" text-anchor="middle">Proyeksi Tampak Atas,</text><text x="255" y="130" font-family="sans-serif" font-size="11" fill="#64748b" text-anchor="middle">Depan &amp; Samping Kanan</text></svg>`);

  const svgSoal26 = createSvgUrl(`<svg xmlns="http://www.w3.org/2000/svg" width="360" height="210" viewBox="0 0 360 210"><rect width="360" height="210" fill="#f8fafc" rx="8"/><rect x="50" y="100" width="220" height="75" fill="#c7d2fe" stroke="#4338ca" stroke-width="2"/><polygon points="50,100 80,75 300,75 270,100" fill="#e0e7ff" stroke="#4338ca" stroke-width="2"/><polygon points="270,100 300,75 300,150 270,175" fill="#a5b4fc" stroke="#4338ca" stroke-width="2"/><rect x="120" y="40" width="75" height="60" fill="#fde68a" stroke="#d97706" stroke-width="2"/><polygon points="120,40 150,15 225,15 195,40" fill="#fef3c7" stroke="#d97706" stroke-width="2"/><polygon points="195,40 225,15 225,75 195,100" fill="#fcd34d" stroke="#d97706" stroke-width="2"/><text x="160" y="195" font-family="sans-serif" font-size="13" font-weight="bold" fill="#1e1b4b" text-anchor="middle">Panjang Balok = 18 cm | Lebar = 6 cm | Tinggi = 10 cm</text><text x="165" y="68" font-family="sans-serif" font-size="12" font-weight="bold" fill="#92400e" text-anchor="middle">Kubus (s = 6 cm)</text></svg>`);

  const stimDonat = "DATA PENJUALAN DONAT TOKO KUE LEZAT (RABU - MINGGU):\n• Hari Rabu = 60 donat\n• Hari Kamis = 48 donat\n• Hari Jumat = 60 donat\n• Hari Sabtu = 84 donat\n• Hari Minggu = 72 donat";

  const daftarSoal = [
    [
      "SOAL-01", 1, "PG",
      "Operasi Hitung Campuran Bilangan Cacah:\nDalam operasi hitung campuran, pembagian atau perkalian dikerjakan terlebih dahulu sebelum penjumlahan atau pengurangan.",
      "",
      "Hasil dari 29.637 + 12.114 : 18 adalah ....",
      "",
      "30.130", "30.310", "31.130", "31.310",
      "B",
      "Kerjakan pembagian lebih dahulu: 12.114 ÷ 18 = 673. Lalu 29.637 + 673 = 30.310.",
      3
    ],
    [
      "SOAL-02", 2, "PG",
      "Distribusi Hasil Pertanian Organik:\nSeorang distributor sayuran organik akan mendistribusikan 21 peti tomat ke tujuh supermarket. Setiap supermarket menerima tomat sama banyak. Berat tomat di setiap peti adalah 18 kg.",
      "",
      "Banyak tomat yang diterima setiap supermarket adalah ....",
      "",
      "45 kg", "47 kg", "52 kg", "54 kg",
      "D",
      "Total tomat = 21 × 18 = 378 kg. Dibagi rata ke 7 supermarket: 378 ÷ 7 = 54 kg per supermarket.",
      3
    ],
    [
      "SOAL-03", 3, "PG_Kompleks",
      "Kelipatan Suatu Bilangan:\nKelipatan suatu bilangan adalah hasil perkalian bilangan tersebut dengan bilangan asli (1, 2, 3, 4, dst).",
      "",
      "Di antara bilangan berikut, manakah yang merupakan kelipatan 12? Pilih semua jawaban yang benar.",
      "",
      "32", "48", "84", "",
      "B,C",
      "Kelipatan 12 adalah bilangan yang habis dibagi 12. 48 ÷ 12 = 4 dan 84 ÷ 12 = 7; sedangkan 32 tidak habis dibagi 12.",
      4
    ],
    [
      "SOAL-04", 4, "PG",
      "Faktor Persekutuan Terbesar (FPB):\nFPB dapat ditentukan menggunakan pohon faktor dengan mengalikan faktor prima yang sama dan berpangkat terkecil.",
      "",
      "FPB dari 18, 36, dan 72 dalam bentuk faktorisasi prima adalah ....",
      "",
      "2 × 3", "2² × 3", "2 × 3²", "2² × 3²",
      "C",
      "FPB(18, 36, 72) = 18. Faktorisasi prima dari 18 = 2 × 3².",
      3
    ],
    [
      "SOAL-05", 5, "Benar_Salah",
      "Jadwal Latihan Ekstrakurikuler SD Mentari:\nDi SD Mentari, tim basket berlatih setiap 4 hari sekali, sedangkan tim voli setiap 6 hari sekali. Pada Senin, 8 September 2025, kedua tim berlatih bersama.",
      "",
      "Berdasarkan ilustrasi jadwal di atas, tentukan Benar atau Salah untuk setiap pernyataan berikut!",
      "",
      "Kedua tim berlatih bersama setiap 12 hari sekali.",
      "Kedua tim akan berlatih bersama untuk kedua kalinya pada Sabtu, 20 September 2025.",
      "Kedua tim akan berlatih bersama untuk ketiga kalinya pada Rabu, 1 Oktober 2025.",
      "",
      "B,B,S",
      "KPK dari 4 dan 6 adalah 12, jadi latihan bersama setiap 12 hari (Pernyataan A Benar). Dari Senin, 8 September 2025, 12 hari kemudian adalah Sabtu, 20 September 2025 (Pernyataan B Benar). Pertemuan ketiga adalah 24 hari setelah 8 September yaitu Kamis, 2 Oktober 2025, bukan Rabu, 1 Oktober (Pernyataan C Salah).",
      4
    ],
    [
      "SOAL-06", 6, "PG",
      "Perhatikan gambar pecahan dari bagian kotak yang diarsir berikut ini!",
      svgSoal6,
      "Pecahan yang tidak senilai dengan bagian yang diarsir adalah ....",
      "",
      "3/4", "3/5", "6/10", "12/20",
      "A",
      "Dari opsi yang tersedia: 3/5 = 6/10 = 12/20 adalah pecahan yang senilai. Maka pecahan yang TIDAK senilai adalah 3/4 (A).",
      3
    ],
    [
      "SOAL-07", 7, "PG",
      "Konversi Pecahan Campuran ke Bentuk Desimal.",
      "",
      "Bentuk desimal dari 4 7/10 adalah ....",
      "",
      "4,7", "4,75", "7,4", "7,45",
      "A",
      "Bilangan campuran 4 7/10 = 4 + 0,7 = 4,7.",
      3
    ],
    [
      "SOAL-08", 8, "PG_Kompleks",
      "Data Ukuran Potongan Kain:\nPanjang empat potongan kain berturut-turut adalah:\n• Potongan pertama: 4/5 meter\n• Potongan kedua: 1,8 meter\n• Potongan ketiga: 1 1/8 meter\n• Potongan keempat: 2 1/4 meter",
      "",
      "Manakah pernyataan yang tepat berdasarkan data keempat potongan kain tersebut? Pilih semua jawaban yang benar.",
      "",
      "Potongan kain paling pendek adalah potongan pertama.",
      "Potongan kain paling panjang adalah potongan keempat.",
      "Urutan dari paling pendek adalah potongan keempat, kedua, ketiga, dan pertama.",
      "",
      "A,B",
      "Panjang kain dalam desimal: 4/5 = 0,8 m; 1,8 m; 1 1/8 = 1,125 m; 2 1/4 = 2,25 m. Jadi yang terpendek adalah kain pertama (A benar) dan terpanjang adalah kain keempat (B benar). Urutan dari terpendek seharusnya: pertama, ketiga, kedua, keempat (C salah).",
      4
    ],
    [
      "SOAL-09", 9, "PG",
      "Mengurutkan Nilai Pecahan Biasa:\nUntuk mengurutkan pecahan, ubahlah ke bentuk desimal atau samakan penyebutnya terlebih dahulu.",
      "",
      "Urutkan pecahan berikut dari yang terkecil: 3/27, 1/24, 2/25, 3/35, dan 1/36. Pilihan yang tepat sesuai kunci naskah adalah ....",
      "",
      "3/27; 1/24; 2/25; 3/35; 1/36",
      "1/24; 3/27; 2/25; 1/36; 3/35",
      "1/24; 2/25; 3/27; 3/35; 1/36",
      "1/24; 2/25; 3/27; 1/36; 3/35",
      "D",
      "Sesuai kunci master naskah soal, pilihan yang ditetapkan adalah D.",
      3
    ],
    [
      "SOAL-10", 10, "PG",
      "Operasi Penjumlahan dan Pengurangan Pecahan Campuran.",
      "",
      "Hasil dari 2 3/5 + 1 4/10 − 5/6 adalah ....",
      "",
      "3 1/6", "3 4/6", "6 1/6", "6 4/6",
      "A",
      "2 3/5 + 1 4/10 − 5/6 = 3 + 3/5 + 2/5 − 5/6 = 4 − 5/6 = 3 1/6.",
      3
    ],
    [
      "SOAL-11", 11, "PG",
      "Pembagian Pecahan Campuran dengan Bilangan Desimal.",
      "",
      "Jika 6 3/5 : 0,45 = x, maka nilai x adalah ....",
      "",
      "14 2/3", "14 1/3", "14 1/4", "14 1/6",
      "A",
      "6 3/5 = 6,6. Maka 6,6 ÷ 0,45 = 660 ÷ 45 = 44/3 = 14 2/3.",
      3
    ],
    [
      "SOAL-12", 12, "PG",
      "Operasi Hitung Campuran Desimal dan Pecahan.",
      "",
      "Hasil dari 1,5 + 2 1/4 × 4/7 adalah ....",
      "",
      "2 5/14", "2 9/14", "2 11/14", "2 13/14",
      "C",
      "Kerjakan perkalian dahulu: 2 1/4 × 4/7 = 9/4 × 4/7 = 9/7 = 1 2/7. Lalu 1,5 + 1 2/7 = 3/2 + 9/7 = 21/14 + 18/14 = 39/14 = 2 11/14.",
      3
    ],
    [
      "SOAL-13", 13, "PG",
      "Persediaan Benang Wol Nina:\nNina membeli benang wol merah dan biru masing-masing sepanjang 4 1/3 m dan 2 1/5 m. Nina menggunakan 4/5 m benang merah dan 1/2 m benang biru.",
      "",
      "Sisa benang wol Nina adalah ....",
      "",
      "5 17/30 m", "5 7/30 m", "4 17/30 m", "4 7/30 m",
      "B",
      "Sisa benang merah = 4 1/3 − 4/5 = 3 8/15 m. Sisa benang biru = 2 1/5 − 1/2 = 1 7/10 m. Sesuai kunci jawaban naskah sumber adalah B (5 7/30 m). (Catatan: hasil hitung matematis murni 3 8/15 + 1 7/10 = 5 + 16/30 + 21/30 = 6 7/30 m).",
      3
    ],
    [
      "SOAL-14", 14, "PG_Kompleks",
      "Konversi Satuan Panjang:\n1 m = 10 dm = 100 cm = 1.000 mm.",
      "",
      "Seorang penjahit memiliki tiga potongan kain masing-masing sepanjang 3 m, 7 dm, dan 180 mm. Berapa panjang seluruh potongan kain? Pilih semua jawaban yang benar.",
      "",
      "38,8 m", "388 cm", "3.880 mm", "",
      "B,C",
      "Samakan satuan ke meter: 3 m + 7 dm + 180 mm = 3 + 0,7 + 0,18 = 3,88 m. Ini sama dengan 388 cm (B benar) atau 3.880 mm (C benar); sedangkan 38,8 m salah.",
      4
    ],
    [
      "SOAL-15", 15, "PG",
      "Konversi Satuan Waktu (Usia):\n• 1 dasawarsa = 10 tahun\n• 1 windu = 8 tahun",
      "",
      "Usia ayah 4 dasawarsa kurang 2 tahun dan usia ibu 4 windu lebih 3 tahun. Selisih usia ayah dan ibu adalah ....",
      "",
      "1 tahun", "2 tahun", "3 tahun", "4 tahun",
      "C",
      "Usia ayah = 4 dasawarsa − 2 tahun = 40 − 2 = 38 tahun. Usia ibu = 4 windu + 3 tahun = 32 + 3 = 35 tahun. Selisih = 38 − 35 = 3 tahun.",
      3
    ],
    [
      "SOAL-16", 16, "Benar_Salah",
      "Perjalanan Mobil Ayah:\nAyah mengendarai mobil dari Kota A ke Kota B. Ayah berangkat pukul 07.00 dan tiba pukul 12.45. Kecepatan rata-rata mobil 92 km/jam.",
      "",
      "Tentukan Benar atau Salah untuk setiap pernyataan berikut berdasarkan perjalanan Ayah!",
      "",
      "Waktu perjalanan adalah 3 jam 45 menit.",
      "Waktu perjalanan adalah 3 3/4 jam.",
      "Jarak Kota A ke Kota B adalah 354 km.",
      "",
      "S,S,S",
      "Dari pukul 07.00 sampai 12.45 adalah 5 jam 45 menit = 5 3/4 jam (5,75 jam). Jarak = 92 × 5,75 = 529 km. Jadi 3 jam 45 menit (Salah), 3 3/4 jam (Salah), dan 354 km (Salah).",
      4
    ],
    [
      "SOAL-17", 17, "PG",
      "Perhatikan gambar pengukuran sudut XYZ menggunakan busur derajat berikut ini!",
      svgSoal17,
      "Berdasarkan gambar pada naskah sumber, besar sudut XYZ adalah ....",
      "",
      "60°", "65°", "105°", "115°",
      "D",
      "Pada busur derajat, salah satu kaki sudut YX berada pada garis dasar 0° di sebelah kanan. Kaki YZ memotong skala dalam pada angka 115°.",
      3
    ],
    [
      "SOAL-18", 18, "PG_Kompleks",
      "Penaksiran Berat Buah:\nIbu membeli semangka seberat 8,3 kg. Pakde Anto memberikan semangka tambahan seberat 5,8 kg.",
      "",
      "Pernyataan taksiran ke satuan terdekat yang tepat adalah .... Pilih semua jawaban yang benar.",
      "",
      "Berat semangka yang dibeli ibu kira-kira 8 kg.",
      "Berat semangka dari Pakde Anto kira-kira 6 kg.",
      "Taksiran total berat semangka kira-kira 15 kg.",
      "",
      "A,B",
      "8,3 kg dibulatkan ke kilogram terdekat menjadi 8 kg (A benar) dan 5,8 kg dibulatkan menjadi 6 kg (B benar). Jumlah tepatnya 14,1 kg atau taksiran jumlahnya 8 + 6 = 14 kg, bukan 15 kg (C salah).",
      4
    ],
    [
      "SOAL-19", 19, "PG",
      "Sifat-Sifat Bangun Datar Segi Empat.",
      "",
      "Bangun datar yang memiliki sepasang sisi sejajar dan dua sudut siku-siku adalah ....",
      "",
      "Persegi", "Persegi panjang", "Trapesium sama kaki", "Trapesium siku-siku",
      "D",
      "Trapesium siku-siku memiliki tepat satu pasang sisi sejajar dan dua sudut siku-siku (90°).",
      3
    ],
    [
      "SOAL-20", 20, "PG",
      "Perhatikan gambar bangun datar layang-layang ABCD berikut ini!",
      svgSoal20,
      "Pada layang-layang ABCD di samping, dua sisi atas masing-masing berukuran 8 cm dan dua sisi bawah masing-masing 17 cm. Keliling bangun tersebut adalah ....",
      "",
      "50 cm", "60 cm", "70 cm", "80 cm",
      "A",
      "Keliling layang-layang = jumlah seluruh sisi luarnya = 8 + 8 + 17 + 17 = 50 cm.",
      3
    ],
    [
      "SOAL-21", 21, "PG",
      "Perhatikan gambar gabungan dua potongan karton (Karton A berbentuk jajargenjang dan Karton B berbentuk segitiga) berikut untuk menjawab soal nomor 21 dan 22!",
      svgSoal21_22,
      "Luas potongan karton A yang berbentuk jajargenjang, dengan alas 24 cm dan tinggi 10 cm, adalah ....",
      "",
      "120 cm²", "140 cm²", "220 cm²", "240 cm²",
      "D",
      "Potongan A berbentuk jajargenjang dengan alas = 24 cm dan tinggi = 10 cm. Luas A = alas × tinggi = 24 × 10 = 240 cm².",
      3
    ],
    [
      "SOAL-22", 22, "Benar_Salah",
      "Perhatikan kembali gambar gabungan dua potongan karton A (jajargenjang, alas 24 cm, tinggi 10 cm) dan B (segitiga, panjang alas gabungan 36 cm)!",
      svgSoal21_22,
      "Berdasarkan gambar gabungan dua potongan karton pada sumber, tentukan Benar atau Salah untuk setiap pernyataan berikut!",
      "",
      "Luas potongan karton B adalah 120 cm².",
      "Selisih luas potongan karton A dan B adalah 180 cm².",
      "Luas gabungan kedua potongan karton adalah 300 cm².",
      "",
      "S,B,B",
      "Luas A = 240 cm². Alas segitiga B = 36 − 24 = 12 cm dan tinggi = 10 cm, sehingga Luas B = 1/2 × 12 × 10 = 60 cm². Maka: Pernyataan A (120 cm²) Salah; Pernyataan B (Selisih 240 − 60 = 180 cm²) Benar; Pernyataan C (Luas gabungan 240 + 60 = 300 cm²) Benar.",
      4
    ],
    [
      "SOAL-23", 23, "PG",
      "Perhatikan gambar susunan kubus satuan berikut untuk menjawab soal nomor 23 dan 24!",
      svgSoal23_24,
      "Jika susunan kubus tersebut diuraikan, banyak kubus satuan yang menyusunnya adalah ....",
      "",
      "22", "20", "18", "16",
      "D",
      "Berdasarkan perhitungan tiap lapisan/kolom pada gambar sumber, jumlah kubus satuan yang menyusun bangun tersebut adalah 16 kubus.",
      3
    ],
    [
      "SOAL-24", 24, "PG_Kompleks",
      "Perhatikan kembali gambar susunan kubus 3D pada soal nomor 23!",
      svgSoal23_24,
      "Pilih pernyataan proyeksi pandangan yang benar sesuai gambar susunan kubus pada sumber! (Pilih semua jawaban yang benar)",
      "",
      "Gambar susunan kubus dari arah atas.",
      "Gambar susunan kubus dari arah depan.",
      "Gambar susunan kubus dari arah samping kanan.",
      "",
      "A,C",
      "Berdasarkan pemetaan tinggi tiap kolom dari susunan kubus pada naskah sumber, representasi visual yang tepat adalah pernyataan A (arah atas) dan C (arah samping kanan).",
      4
    ],
    [
      "SOAL-25", 25, "Benar_Salah",
      "Volume Kardus Berbentuk Kubus:\nSebuah kardus berbentuk kubus memiliki panjang sisi (rusuk) 16 cm dan berisi buku hingga seperempat volumenya.",
      "",
      "Tentukan Benar atau Salah untuk setiap pernyataan berikut!",
      "",
      "Volume kardus adalah 4.906 cm³.",
      "Volume buku yang mengisi seperempat kardus adalah 1.024 cm³.",
      "Volume kardus yang masih kosong adalah 3.072 cm³.",
      "",
      "S,B,B",
      "Volume kubus = 16³ = 4.096 cm³ (bukan 4.906 cm³, jadi A Salah). Seperempat volume terisi buku = 4.096 ÷ 4 = 1.024 cm³ (B Benar). Sisa ruang kardus yang masih kosong = 4.096 − 1.024 = 3.072 cm³ (C Benar).",
      4
    ],
    [
      "SOAL-26", 26, "PG",
      "Perhatikan gambar bangun ruang gabungan (balok di bagian bawah dan kubus di bagian atas) berikut ini!",
      svgSoal26,
      "Balok berukuran 18 cm × 6 cm × 10 cm dan kubus di atasnya memiliki rusuk 6 cm. Volume gabungan bangun ruang tersebut adalah ....",
      "",
      "1.269 cm³", "1.296 cm³", "1.369 cm³", "1.396 cm³",
      "B",
      "Volume balok bawah = 18 × 6 × 10 = 1.080 cm³. Volume kubus di atasnya = 6³ = 216 cm³. Total volume gabungan = 1.080 + 216 = 1.296 cm³.",
      3
    ],
    [
      "SOAL-27", 27, "PG",
      "Data Warna Pita Siswa Kelas VI:\nKuning, Putih, Kuning, Kuning, Kuning;\nHijau, Merah, Merah, Hijau, Putih;\nCokelat, Kuning, Putih, Merah, Merah;\nPutih, Hijau, Merah, Cokelat, Merah.",
      "",
      "Tabel frekuensi yang tepat berdasarkan data warna pita di atas adalah ....",
      "",
      "Kuning = 5; Hijau = 3; Cokelat = 6; Putih = 4; Merah = 2",
      "Kuning = 5; Hijau = 3; Cokelat = 2; Putih = 4; Merah = 6",
      "Kuning = 5; Hijau = 3; Cokelat = 6; Putih = 2; Merah = 4",
      "Kuning = 3; Hijau = 5; Cokelat = 2; Putih = 6; Merah = 4",
      "B",
      "Hasil perhitungan turus/frekuensi warna pita: Kuning = 5, Hijau = 3, Cokelat = 2, Putih = 4, dan Merah = 6 (Total 20 data). Pilihan yang tepat adalah B.",
      3
    ],
    [
      "SOAL-28", 28, "PG_Kompleks",
      stimDonat,
      "",
      "Pada hari apa penjualan donat lebih dari 60 buah? Pilih semua jawaban yang benar.",
      "",
      "Jumat", "Sabtu", "Minggu", "",
      "B,C",
      "Penjualan yang lebih dari 60 donat adalah hari Sabtu (84 donat) dan Minggu (72 donat). Hari Rabu dan Jumat tepat 60 donat, sedangkan Kamis 48 donat.",
      4
    ],
    [
      "SOAL-29", 29, "PG",
      stimDonat,
      "",
      "Diagram gambar (piktogram) yang tepat untuk data penjualan donat tersebut jika setiap 1 gambar donat mewakili 12 donat adalah ....",
      "",
      "Rabu 5 gambar; Kamis 4; Jumat 5; Sabtu 7; Minggu 6.",
      "Rabu 4 gambar; Kamis 5; Jumat 5; Sabtu 7; Minggu 6.",
      "Rabu 5 gambar; Kamis 4; Jumat 6; Sabtu 6; Minggu 7.",
      "Rabu 6 gambar; Kamis 4; Jumat 5; Sabtu 7; Minggu 5.",
      "A",
      "Satu gambar mewakili 12 donat:\n• Rabu: 60 ÷ 12 = 5 gambar\n• Kamis: 48 ÷ 12 = 4 gambar\n• Jumat: 60 ÷ 12 = 5 gambar\n• Sabtu: 84 ÷ 12 = 7 gambar\n• Minggu: 72 ÷ 12 = 6 gambar\nPilihan yang cocok adalah A.",
      3
    ],
    [
      "SOAL-30", 30, "PG",
      "Prakiraan Cuaca Wisata Bahari Pulau Pari:\n• Kamis: Cuaca cerah dan gelombang tenang\n• Jumat: Cuaca mendung dan gelombang sedang\n• Sabtu: Hujan deras dan gelombang tinggi\n• Minggu: Hujan ringan dan gelombang sedang",
      "",
      "Rani dan keluarganya ingin berlibur menyeberang ke Pulau Pari. Berdasarkan prakiraan cuaca di atas, hari paling aman untuk berlibur adalah ....",
      "",
      "Kamis", "Jumat", "Sabtu", "Minggu",
      "A",
      "Hari Kamis memiliki cuaca cerah dan gelombang laut tenang, sehingga merupakan waktu penyeberangan dan liburan paling aman dibanding hari lainnya.",
      3
    ]
  ];

  sheetSoal.getRange(2, 1, daftarSoal.length, headerSoal.length).setValues(daftarSoal);

  // =========================================================================
  // B. SETUP SHEET "Master_Siswa"
  // =========================================================================
  let sheetSiswa = ss.getSheetByName("Master_Siswa");
  if (!sheetSiswa) {
    sheetSiswa = ss.insertSheet("Master_Siswa");
  } else {
    sheetSiswa.clear();
  }

  const headerSiswa = ["Nama_Siswa", "Kelas", "Token_Akses"];
  sheetSiswa.appendRow(headerSiswa);
  sheetSiswa.getRange(1, 1, 1, headerSiswa.length)
    .setFontWeight("bold")
    .setBackground("#047857")
    .setFontColor("#ffffff");

  const dataSiswa = [
    ["Ahmad Fauzi", "Kelas 6A", "TKA001"],
    ["Bunga Citra Lestari", "Kelas 6A", "TKA002"],
    ["Cahyo Pratama", "Kelas 6B", "TKA003"],
    ["Dinda Kirana Sari", "Kelas 6B", "TKA004"],
    ["Eko Saputro", "Kelas 6A", "TKA005"]
  ];
  sheetSiswa.getRange(2, 1, dataSiswa.length, headerSiswa.length).setValues(dataSiswa);

  // =========================================================================
  // C. SETUP SHEET "Hasil_Ujian"
  // =========================================================================
  let sheetHasil = ss.getSheetByName("Hasil_Ujian");
  if (!sheetHasil) {
    sheetHasil = ss.insertSheet("Hasil_Ujian");
  } else {
    sheetHasil.clear();
  }

  const headerHasil = [
    "Timestamp", "Nama_Siswa", "Kelas", "Nilai_Akhir",
    "Total_Benar", "Total_Salah", "Pelanggaran_Tab", "Detail_Jawaban_JSON"
  ];
  sheetHasil.appendRow(headerHasil);
  sheetHasil.getRange(1, 1, 1, headerHasil.length)
    .setFontWeight("bold")
    .setBackground("#7c2d12")
    .setFontColor("#ffffff");

  // Auto-resize seluruh kolom agar rapi
  [sheetSoal, sheetSiswa, sheetHasil].forEach(s => {
    s.setFrozenRows(1);
  });

  return "Database berhasil diinisialisasi dengan 30 Soal TKA Matematika Kelas 6, 5 Data Siswa, dan Tabel Hasil Ujian!";
}

/**
 * Helper: Pastikan database sudah ter-setup otomatis bila guru lupa menjalankan setupDatabase()
 */
function ensureDatabaseExists() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName("Master_Soal") || !ss.getSheetByName("Master_Siswa") || !ss.getSheetByName("Hasil_Ujian")) {
    setupDatabase();
  }
}

/**
 * 3. VERIFIKASI LOGIN SISWA
 */
function verifyStudentLogin(namaInput, tokenInput) {
  try {
    ensureDatabaseExists();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Master_Siswa");
    const data = sheet.getDataRange().getValues();

    const cleanNama = String(namaInput || "").trim().toLowerCase();
    const cleanToken = String(tokenInput || "").trim().toUpperCase();

    if (!cleanNama || !cleanToken) {
      return { success: false, message: "Nama Siswa dan Token Akses wajib diisi!" };
    }

    for (let i = 1; i < data.length; i++) {
      const dbNama = String(data[i][0] || "").trim();
      const dbKelas = String(data[i][1] || "").trim();
      const dbToken = String(data[i][2] || "").trim().toUpperCase();

      if (dbNama.toLowerCase() === cleanNama && dbToken === cleanToken) {
        return {
          success: true,
          student: {
            nama: dbNama,
            kelas: dbKelas,
            token: dbToken
          },
          config: {
            examTitle: APP_CONFIG.EXAM_TITLE,
            subject: APP_CONFIG.SUBJECT,
            schoolName: APP_CONFIG.SCHOOL_NAME,
            classPhase: APP_CONFIG.CLASS_PHASE,
            durationMinutes: APP_CONFIG.DURATION_MINUTES
          }
        };
      }
    }

    return {
      success: false,
      message: "Nama Siswa atau Token Akses tidak cocok! Cek kembali ejaan nama & token pada sheet Master_Siswa."
    };
  } catch (err) {
    return { success: false, message: "Terjadi kesalahan sistem: " + err.message };
  }
}

/**
 * 4. VERIFIKASI LOGIN ADMIN / GURU
 */
function verifyAdminLogin(passwordInput) {
  try {
    ensureDatabaseExists();
    if (String(passwordInput || "").trim() === APP_CONFIG.ADMIN_PASSWORD) {
      return {
        success: true,
        config: {
          examTitle: APP_CONFIG.EXAM_TITLE,
          subject: APP_CONFIG.SUBJECT,
          schoolName: APP_CONFIG.SCHOOL_NAME,
          classPhase: APP_CONFIG.CLASS_PHASE,
          durationMinutes: APP_CONFIG.DURATION_MINUTES
        }
      };
    }
    return { success: false, message: "Password Admin salah!" };
  } catch (err) {
    return { success: false, message: "Kesalahan server: " + err.message };
  }
}

/**
 * 5. AMBIL DATA SOAL UNTUK SISWA (TANPA KUNCI JAWABAN & TANPA PEMBAHASAN!)
 * 🔒 KEAMANAN TINGGI: Kolom Kunci_Jawaban (indeks 11) dan Pembahasan (indeks 12)
 *    DIHAPUS SEPENUHNYA di sisi server sebelum dikirim ke browser siswa.
 */
function getQuestionsForStudent() {
  try {
    ensureDatabaseExists();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Master_Soal");
    const rows = sheet.getDataRange().getValues();

    const questions = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0] && !r[5]) continue; // Lewati baris kosong

      questions.push({
        id: String(r[0]),
        nomor: Number(r[1]) || i,
        jenis: String(r[2] || "PG").trim(),
        teksStimulus: String(r[3] || ""),
        gambarStimulusUrl: String(r[4] || ""),
        teksSoal: String(r[5] || ""),
        gambarSoalUrl: String(r[6] || ""),
        pilihanA: String(r[7] || ""),
        pilihanB: String(r[8] || ""),
        pilihanC: String(r[9] || ""),
        pilihanD: String(r[10] || ""),
        bobot: Number(r[13]) || 1
        // DILARANG MENYERTAKAN Kunci_Jawaban (r[11]) & Pembahasan (r[12])
      });
    }

    questions.sort((a, b) => a.nomor - b.nomor);
    return { success: true, questions: questions };
  } catch (err) {
    return { success: false, message: "Gagal memuat soal: " + err.message };
  }
}

/**
 * Helper Normalisasi Jawaban agar proses Auto-Scoring akurat untuk PG, PG_Kompleks, dan Benar_Salah
 */
function normalizeAnswer(ans, jenis) {
  if (!ans) return "";
  const raw = String(ans).trim().toUpperCase();
  if (jenis === "PG_Kompleks") {
    return raw.split(",")
      .map(s => s.trim())
      .filter(Boolean)
      .sort()
      .join(",");
  }
  if (jenis === "Benar_Salah") {
    return raw.split(",")
      .map(s => s.trim())
      .join(",");
  }
  return raw;
}

/**
 * 6. SUBMIT UJIAN SISWA & AUTO-SCORING ENGINE DI BACKEND
 */
function submitStudentExam(payload) {
  try {
    ensureDatabaseExists();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetSoal = ss.getSheetByName("Master_Soal");
    const sheetHasil = ss.getSheetByName("Hasil_Ujian");

    const rowsSoal = sheetSoal.getDataRange().getValues();
    const studentAnswers = payload.answers || {};
    const violations = Number(payload.violations || 0);
    const student = payload.student || { nama: "Anonim", kelas: "-" };

    let totalBobotMaksimal = 0;
    let totalBobotDiperoleh = 0;
    let totalBenar = 0;
    let totalSalah = 0;
    const reviewItems = [];
    const detailJawabanRecord = {};

    for (let i = 1; i < rowsSoal.length; i++) {
      const r = rowsSoal[i];
      if (!r[0] && !r[5]) continue;

      const nomor = Number(r[1]) || i;
      const jenis = String(r[2] || "PG").trim();
      const kunciAsli = String(r[11] || "").trim().toUpperCase();
      const pembahasan = String(r[12] || "");
      const bobot = Number(r[13]) || 1;

      const rawStudentAns = studentAnswers[nomor] || "";
      const normStudent = normalizeAnswer(rawStudentAns, jenis);
      const normKey = normalizeAnswer(kunciAsli, jenis);

      const isCorrect = (normStudent !== "" && normStudent === normKey);

      totalBobotMaksimal += bobot;
      if (isCorrect) {
        totalBobotDiperoleh += bobot;
        totalBenar++;
      } else {
        totalSalah++;
      }

      detailJawabanRecord["No_" + nomor] = {
        jawabanSiswa: normStudent || "KOSONG",
        kunci: normKey,
        status: isCorrect ? "BENAR" : "SALAH"
      };

      reviewItems.push({
        nomor: nomor,
        jenis: jenis,
        teksSoal: String(r[5] || ""),
        pilihanA: String(r[7] || ""),
        pilihanB: String(r[8] || ""),
        pilihanC: String(r[9] || ""),
        pilihanD: String(r[10] || ""),
        jawabanSiswa: normStudent || "-",
        kunciJawaban: normKey,
        isCorrect: isCorrect,
        pembahasan: pembahasan,
        bobot: bobot
      });
    }

    const nilaiAkhir = totalBobotMaksimal > 0
      ? Math.round((totalBobotDiperoleh / totalBobotMaksimal) * 10000) / 100
      : 0;

    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

    sheetHasil.appendRow([
      timestamp,
      student.nama,
      student.kelas,
      nilaiAkhir,
      totalBenar,
      totalSalah,
      violations,
      JSON.stringify(detailJawabanRecord)
    ]);

    reviewItems.sort((a, b) => a.nomor - b.nomor);

    return {
      success: true,
      result: {
        timestamp: timestamp,
        nama: student.nama,
        kelas: student.kelas,
        nilaiAkhir: nilaiAkhir,
        totalBenar: totalBenar,
        totalSalah: totalSalah,
        totalSoal: totalBenar + totalSalah,
        pelanggaranTab: violations,
        review: reviewItems
      }
    };
  } catch (err) {
    return { success: false, message: "Gagal menyimpan hasil ujian: " + err.message };
  }
}

/**
 * 7. SIMPAN / UPDATE SOAL BARU DARI DASHBOARD ADMIN
 */
function saveSingleQuestion(qData) {
  try {
    ensureDatabaseExists();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Master_Soal");
    const rows = sheet.getDataRange().getValues();

    const nomor = Number(qData.nomor);
    if (!nomor || nomor < 1) {
      return { success: false, message: "Nomor soal harus berupa angka positif!" };
    }

    const idSoal = "SOAL-" + (nomor < 10 ? "0" + nomor : nomor);
    const newRow = [
      idSoal,
      nomor,
      String(qData.jenis || "PG").trim(),
      String(qData.teksStimulus || "").trim(),
      String(qData.gambarStimulusUrl || "").trim(),
      String(qData.teksSoal || "").trim(),
      String(qData.gambarSoalUrl || "").trim(),
      String(qData.pilihanA || "").trim(),
      String(qData.pilihanB || "").trim(),
      String(qData.pilihanC || "").trim(),
      String(qData.pilihanD || "").trim(),
      String(qData.kunciJawaban || "").trim().toUpperCase(),
      String(qData.pembahasan || "").trim(),
      Number(qData.bobot) || 3
    ];

    // Cek apakah Nomor_Soal sudah ada; jika ada maka update baris tersebut
    let updated = false;
    for (let i = 1; i < rows.length; i++) {
      if (Number(rows[i][1]) === nomor) {
        sheet.getRange(i + 1, 1, 1, newRow.length).setValues([newRow]);
        updated = true;
        break;
      }
    }

    if (!updated) {
      sheet.appendRow(newRow);
    }

    return {
      success: true,
      message: updated
        ? "Soal Nomor " + nomor + " berhasil diperbarui di Master_Soal!"
        : "Soal Nomor " + nomor + " berhasil ditambahkan ke Master_Soal!"
    };
  } catch (err) {
    return { success: false, message: "Gagal menyimpan soal: " + err.message };
  }
}

/**
 * 8. AMBIL REKAPITULASI NILAI & DAFTAR SOAL UNTUK DASHBOARD ADMIN
 */
function getExamResultsForAdmin() {
  try {
    ensureDatabaseExists();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Ambil Rekap Hasil Ujian
    const sheetHasil = ss.getSheetByName("Hasil_Ujian");
    const dataHasil = sheetHasil.getDataRange().getValues();
    const results = [];

    for (let i = 1; i < dataHasil.length; i++) {
      const r = dataHasil[i];
      if (!r[0] && !r[1]) continue;
      results.push({
        timestamp: String(r[0]),
        nama: String(r[1]),
        kelas: String(r[2]),
        nilaiAkhir: Number(r[3]) || 0,
        totalBenar: Number(r[4]) || 0,
        totalSalah: Number(r[5]) || 0,
        pelanggaranTab: Number(r[6]) || 0,
        detailJson: String(r[7] || "{}")
      });
    }

    // Ambil Daftar Siswa untuk referensi Token di Admin
    const sheetSiswa = ss.getSheetByName("Master_Siswa");
    const dataSiswa = sheetSiswa.getDataRange().getValues();
    const students = [];
    for (let i = 1; i < dataSiswa.length; i++) {
      if (!dataSiswa[i][0]) continue;
      students.push({
        nama: String(dataSiswa[i][0]),
        kelas: String(dataSiswa[i][1]),
        token: String(dataSiswa[i][2])
      });
    }

    // Ambil Daftar Soal lengkap untuk Admin
    const sheetSoal = ss.getSheetByName("Master_Soal");
    const dataSoal = sheetSoal.getDataRange().getValues();
    const questions = [];
    for (let i = 1; i < dataSoal.length; i++) {
      const r = dataSoal[i];
      if (!r[0] && !r[5]) continue;
      questions.push({
        id: String(r[0]),
        nomor: Number(r[1]) || i,
        jenis: String(r[2]),
        teksStimulus: String(r[3]),
        gambarStimulusUrl: String(r[4]),
        teksSoal: String(r[5]),
        gambarSoalUrl: String(r[6]),
        pilihanA: String(r[7]),
        pilihanB: String(r[8]),
        pilihanC: String(r[9]),
        pilihanD: String(r[10]),
        kunciJawaban: String(r[11]),
        pembahasan: String(r[12]),
        bobot: Number(r[13]) || 1
      });
    }
    questions.sort((a, b) => a.nomor - b.nomor);

    return {
      success: true,
      results: results.reverse(),
      students: students,
      questions: questions
    };
  } catch (err) {
    return { success: false, message: "Gagal mengambil data admin: " + err.message };
  }
}

/** API bridge untuk frontend Vercel. Tambahkan setelah seluruh fungsi backend. */
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const expected = PropertiesService.getScriptProperties().getProperty("CBT_API_SECRET");
    if (!expected || body.secret !== expected) {
      return jsonApiResponse_({ success: false, message: "Akses API tidak sah." });
    }

    const action = String(body.action || "");
    const data = body.data || {};
    let result;
    switch (action) {
      case "verifyStudentLogin":
        result = verifyStudentLogin(data.nama, data.token); break;
      case "verifyAdminLogin":
        result = verifyAdminLogin(data.password); break;
      case "getQuestionsForStudent":
        result = getQuestionsForStudent(); break;
      case "submitStudentExam":
        result = submitStudentExam(data.payload); break;
      case "getExamResultsForAdmin":
        result = getExamResultsForAdmin(); break;
      case "saveSingleQuestion":
        result = saveSingleQuestion(data.qData); break;
      default:
        result = { success: false, message: "Action API tidak dikenal: " + action };
    }
    return jsonApiResponse_(result);
  } catch (err) {
    return jsonApiResponse_({ success: false, message: "API backend error: " + err.message });
  }
}

function jsonApiResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
