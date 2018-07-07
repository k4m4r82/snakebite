Attribute VB_Name = "ModLang"
'definisi bahasa
'BAHASA INDONESIA
'menu
Public Const id_LangmnuServer As String = "Server"
Public Const id_LangmnuMatikan As String = "Matikan"
Public Const id_LangGantiOp As String = "Ganti Operator"
Public Const id_LangPrintStruk As String = "Print Struk"
Public Const id_LangmnuClient As String = "Client"
Public Const id_LangmnuStop As String = "Stop Client"
Public Const id_LangmnuPindah As String = "Pindah Client"
Public Const id_LangmnuShutdown As String = "Shutdown Client"
Public Const id_LangmnuSetPasswordClient As String = "Set Password"
Public Const id_LangmnuSetting As String = "Setting"
Public Const id_LangmnuKonfig As String = "Konfigurasi"
Public Const id_LangmnuUser As String = "Manajemen User"
Public Const id_LangmnuLaporan As String = "Laporan"
Public Const id_LangmnuRekap As String = "Rekap Billing"
Public Const id_LangmnuDeposit As String = "Informasi Deposit"
Public Const id_LangmnuUtil As String = "Utility"
Public Const id_LangmnuMember As String = "Member"
Public Const id_LangmnuPesan As String = "Pesan"
Public Const id_LangmnuBackup As String = "Backup Database"
Public Const id_LangmnuHelp As String = "Help"
Public Const id_LangmnuAbout As String = "About"
Public Const id_Langmnuhapusclient As String = "Hapus Client"
'toolbar
Public Const id_LangTbarMatikan As String = "Matikan Server"
Public Const id_LangTbarStop As String = "Stop Client"
Public Const id_LangTbarPindah As String = "Pindah Client"
Public Const id_LangTbarShutdown As String = "Shutdown Client"
Public Const id_LangTbarSetPasswordClient As String = "Set Password Admin"
Public Const id_LangTbarKonfig As String = "Konfigurasi"
Public Const id_LangTbarUser As String = "Manajemen User"
Public Const id_LangTbarRekap As String = "Rekap Billing"
Public Const id_LangTbarDeposit As String = "Laporan Deposit"
Public Const id_LangTbarMember As String = "Manajemen Member"
Public Const id_LangTbarPesan As String = "Pesan"
Public Const id_LangTbarBackup As String = "Backup Database"
Public Const id_LangTbarAbout As String = "About"
'jenis
Public Const id_LangPersonal As String = "Personal"
Public Const id_LangMember As String = "Member"
Public Const id_LangGame As String = "Game"
Public Const id_LangKetik As String = "Ketik"
'status
Public Const id_LangBebas As String = "Bebas"
Public Const id_LangAktif As String = "Aktif"
Public Const id_LangPindah As String = "Pindah"
'billing
Public Const id_LangClient As String = "Client "
Public Const id_LangNamaUser As String = "Nama User"
Public Const id_LangStatus As String = "Status"
Public Const id_LangMulai As String = "Mulai"
Public Const id_LangDurasi As String = "Durasi"
Public Const id_LangBiaya As String = "Biaya"
Public Const id_LangJenis As String = "Jenis"
Public Const id_LangDiscount As String = "Discount"
Public Const id_LangTotal As String = "Total"
Public Const id_LangNota As String = "Nota"
Public Const id_LangNo As String = "No."
Public Const id_LangTanggal As String = "Tanggal"
Public Const id_LangSelesai As String = "Selesai"
Public Const id_LangNamaMember As String = "Nama Member"
Public Const id_LangAlamat As String = "Alamat"
Public Const id_LangNoTelp As String = "No. Telp"
Public Const id_LangDeposit As String = "Deposit"
Public Const id_LangLaporanRekapHarian As String = "Laporan Rekap Harian Tanggal "
Public Const id_LangLaporanRekapBulanan As String = "Laporan Rekap Bulanan Periode "
Public Const id_LangLaporanRekapTahunan As String = "Laporan Rekap Tahunan Periode "
'global
Public Const id_LangUser As String = "User"
Public Const id_LangNama As String = "Nama"
Public Const id_LangLogin As String = "Login"
Public Const id_LangGagalCache As String = "Cache Autorisasi Gagal"
Public Const id_LangNoClientTerpakai As String = "Nomor client telah terpakai"
Public Const id_LangClientBelumTerdaftar As String = "Client belum terdaftar"
Public Const id_LangMemberBelumTerdaftar As String = "Nama member belum terdaftar"
Public Const id_LangPassword As String = "Password"
Public Const id_LangPasswordLagi As String = "Password Lagi"
Public Const id_LangBahasa As String = "Bahasa"
Public Const id_LangSetPasswordSukses As String = "Set password sukses. Silahkan jalankan program lagi dan atur konfigurasi server."
Public Const id_LangPasswordSalah As String = "Password Salah"
Public Const id_LangVersi As String = "Versi"
Public Const id_LangGlobLogin As String = "Login"
Public Const id_LangBatal As String = "Batal"
Public Const id_LangGlobLogout As String = "Logout"
Public Const id_LangTidakDitemukan As String = "Tidak ditemukan dalam database"
Public Const id_LangUserSalah As String = "Nama user salah"
Public Const id_LangNomor As String = "Nomor"
Public Const id_LangStop As String = "Stop"
Public Const id_LangGlobPindah As String = "Pindah"
Public Const id_LangDari As String = "Dari"
Public Const id_LangKe As String = "Ke"
Public Const id_LangKonfigurasi As String = "Konfigurasi"
Public Const id_LangSetinganUmum As String = "Setingan Umum"
Public Const id_LangNamaWarnet As String = "Nama Warnet"
Public Const id_LangServerPort As String = "Port Server"
Public Const id_LangCurrency As String = "Mata Uang"
Public Const id_LangHargaAwal As String = "Harga Awal"
Public Const id_LangSetinganHarga As String = "Setingan Harga"
Public Const id_LangKelompokHarga As String = "Kelompok Harga"
Public Const id_LangHargaPersonal As String = "Personal"
Public Const id_LangHargaMember As String = "Member"
Public Const id_LangHargaGame As String = "Game"
Public Const id_LangHargaMengetik As String = "Mengetik"
Public Const id_LangInterval As String = "Interval"
Public Const id_LangMinDeposit As String = "Min Deposit"
Public Const id_LangMulaiJam As String = "Jam Mulai"
Public Const id_LangSelesaiJam As String = "Jam Selesai"
Public Const id_LangUpdate As String = "Update"
Public Const id_LangTutup As String = "Tutup"
Public Const id_LangRefresh As String = "Refresh"
Public Const id_LangDetailTransaksi As String = "Detail Transaksi"
Public Const id_LangCetakLap As String = "Cetak Lap."
Public Const id_LangCetakNota As String = "Cetak Nota"
Public Const id_LangPeriode As String = "Periode"
Public Const id_LangPerHari As String = "Per Hari"
Public Const id_LangPerBulan As String = "Per Bulan"
Public Const id_LangPerTahun As String = "Per Tahun"
Public Const id_LangSemua As String = "Semua"
Public Const id_LangJenisMember As String = "Member"
Public Const id_LangJenisPersonal As String = "Personal"
Public Const id_LangJenisGame As String = "Game"
Public Const id_LangJenisKetik As String = "Ketik"
Public Const id_LangRekapBilling As String = "Rekapitulasi Billing"
Public Const id_LangJanuari As String = "Januari"
Public Const id_LangFebruari As String = "Februari"
Public Const id_LangMaret As String = "Maret"
Public Const id_LangApril As String = "April"
Public Const id_LangMei As String = "Mei"
Public Const id_LangJuni As String = "Juni"
Public Const id_LangJuli As String = "Juli"
Public Const id_LangAgustus As String = "Agustus"
Public Const id_LangSeptember As String = "September"
Public Const id_LangOktober As String = "Oktober"
Public Const id_LangNovember As String = "November"
Public Const id_LangDesember As String = "Desember"
Public Const id_LangOperator As String = "Operator"
Public Const id_LangLaporanDeposit As String = "Laporan Deposit Member"
Public Const id_LangSimpan As String = "Simpan"
Public Const id_LangUbah As String = "Ubah"
Public Const id_LangHapus As String = "Hapus"
Public Const id_LangManajemenMember As String = "Manajemen Member"
Public Const id_LangNamaAsli As String = "Nama Asli"
Public Const id_LangSetAdminClient As String = "Set Password Admin Client"
Public Const id_LangUlangPassword As String = "Password Lagi"
Public Const id_LangKirimPesan As String = "Kirim dan Terima Pesan"
Public Const id_LangKirim As String = "Kirim"
Public Const id_LangHapusLog As String = "Hapus Log"
Public Const id_LangShutdownClient As String = "Shutdown Client"
Public Const id_LangShutdown As String = "Shutdown"
Public Const id_LangManajemenUser As String = "Manajemen User"
Public Const id_LangAdministrator As String = "Administrator"
Public Const id_LangJenisUser As String = "Jenis User"
Public Const id_StatusClientKonek As String = "Status Client : Mencoba Konek"
Public Const id_StatusClientSeting As String = "Status Client : Belum diseting"
Public Const id_LangTerkonek As String = "Status Client : Tersambung"
Public Const id_LangTerkunci As String = "Terkunci"
Public Const id_LangInformasi As String = "Informasi"
Public Const id_LangDepositTidakCukup As String = "Deposit tidak mencukupi"
Public Const id_LangUserKosong As String = "Nama User tidak boleh kosong"

'greeting
Public Const id_LangGreeting As String = "Terima Kasih Atas Kunjungan Anda"



'ENGLISH
'menu
Public Const en_LangmnuServer As String = "Server"
Public Const en_LangmnuMatikan As String = "Close"
Public Const en_LangGantiOp As String = "Change Operator"
Public Const en_LangPrintStruk As String = "Print Struct"
Public Const en_LangmnuClient As String = "Client"
Public Const en_LangmnuStop As String = "Stop Client"
Public Const en_LangmnuPindah As String = "Move Client"
Public Const en_LangmnuShutdown As String = "Shutdown Client"
Public Const en_LangmnuSetPasswordClient As String = "Set Password"
Public Const en_LangmnuSetting As String = "Setting"
Public Const en_LangmnuKonfig As String = "Configuration"
Public Const en_LangmnuUser As String = "User Management"
Public Const en_LangmnuLaporan As String = "Report"
Public Const en_LangmnuRekap As String = "Billing Report"
Public Const en_LangmnuDeposit As String = "Deposit Information"
Public Const en_LangmnuUtil As String = "Utility"
Public Const en_LangmnuMember As String = "Member Management"
Public Const en_LangmnuPesan As String = "Message"
Public Const en_LangmnuBackup As String = "Backup Database"
Public Const en_LangmnuHelp As String = "Help"
Public Const en_LangmnuAbout As String = "About"
Public Const en_Langmnuhapusclient As String = "Delete Client"
'toolbar
Public Const en_LangTbarMatikan As String = "Close Billing Server"
Public Const en_LangTbarStop As String = "Stop Client"
Public Const en_LangTbarPindah As String = "Move Client"
Public Const en_LangTbarShutdown As String = "Shutdown Client"
Public Const en_LangTbarSetPasswordClient As String = "Set Password Admin"
Public Const en_LangTbarKonfig As String = "Configuration"
Public Const en_LangTbarUser As String = "User Management"
Public Const en_LangTbarRekap As String = "Billing Report"
Public Const en_LangTbarDeposit As String = "Deposit Report"
Public Const en_LangTbarMember As String = "Member"
Public Const en_LangTbarPesan As String = "Message"
Public Const en_LangTbarBackup As String = "Backup Database"
Public Const en_LangTbarAbout As String = "About"
'jenis
Public Const en_LangPersonal As String = "Personal"
Public Const en_LangMember As String = "Member"
Public Const en_LangGame As String = "Game"
Public Const en_LangKetik As String = "Rent"
'status
Public Const en_LangBebas As String = "Idle"
Public Const en_LangAktif As String = "Active"
Public Const en_LangPindah As String = "Move"
'billing
Public Const en_LangClient As String = "Client "
Public Const en_LangNamaUser As String = "User Name"
Public Const en_LangStatus As String = "Status"
Public Const en_LangMulai As String = "Begin"
Public Const en_LangDurasi As String = "Duration"
Public Const en_LangBiaya As String = "Cost"
Public Const en_LangJenis As String = "Type"
Public Const en_LangDiscount As String = "Discount"
Public Const en_LangTotal As String = "Total"
Public Const en_LangNota As String = "Nota"
Public Const en_LangNo As String = "No."
Public Const en_LangTanggal As String = "Date"
Public Const en_LangSelesai As String = "End"
Public Const en_LangNamaMember As String = "Member Name"
Public Const en_LangAlamat As String = "Address"
Public Const en_LangNoTelp As String = "Telp No."
Public Const en_LangDeposit As String = "Deposit"
Public Const en_LangLaporanRekapHarian As String = "Daily Report at "
Public Const en_LangLaporanRekapBulanan As String = "Monthly Report Periode "
Public Const en_LangLaporanRekapTahunan As String = "Yearly Report Periode "
'global
Public Const en_LangUser As String = "User"
Public Const en_LangNama As String = "Name"
Public Const en_LangLogin As String = "Login"
Public Const en_LangGagalCache As String = "Cache Autorisation Failed"
Public Const en_LangNoClientTerpakai As String = "Client number are exist"
Public Const en_LangClientBelumTerdaftar As String = "Client not registered yet"
Public Const en_LangMemberBelumTerdaftar As String = "Member name not registered yet"
Public Const en_LangPassword As String = "Password"
Public Const en_LangPasswordLagi As String = "Retype Password"
Public Const en_LangBahasa As String = "Language"
Public Const en_LangSetPasswordSukses As String = "Set password succeed. Please restart program."
Public Const en_LangPasswordSalah As String = "Wrong Password"
Public Const en_LangVersi As String = "Version"
Public Const en_LangGlobLogin As String = "Login"
Public Const en_LangBatal As String = "Cancel"
Public Const en_LangGlobLogout As String = "Logout"
Public Const en_LangTidakDitemukan As String = "Doesnt exists in database"
Public Const en_LangUserSalah As String = "Wrong user name"
Public Const en_LangNomor As String = "Number"
Public Const en_LangStop As String = "Stop"
Public Const en_LangGlobPindah As String = "Move"
Public Const en_LangDari As String = "From"
Public Const en_LangKe As String = "To"
Public Const en_LangKonfigurasi As String = "Configuration"
Public Const en_LangSetinganUmum As String = "Global Setting"
Public Const en_LangNamaWarnet As String = "Netcafe Name"
Public Const en_LangServerPort As String = "Server Port"
Public Const en_LangCurrency As String = "Currency"
Public Const en_LangHargaAwal As String = "1st Cost"
Public Const en_LangSetinganHarga As String = "Cost Setting"
Public Const en_LangKelompokHarga As String = "Cost Group"
Public Const en_LangHargaPersonal As String = "Personal"
Public Const en_LangHargaMember As String = "Member"
Public Const en_LangHargaGame As String = "Game"
Public Const en_LangHargaMengetik As String = "Rent"
Public Const en_LangInterval As String = "Interval"
Public Const en_LangMinDeposit As String = "Min Deposit"
Public Const en_LangMulaiJam As String = "Time Begin"
Public Const en_LangSelesaiJam As String = "Time End"
Public Const en_LangUpdate As String = "Update"
Public Const en_LangTutup As String = "Close"
Public Const en_LangRefresh As String = "Refresh"
Public Const en_LangDetailTransaksi As String = "Transaction Detail"
Public Const en_LangCetakLap As String = "Print Rep."
Public Const en_LangCetakNota As String = "Print Nota"
Public Const en_LangPeriode As String = "Periode"
Public Const en_LangPerHari As String = "Daily"
Public Const en_LangPerBulan As String = "Monthly"
Public Const en_LangPerTahun As String = "Yearly"
Public Const en_LangSemua As String = "All"
Public Const en_LangJenisMember As String = "Member"
Public Const en_LangJenisPersonal As String = "Personal"
Public Const en_LangJenisGame As String = "Game"
Public Const en_LangJenisKetik As String = "Rent"
Public Const en_LangRekapBilling As String = "Billing Report"
Public Const en_LangJanuari As String = "January"
Public Const en_LangFebruari As String = "February"
Public Const en_LangMaret As String = "March"
Public Const en_LangApril As String = "April"
Public Const en_LangMei As String = "May"
Public Const en_LangJuni As String = "June"
Public Const en_LangJuli As String = "July"
Public Const en_LangAgustus As String = "August"
Public Const en_LangSeptember As String = "September"
Public Const en_LangOktober As String = "October"
Public Const en_LangNovember As String = "November"
Public Const en_LangDesember As String = "December"
Public Const en_LangOperator As String = "Operator"
Public Const en_LangLaporanDeposit As String = "Member Deposit Report"
Public Const en_LangSimpan As String = "Save"
Public Const en_LangUbah As String = "Edit"
Public Const en_LangHapus As String = "Delete"
Public Const en_LangManajemenMember As String = "Member Management"
Public Const en_LangNamaAsli As String = "Real Name"
Public Const en_LangSetAdminClient As String = "Set Password Admin Client"
Public Const en_LangUlangPassword As String = "Retype Password"
Public Const en_LangKirimPesan As String = "Send and Receive Message"
Public Const en_LangKirim As String = "Send"
Public Const en_LangHapusLog As String = "Delete Log"
Public Const en_LangShutdownClient As String = "Shutdown Client"
Public Const en_LangShutdown As String = "Shutdown"
Public Const en_LangManajemenUser As String = "User Management"
Public Const en_LangAdministrator As String = "Administrator"
Public Const en_LangJenisUser As String = "User Type"
Public Const en_StatusClientKonek As String = "Status Client : Connecting"
Public Const en_StatusClientSeting As String = "Status Client : Not Registered Yet"
Public Const en_LangTerkonek As String = "Client Status : Connected"
Public Const en_LangTerkunci As String = "Locked"
Public Const en_LangInformasi As String = "Information"
Public Const en_LangDepositTidakCukup As String = "Not Enough Deposit"
Public Const en_LangUserKosong As String = "User name may not empty"

'greeting
Public Const en_LangGreeting As String = "Thank You For Your Visit"



Public LangmnuServer As String
Public LangmnuMatikan As String
Public LangmnuClient As String
Public LangmnuStop As String
Public LangmnuPindah As String
Public LangmnuShutdown As String
Public LangmnuSetPasswordClient As String
Public LangmnuSetting As String
Public LangmnuKonfig As String
Public LangmnuUser As String
Public LangmnuLaporan As String
Public LangmnuRekap As String
Public LangmnuDeposit As String
Public LangmnuUtil As String
Public LangmnuMember As String
Public LangmnuPesan As String
Public LangmnuBackup As String
Public LangmnuHelp As String
Public LangmnuAbout As String
Public Langmnuhapusclient As String
Public LangTbarMatikan As String
Public LangTbarStop As String
Public LangTbarPindah As String
Public LangTbarShutdown As String
Public LangTbarSetPasswordClient As String
Public LangTbarKonfig As String
Public LangTbarUser As String
Public LangTbarRekap As String
Public LangTbarDeposit As String
Public LangTbarMember As String
Public LangTbarPesan As String
Public LangTbarBackup As String
Public LangTbarAbout As String
Public LangPersonal As String
Public LangMember As String
Public LangGame As String
Public LangKetik As String
Public LangBebas As String
Public LangAktif As String
Public LangPindah As String
Public LangClient As String
Public LangNamaUser As String
Public LangStatus As String
Public LangMulai As String
Public LangDurasi As String
Public LangBiaya As String
Public LangJenis As String
Public LangDiscount As String
Public LangTotal As String
Public LangNota As String
Public LangNo As String
Public LangTanggal As String
Public LangSelesai As String
Public LangNamaMember As String
Public LangAlamat As String
Public LangNoTelp As String
Public LangDeposit As String
Public LangLaporanRekapHarian As String
Public LangLaporanRekapBulanan As String
Public LangLaporanRekapTahunan As String
Public LangUser As String
Public LangNama As String
Public LangLogin As String
Public LangGagalCache As String
Public LangNoClientTerpakai As String
Public LangClientBelumTerdaftar As String
Public LangMemberBelumTerdaftar As String
Public LangGreeting As String
Public LangGantiOp As String
Public LangPrintStruk As String
Public LangPassword As String
Public LangPasswordLagi As String
Public LangBahasa As String
Public LangSetPasswordSukses As String
Public LangPasswordSalah As String
Public LangVersi As String
Public LangGlobLogin As String
Public LangBatal As String
Public LangGlobLogout As String
Public LangTidakDitemukan As String
Public LangUserSalah As String
Public LangNomor As String
Public LangStop As String
Public LangGlobPindah As String
Public LangDari As String
Public LangKe As String
Public LangKonfigurasi As String
Public LangSetinganUmum As String
Public LangNamaWarnet As String
Public LangServerPort As String
Public LangCurrency As String
Public LangHargaAwal As String
Public LangSetinganHarga As String
Public LangKelompokHarga As String
Public LangHargaPersonal As String
Public LangHargaMember As String
Public LangHargaGame As String
Public LangHargaMengetik As String
Public LangInterval As String
Public LangMinDeposit As String
Public LangMulaiJam As String
Public LangSelesaiJam As String
Public LangUpdate As String
Public LangTutup As String
Public LangRefresh As String
Public LangDetailTransaksi As String
Public LangCetakLap As String
Public LangCetakNota As String
Public LangPeriode As String
Public LangPerHari As String
Public LangPerBulan As String
Public LangPerTahun As String
Public LangSemua As String
Public LangJenisPersonal As String
Public LangJenisMember As String
Public LangJenisGame As String
Public LangJenisKetik As String
Public LangRekapBilling As String
Public LangJanuari As String
Public LangFebruari As String
Public LangMaret As String
Public LangApril As String
Public LangMei As String
Public LangJuni As String
Public LangJuli As String
Public LangAgustus As String
Public LangSeptember As String
Public LangOktober As String
Public LangNovember As String
Public LangDesember As String
Public LangOperator As String
Public LangLaporanDeposit As String
Public LangSimpan As String
Public LangUbah As String
Public LangHapus As String
Public LangManajemenMember As String
Public LangNamaAsli As String
Public LangSetAdminClient As String
Public LangUlangPassword As String
Public LangKirimPesan As String
Public LangKirim As String
Public LangHapusLog As String
Public LangShutdownClient As String
Public LangShutdown As String
Public LangManajemenUser As String
Public LangAdministrator As String
Public LangJenisUser As String
Public LangStatusClientKonek As String
Public LangStatusClientSet As String
Public LangTerkonek As String
Public LangTerkunci As String
Public LangInformasi As String
Public LangDepositTidakCukup As String
Public LangUserKosong As String


Public Sub Bahasa()
    If lang = "id" Then
        LangmnuServer = id_LangmnuServer
        LangmnuMatikan = id_LangmnuMatikan
        LangGantiOp = id_LangGantiOp
        LangPrintStruk = id_LangPrintStruk
        LangmnuClient = id_LangmnuClient
        LangmnuStop = id_LangmnuStop
        LangmnuPindah = id_LangmnuPindah
        LangmnuShutdown = id_LangmnuShutdown
        LangmnuSetPasswordClient = id_LangmnuSetPasswordClient
        LangmnuSetting = id_LangmnuSetting
        LangmnuKonfig = id_LangmnuKonfig
        LangmnuUser = id_LangmnuUser
        LangmnuLaporan = id_LangmnuLaporan
        LangmnuRekap = id_LangmnuRekap
        LangmnuDeposit = id_LangmnuDeposit
        LangmnuUtil = id_LangmnuUtil
        LangmnuMember = id_LangmnuMember
        LangmnuPesan = id_LangmnuPesan
        LangmnuBackup = id_LangmnuBackup
        LangmnuHelp = id_LangmnuHelp
        LangmnuAbout = id_LangmnuAbout
        Langmnuhapusclient = id_Langmnuhapusclient
        LangTbarMatikan = id_LangTbarMatikan
        LangTbarStop = id_LangTbarStop
        LangTbarPindah = id_LangTbarPindah
        LangTbarShutdown = id_LangTbarShutdown
        LangTbarSetPasswordClient = id_LangTbarSetPasswordClient
        LangTbarKonfig = id_LangTbarKonfig
        LangTbarUser = id_LangTbarUser
        LangTbarRekap = id_LangTbarRekap
        LangTbarDeposit = id_LangTbarDeposit
        LangTbarMember = id_LangTbarMember
        LangTbarPesan = id_LangTbarPesan
        LangTbarBackup = id_LangTbarBackup
        LangTbarAbout = id_LangTbarAbout
        LangPersonal = id_LangPersonal
        LangMember = id_LangMember
        LangGame = id_LangGame
        LangKetik = id_LangKetik
        LangBebas = id_LangBebas
        LangAktif = id_LangAktif
        LangPindah = id_LangPindah
        LangClient = id_LangClient
        LangNamaUser = id_LangNamaUser
        LangStatus = id_LangStatus
        LangMulai = id_LangMulai
        LangDurasi = id_LangDurasi
        LangBiaya = id_LangBiaya
        LangJenis = id_LangJenis
        LangDiscount = id_LangDiscount
        LangTotal = id_LangTotal
        LangNota = id_LangNota
        LangNo = id_LangNo
        LangTanggal = id_LangTanggal
        LangSelesai = id_LangSelesai
        LangNamaMember = id_LangNamaMember
        LangAlamat = id_LangAlamat
        LangNoTelp = id_LangNoTelp
        LangDeposit = id_LangDeposit
        LangLaporanRekapHarian = id_LangLaporanRekapHarian
        LangLaporanRekapBulanan = id_LangLaporanRekapBulanan
        LangLaporanRekapTahunan = id_LangLaporanRekapTahunan
        LangUser = id_LangUser
        LangNama = id_LangNama
        LangLogin = id_LangLogin
        LangGagalCache = id_LangGagalCache
        LangNoClientTerpakai = id_LangNoClientTerpakai
        LangClientBelumTerdaftar = id_LangClientBelumTerdaftar
        LangMemberBelumTerdaftar = id_LangMemberBelumTerdaftar
        LangPassword = id_LangPassword
        LangPasswordLagi = id_LangPasswordLagi
        LangBahasa = id_LangBahasa
        LangSetPasswordSukses = id_LangSetPasswordSukses
        LangPasswordSalah = id_LangPasswordSalah
        LangVersi = id_LangVersi
        LangGlobLogin = id_LangGlobLogin
        LangBatal = id_LangBatal
        LangGlobLogout = id_LangGlobLogout
        LangTidakDitemukan = id_LangTidakDitemukan
        LangUserSalah = id_LangUserSalah
        LangNomor = id_LangNomor
        LangStop = id_LangStop
        LangGlobPindah = id_LangGlobPindah
        LangDari = id_LangDari
        LangKe = id_LangKe
        LangKonfigurasi = id_LangKonfigurasi
        LangSetinganUmum = id_LangSetinganUmum
        LangNamaWarnet = id_LangNamaWarnet
        LangServerPort = id_LangServerPort
        LangCurrency = id_LangCurrency
        LangHargaAwal = id_LangHargaAwal
        LangSetinganHarga = id_LangSetinganHarga
        LangKelompokHarga = id_LangKelompokHarga
        LangHargaPersonal = id_LangHargaPersonal
        LangHargaMember = id_LangHargaMember
        LangHargaGame = id_LangHargaGame
        LangHargaMengetik = id_LangHargaMengetik
        LangInterval = id_LangInterval
        LangMinDeposit = id_LangMinDeposit
        LangMulaiJam = id_LangMulaiJam
        LangSelesaiJam = id_LangSelesaiJam
        LangUpdate = id_LangUpdate
        LangTutup = id_LangTutup
        LangRefresh = id_LangRefresh
        LangDetailTransaksi = id_LangDetailTransaksi
        LangCetakLap = id_LangCetakLap
        LangCetakNota = id_LangCetakNota
        LangPeriode = id_LangPeriode
        LangPerHari = id_LangPerHari
        LangPerBulan = id_LangPerBulan
        LangPerTahun = id_LangPerTahun
        LangSemua = id_LangSemua
        LangJenisPersonal = id_LangJenisPersonal
        LangJenisMember = id_LangJenisMember
        LangJenisGame = id_LangJenisGame
        LangJenisKetik = id_LangJenisKetik
        LangRekapBilling = id_LangRekapBilling
        LangJanuari = id_LangJanuari
        LangFebruari = id_LangFebruari
        LangMaret = id_LangMaret
        LangApril = id_LangApril
        LangMei = id_LangMei
        LangJuni = id_LangJuni
        LangJuli = id_LangJuli
        LangAgustus = id_LangAgustus
        LangSeptember = id_LangSeptember
        LangOktober = id_LangOktober
        LangNovember = id_LangNovember
        LangDesember = id_LangDesember
        LangOperator = id_LangOperator
        LangLaporanDeposit = id_LangLaporanDeposit
        LangSimpan = id_LangSimpan
        LangUbah = id_LangUbah
        LangHapus = id_LangHapus
        LangManajemenMember = id_LangManajemenMember
        LangNamaAsli = id_LangNamaAsli
        LangSetAdminClient = id_LangSetAdminClient
        LangUlangPassword = id_LangUlangPassword
        LangKirimPesan = id_LangKirimPesan
        LangKirim = id_LangKirim
        LangHapusLog = id_LangHapusLog
        LangShutdown = id_LangShutdown
        LangShutdownClient = id_LangShutdownClient
        LangManajemenUser = id_LangManajemenUser
        LangAdministrator = id_LangAdministrator
        LangJenisUser = id_LangJenisUser
        LangStatusClientKonek = id_StatusClientKonek
        LangStatusClientSet = id_StatusClientSeting
        LangTerkonek = id_LangTerkonek
        LangTerkunci = id_LangTerkunci
        LangInformasi = id_LangInformasi
        LangDepositTidakCukup = id_LangDepositTidakCukup
        LangUserKosong = id_LangUserKosong
    ElseIf lang = "en" Then
        LangmnuServer = en_LangmnuServer
        LangmnuMatikan = en_LangmnuMatikan
        LangGantiOp = en_LangGantiOp
        LangPrintStruk = en_LangPrintStruk
        LangmnuClient = en_LangmnuClient
        LangmnuStop = en_LangmnuStop
        LangmnuPindah = en_LangmnuPindah
        LangmnuShutdown = en_LangmnuShutdown
        LangmnuSetPasswordClient = en_LangmnuSetPasswordClient
        LangmnuSetting = en_LangmnuSetting
        LangmnuKonfig = en_LangmnuKonfig
        LangmnuUser = en_LangmnuUser
        LangmnuLaporan = en_LangmnuLaporan
        LangmnuRekap = en_LangmnuRekap
        LangmnuDeposit = en_LangmnuDeposit
        LangmnuUtil = en_LangmnuUtil
        LangmnuMember = en_LangmnuMember
        LangmnuPesan = en_LangmnuPesan
        LangmnuBackup = en_LangmnuBackup
        LangmnuHelp = en_LangmnuHelp
        LangmnuAbout = en_LangmnuAbout
        Langmnuhapusclient = en_Langmnuhapusclient
        LangTbarMatikan = en_LangTbarMatikan
        LangTbarStop = en_LangTbarStop
        LangTbarPindah = en_LangTbarPindah
        LangTbarShutdown = en_LangTbarShutdown
        LangTbarSetPasswordClient = en_LangTbarSetPasswordClient
        LangTbarKonfig = en_LangTbarKonfig
        LangTbarUser = en_LangTbarUser
        LangTbarRekap = en_LangTbarRekap
        LangTbarDeposit = en_LangTbarDeposit
        LangTbarMember = en_LangTbarMember
        LangTbarPesan = en_LangTbarPesan
        LangTbarBackup = en_LangTbarBackup
        LangTbarAbout = en_LangTbarAbout
        LangPersonal = en_LangPersonal
        LangMember = en_LangMember
        LangGame = en_LangGame
        LangKetik = en_LangKetik
        LangBebas = en_LangBebas
        LangAktif = en_LangAktif
        LangPindah = en_LangPindah
        LangClient = en_LangClient
        LangNamaUser = en_LangNamaUser
        LangStatus = en_LangStatus
        LangMulai = en_LangMulai
        LangDurasi = en_LangDurasi
        LangBiaya = en_LangBiaya
        LangJenis = en_LangJenis
        LangDiscount = en_LangDiscount
        LangTotal = en_LangTotal
        LangNota = en_LangNota
        LangNo = en_LangNo
        LangTanggal = en_LangTanggal
        LangSelesai = en_LangSelesai
        LangNamaMember = en_LangNamaMember
        LangAlamat = en_LangAlamat
        LangNoTelp = en_LangNoTelp
        LangDeposit = en_LangDeposit
        LangLaporanRekapHarian = en_LangLaporanRekapHarian
        LangLaporanRekapBulanan = en_LangLaporanRekapBulanan
        LangLaporanRekapTahunan = en_LangLaporanRekapTahunan
        LangUser = en_LangUser
        LangNama = en_LangNama
        LangLogin = en_LangLogin
        LangGagalCache = en_LangGagalCache
        LangNoClientTerpakai = en_LangNoClientTerpakai
        LangClientBelumTerdaftar = en_LangClientBelumTerdaftar
        LangMemberBelumTerdaftar = en_LangMemberBelumTerdaftar
        LangPassword = en_LangPassword
        LangPasswordLagi = en_LangPasswordLagi
        LangBahasa = en_LangBahasa
        LangSetPasswordSukses = en_LangSetPasswordSukses
        LangPasswordSalah = en_LangPasswordSalah
        LangVersi = en_LangVersi
        LangGlobLogin = en_LangGlobLogin
        LangBatal = en_LangBatal
        LangGlobLogout = en_LangGlobLogout
        LangTidakDitemukan = en_LangTidakDitemukan
        LangUserSalah = en_LangUserSalah
        LangNomor = en_LangNomor
        LangStop = en_LangStop
        LangGlobPindah = en_LangGlobPindah
        LangDari = en_LangDari
        LangKe = en_LangKe
        LangKonfigurasi = en_LangKonfigurasi
        LangSetinganUmum = en_LangSetinganUmum
        LangNamaWarnet = en_LangNamaWarnet
        LangServerPort = en_LangServerPort
        LangCurrency = en_LangCurrency
        LangHargaAwal = en_LangHargaAwal
        LangSetinganHarga = en_LangSetinganHarga
        LangKelompokHarga = en_LangKelompokHarga
        LangHargaPersonal = en_LangHargaPersonal
        LangHargaMember = en_LangHargaMember
        LangHargaGame = en_LangHargaGame
        LangHargaMengetik = en_LangHargaMengetik
        LangInterval = en_LangInterval
        LangMinDeposit = en_LangMinDeposit
        LangMulaiJam = en_LangMulaiJam
        LangSelesaiJam = en_LangSelesaiJam
        LangUpdate = en_LangUpdate
        LangTutup = en_LangTutup
        LangRefresh = en_LangRefresh
        LangDetailTransaksi = en_LangDetailTransaksi
        LangCetakLap = en_LangCetakLap
        LangCetakNota = en_LangCetakNota
        LangPeriode = en_LangPeriode
        LangPerHari = en_LangPerHari
        LangPerBulan = en_LangPerBulan
        LangPerTahun = en_LangPerTahun
        LangSemua = en_LangSemua
        LangJenisPersonal = en_LangJenisPersonal
        LangJenisMember = en_LangJenisMember
        LangJenisGame = en_LangJenisGame
        LangJenisKetik = en_LangJenisKetik
        LangRekapBilling = en_LangRekapBilling
        LangJanuari = en_LangJanuari
        LangFebruari = en_LangFebruari
        LangMaret = en_LangMaret
        LangApril = en_LangApril
        LangMei = en_LangMei
        LangJuni = en_LangJuni
        LangJuli = en_LangJuli
        LangAgustus = en_LangAgustus
        LangSeptember = en_LangSeptember
        LangOktober = en_LangOktober
        LangNovember = en_LangNovember
        LangDesember = en_LangDesember
        LangOperator = en_LangOperator
        LangLaporanDeposit = en_LangLaporanDeposit
        LangSimpan = en_LangSimpan
        LangUbah = en_LangUbah
        LangHapus = en_LangHapus
        LangManajemenMember = en_LangManajemenMember
        LangNamaAsli = en_LangNamaAsli
        LangSetAdminClient = en_LangSetAdminClient
        LangUlangPassword = en_LangUlangPassword
        LangKirimPesan = en_LangKirimPesan
        LangKirim = en_LangKirim
        LangHapusLog = en_LangHapusLog
        LangShutdown = en_LangShutdown
        LangShutdownClient = en_LangShutdownClient
        LangManajemenUser = en_LangManajemenUser
        LangAdministrator = en_LangAdministrator
        LangJenisUser = en_LangJenisUser
        LangStatusClientKonek = en_StatusClientKonek
        LangStatusClientSet = en_StatusClientSeting
        LangTerkonek = en_LangTerkonek
        LangTerkunci = en_LangTerkunci
        LangInformasi = en_LangInformasi
        LangDepositTidakCukup = en_LangDepositTidakCukup
        LangUserKosong = en_LangUserKosong
    Else
        LangmnuServer = id_LangmnuServer
        LangmnuMatikan = id_LangmnuMatikan
        LangGantiOp = id_LangGantiOp
        LangPrintStruk = id_LangPrintStruk
        LangmnuClient = id_LangmnuClient
        LangmnuStop = id_LangmnuStop
        LangmnuPindah = id_LangmnuPindah
        LangmnuShutdown = id_LangmnuShutdown
        LangmnuSetPasswordClient = id_LangmnuSetPasswordClient
        LangmnuSetting = id_LangmnuSetting
        LangmnuKonfig = id_LangmnuKonfig
        LangmnuUser = id_LangmnuUser
        LangmnuLaporan = id_LangmnuLaporan
        LangmnuRekap = id_LangmnuRekap
        LangmnuDeposit = id_LangmnuDeposit
        LangmnuUtil = id_LangmnuUtil
        LangmnuMember = id_LangmnuMember
        LangmnuPesan = id_LangmnuPesan
        LangmnuBackup = id_LangmnuBackup
        LangmnuHelp = id_LangmnuHelp
        LangmnuAbout = id_LangmnuAbout
        Langmnuhapusclient = id_Langmnuhapusclient
        LangTbarMatikan = id_LangTbarMatikan
        LangTbarStop = id_LangTbarStop
        LangTbarPindah = id_LangTbarPindah
        LangTbarShutdown = id_LangTbarShutdown
        LangTbarSetPasswordClient = id_LangTbarSetPasswordClient
        LangTbarKonfig = id_LangTbarKonfig
        LangTbarUser = id_LangTbarUser
        LangTbarRekap = id_LangTbarRekap
        LangTbarDeposit = id_LangTbarDeposit
        LangTbarMember = id_LangTbarMember
        LangTbarPesan = id_LangTbarPesan
        LangTbarBackup = id_LangTbarBackup
        LangTbarAbout = id_LangTbarAbout
        LangPersonal = id_LangPersonal
        LangMember = id_LangMember
        LangGame = id_LangGame
        LangKetik = id_LangKetik
        LangBebas = id_LangBebas
        LangAktif = id_LangAktif
        LangPindah = id_LangPindah
        LangClient = id_LangClient
        LangNamaUser = id_LangNamaUser
        LangStatus = id_LangStatus
        LangMulai = id_LangMulai
        LangDurasi = id_LangDurasi
        LangBiaya = id_LangBiaya
        LangJenis = id_LangJenis
        LangDiscount = id_LangDiscount
        LangTotal = id_LangTotal
        LangNota = id_LangNota
        LangNo = id_LangNo
        LangTanggal = id_LangTanggal
        LangSelesai = id_LangSelesai
        LangNamaMember = id_LangNamaMember
        LangAlamat = id_LangAlamat
        LangNoTelp = id_LangNoTelp
        LangDeposit = id_LangDeposit
        LangLaporanRekapHarian = id_LangLaporanRekapHarian
        LangLaporanRekapBulanan = id_LangLaporanRekapBulanan
        LangLaporanRekapTahunan = id_LangLaporanRekapTahunan
        LangUser = id_LangUser
        LangNama = id_LangNama
        LangLogin = id_LangLogin
        LangGagalCache = id_LangGagalCache
        LangNoClientTerpakai = id_LangNoClientTerpakai
        LangClientBelumTerdaftar = id_LangClientBelumTerdaftar
        LangMemberBelumTerdaftar = id_LangMemberBelumTerdaftar
        LangPassword = id_LangPassword
        LangPasswordLagi = id_LangPasswordLagi
        LangBahasa = id_LangBahasa
        LangSetPasswordSukses = id_LangSetPasswordSukses
        LangPasswordSalah = id_LangPasswordSalah
        LangVersi = id_LangVersi
        LangGlobLogin = id_LangGlobLogin
        LangBatal = id_LangBatal
        LangGlobLogout = id_LangGlobLogout
        LangTidakDitemukan = id_LangTidakDitemukan
        LangUserSalah = id_LangUserSalah
        LangNomor = id_LangNomor
        LangStop = id_LangStop
        LangGlobPindah = id_LangGlobPindah
        LangDari = id_LangDari
        LangKe = id_LangKe
        LangKonfigurasi = id_LangKonfigurasi
        LangSetinganUmum = id_LangSetinganUmum
        LangNamaWarnet = id_LangNamaWarnet
        LangServerPort = id_LangServerPort
        LangCurrency = id_LangCurrency
        LangHargaAwal = id_LangHargaAwal
        LangSetinganHarga = id_LangSetinganHarga
        LangKelompokHarga = id_LangKelompokHarga
        LangHargaPersonal = id_LangHargaPersonal
        LangHargaMember = id_LangHargaMember
        LangHargaGame = id_LangHargaGame
        LangHargaMengetik = id_LangHargaMengetik
        LangInterval = id_LangInterval
        LangMinDeposit = id_LangMinDeposit
        LangMulaiJam = id_LangMulaiJam
        LangSelesaiJam = id_LangSelesaiJam
        LangUpdate = id_LangUpdate
        LangTutup = id_LangTutup
        LangRefresh = id_LangRefresh
        LangDetailTransaksi = id_LangDetailTransaksi
        LangCetakLap = id_LangCetakLap
        LangCetakNota = id_LangCetakNota
        LangPeriode = id_LangPeriode
        LangPerHari = id_LangPerHari
        LangPerBulan = id_LangPerBulan
        LangPerTahun = id_LangPerTahun
        LangSemua = id_LangSemua
        LangJenisPersonal = id_LangJenisPersonal
        LangJenisMember = id_LangJenisMember
        LangJenisGame = id_LangJenisGame
        LangJenisKetik = id_LangJenisKetik
        LangRekapBilling = id_LangRekapBilling
        LangJanuari = id_LangJanuari
        LangFebruari = id_LangFebruari
        LangMaret = id_LangMaret
        LangApril = id_LangApril
        LangMei = id_LangMei
        LangJuni = id_LangJuni
        LangJuli = id_LangJuli
        LangAgustus = id_LangAgustus
        LangSeptember = id_LangSeptember
        LangOktober = id_LangOktober
        LangNovember = id_LangNovember
        LangDesember = id_LangDesember
        LangOperator = id_LangOperator
        LangLaporanDeposit = id_LangLaporanDeposit
        LangSimpan = id_LangSimpan
        LangUbah = id_LangUbah
        LangHapus = id_LangHapus
        LangManajemenMember = id_LangManajemenMember
        LangNamaAsli = id_LangNamaAsli
        LangSetAdminClient = id_LangSetAdminClient
        LangUlangPassword = id_LangUlangPassword
        LangKirimPesan = id_LangKirimPesan
        LangKirim = id_LangKirim
        LangHapusLog = id_LangHapusLog
        LangShutdown = id_LangShutdown
        LangShutdownClient = id_LangShutdownClient
        LangManajemenUser = id_LangManajemenUser
        LangAdministrator = id_LangAdministrator
        LangJenisUser = id_LangJenisUser
        LangStatusClientKonek = id_StatusClientKonek
        LangStatusClientSet = id_StatusClientSeting
        LangTerkonek = id_LangTerkonek
        LangTerkunci = id_LangTerkunci
        LangInformasi = id_LangInformasi
        LangDepositTidakCukup = id_LangDepositTidakCukup
        LangUserKosong = id_LangUserKosong
    End If
End Sub

