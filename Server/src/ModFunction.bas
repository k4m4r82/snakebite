Attribute VB_Name = "ModFunction"
Type DbSeting
    nama_warnet As String
    alamat1 As String
    alamat2  As String
    port As Long
    curr As String
    harga_awal As Single
End Type

Type DbHarga
    personal As Single
    member As Single
    game As Single
    ketik As Single
    interval As Long
    depositmin As Single
End Type

Type DbDiskon
    disc As Long
    jam_awal As String
    jam_akhir As String
End Type

Public TDbSeting As DbSeting
Public TDbHarga As DbHarga
Public TDbDiskon As DbDiskon

Function UpdateDbSeting(nama_warnet As String, alamat1 As String, alamat2 As String, port As Long, curr As String, harga_awal As Single) As Long
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open "SELECT seting.* FROM seting", cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        rs.MoveFirst
        rs(0) = nama_warnet
        rs(1) = alamat1
        rs(2) = alamat2
        rs(3) = port
        rs(4) = curr
        rs(5) = harga_awal
        rs.Update
    Else
        rs.AddNew
        rs(0) = nama_warnet
        rs(1) = alamat1
        rs(2) = alamat2
        rs(3) = port
        rs(4) = curr
        rs(5) = harga_awal
        rs.Update
    End If
    cn.Close
    UpdateDbSeting = 1
End Function

Function LihatDbSeting() As Long
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open "SELECT seting.* FROM seting", cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        rs.MoveFirst
        TDbSeting.nama_warnet = rs(0)
        TDbSeting.alamat1 = rs(1)
        TDbSeting.alamat2 = rs(2)
        TDbSeting.port = rs(3)
        TDbSeting.curr = rs(4)
        TDbSeting.harga_awal = rs(5)
        LihatDbSeting = 1
    Else
        TDbSeting.nama_warnet = "Warnet Citra Mandiri"
        TDbSeting.alamat1 = "Jl. Dewi Sartika No. 4"
        TDbSeting.alamat2 = "Abepura - Jayapura"
        TDbSeting.port = 1500
        TDbSeting.curr = "Rp."
        TDbSeting.harga_awal = 1000
        LihatDbSeting = 0
    End If
    cn.Close
End Function

Function LihatDbHarga() As Long
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open "SELECT harga.* FROM harga", cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        rs.MoveFirst
        TDbHarga.personal = rs(1)
        TDbHarga.member = rs(2)
        TDbHarga.game = rs(3)
        TDbHarga.ketik = rs(4)
        TDbHarga.interval = rs(5)
        TDbHarga.depositmin = rs(6)
        LihatDbHarga = 1
    Else
        TDbHarga.personal = 2500
        TDbHarga.member = 2000
        TDbHarga.game = 2000
        TDbHarga.ketik = 1000
        TDbHarga.interval = 15
        TDbHarga.depositmin = 25000
        LihatDbHarga = 0
    End If
    cn.Close
End Function

Function UpdateDbHarga(personal As Single, member As Single, game As Single, ketik As Single, interval As Long, depositmin As Single) As Long
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open "SELECT harga.* FROM harga", cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        rs.MoveFirst
        rs(0) = 0
        rs(1) = personal
        rs(2) = member
        rs(3) = game
        rs(4) = ketik
        rs(5) = interval
        rs(6) = depositmin
        rs.Update
    Else
        rs.AddNew
        rs(0) = 0
        rs(1) = personal
        rs(2) = member
        rs(3) = game
        rs(4) = ketik
        rs(5) = interval
        rs(6) = depositmin
        rs.Update
    End If
    cn.Close
    UpdateDbHarga = 1
End Function

Function LihatDbDiskon() As Long
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open "SELECT discount.* FROM discount", cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        rs.MoveFirst
        TDbDiskon.disc = rs(0)
        TDbDiskon.jam_awal = rs(1)
        TDbDiskon.jam_akhir = rs(2)
    Else
        TDbDiskon.disc = 10
        TDbDiskon.jam_awal = "22:00:00"
        TDbDiskon.jam_akhir = "04:00:00"
    End If
    cn.Close
    LihatDbDiskon = 1
End Function

Function UpdateDbDiskon(disc As Long, jam_awal As String, jam_akhir As String) As Long
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open "SELECT discount.* FROM discount", cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        rs.MoveFirst
        rs(0) = disc
        rs(1) = jam_awal
        rs(2) = jam_akhir
        rs.Update
    Else
        rs.AddNew
        rs(0) = 10
        rs(1) = jam_awal
        rs(2) = jam_akhir
        rs.Update
    End If
    cn.Close
    UpdateDbDiskon = 1
End Function

Function AmbilAplikasi() As Long
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open "SELECT aplikasi.* FROM aplikasi", cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        rs.MoveFirst
        ketik1 = rs(0)
        ketik2 = rs(1)
        ketik3 = rs(2)
        game1 = rs(3)
        game2 = rs(4)
        game3 = rs(5)
    Else
        ketik1 = "WINWORD.EXE"
        ketik2 = "EXCEL.EXE"
        ketik3 = "POWERPOINT.EXE"
        game1 = "CS.EXE"
        game2 = "GP2.EXE"
        game3 = "dfx.exe"
        rs.AddNew
        rs(0) = ketik1
        rs(1) = ketik2
        rs(2) = ketik3
        rs(3) = game1
        rs(4) = game2
        rs(5) = game3
        rs.Update
    End If
    cn.Close
    AmbilAplikasi = 1
End Function

Function UpdateDbApp(ketik1 As String, ketik2 As String, ketik3 As String, game1 As String, game2 As String, game3 As String) As Long
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open "SELECT aplikasi.* FROM aplikasi", cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        rs.MoveFirst
        rs(0) = ketik1
        rs(1) = ketik2
        rs(2) = ketik3
        rs(3) = game1
        rs(4) = game2
        rs(5) = game3
        rs.Update
    Else
        rs.AddNew
        rs(0) = ketik1
        rs(1) = ketik2
        rs(2) = ketik3
        rs(3) = game1
        rs(4) = game2
        rs(5) = game3
        rs.Update
    End If
    cn.Close
    UpdateDbApp = 1
End Function

Function CekClientKonek(ByVal no As Integer) As String
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open "SELECT ip.no_client,socket.socket FROM ip,socket WHERE ip.ip=socket.ip AND ip.no_client=" & no, cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        CekClientKonek = LangAktif
    Else
        CekClientKonek = LangInAktif
    End If
    cn.Close
End Function

Function CekClientSocket(ByVal no As Integer) As Integer
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open "SELECT ip.no_client,socket.socket FROM ip,socket WHERE ip.ip=socket.ip AND ip.no_client=" & no, cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        CekClientSocket = rs(1)
    End If
    cn.Close
End Function

Function JmlHari(ByVal vDate As Date) As Integer
    JmlHari = Day(DateSerial(Year(vDate), Month(vDate) + 1, 0))
End Function

Sub CtrlHover(ctl As Control, X As Single, Y As Single, OrigBackColor As Long, NewBackColor As Long)
Dim HitTest As Long
    On Error Resume Next
    HitTest = ctl.hwnd
    If Err.Number <> 0 Then Exit Sub
    If (X < 0) Or (Y < 0) Or (X > ctl.Width) Or (Y > ctl.Height) Then
        ReleaseCapture
        ctl.Backcolor = OrigBackColor
    Else
        SetCapture ctl.hwnd
        ctl.Backcolor = NewBackColor
    End If
    On Error GoTo 0
End Sub

Function Terbilang(ByVal n As Currency, ByVal lng As String) As String
Dim satuan, satuan1 As Variant
    If lng = "id" Then
        satuan = Array("", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan", "Sepuluh", "Sebelas")
        Select Case n
            Case 0 To 11
                Terbilang = " " + satuan(Fix(n))
            Case 12 To 19
                Terbilang = Terbilang(n Mod 10, lng) + " Belas"
            Case 20 To 99
                Terbilang = Terbilang(Fix(n / 10), lng) + " Puluh" + Terbilang(n Mod 10, lng)
            Case 100 To 199
                Terbilang = " Seratus" + Terbilang(n - 100, lng)
            Case 200 To 999
                Terbilang = Terbilang(Fix(n / 100), lng) + " Ratus" + Terbilang(n Mod 100, lng)
            Case 1000 To 1999
                Terbilang = " Seribu" + Terbilang(n - 1000, lng)
            Case 2000 To 999999
                Terbilang = Terbilang(Fix(n / 1000), lng) + " Ribu" + Terbilang(n Mod 1000, lng)
            Case 1000000 To 999999999
                Terbilang = Terbilang(Fix(n / 1000000), lng) + " Juta" + Terbilang(n Mod 1000000, lng)
            Case Else
                Terbilang = Terbilang(Fix(n / 1000000000), lng) + " Milyar" + Terbilang(n Mod 1000000000, lng)
        End Select
    ElseIf lng = "en" Then
        satuan = Array("", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve")
        satuan1 = Array("", "One", "Twen", "Thir", "Four", "Fif", "Six", "Seven", "Eigh", "Nine")
        Select Case n
            Case 0 To 12
                Terbilang = " " + satuan(Fix(n))
            Case 13 To 19
                Terbilang = satuan1(Fix(n - 10)) + "teen "
            Case 20 To 99
                Terbilang = satuan1(Fix(n / 10)) + "ty " + Terbilang(n Mod 10, lng)
            Case 100 To 199
                Terbilang = satuan1(Fix(n / 100)) + " Hundred " + Terbilang(n - 100, lng)
            Case 200 To 999
                Terbilang = Terbilang(Fix(n / 100), lng) + " Hundred " + Terbilang(n Mod 100, lng)
            Case 1000 To 1999
                Terbilang = Terbilang(Fix(n / 1000), lng) + " Thousand " + Terbilang(n - 1000, lng)
            Case 2000 To 999999
                Terbilang = Terbilang(Fix(n / 1000), lng) + " Thousand " + Terbilang(n Mod 1000, lng)
            Case 1000000 To 999999999
                Terbilang = Terbilang(Fix(n / 1000000), lng) + " Billion " + Terbilang(n Mod 1000000, lng)
            Case Else
                Terbilang = Terbilang(Fix(n / 1000000000), lng) + " Million " + Terbilang(n Mod 1000000000, lng)
        End Select
    End If
End Function

Sub TampilTip(ctl As Label, lpstr As String)
    ctl.Caption = lpstr
End Sub

Sub BuatDatabase(strDBPath As String)
    Dim catNewDB As ADOX.Catalog
    Set catNewDB = New ADOX.Catalog
    catNewDB.Create "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & strDBPath
    Set catNewDB = Nothing
End Sub

Sub BuatTabel(strDBPath As String)
Dim catDB As ADOX.Catalog
Dim tblNew As ADOX.Table
    Set catDB = New ADOX.Catalog
    catDB.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & strDBPath
    Set tblNew = New ADOX.Table
    tblNew.Name = "account"
    Set tblNew.ParentCatalog = catDB
    tblNew.Columns.Append "user", adVarWChar, 15
    tblNew.Columns.Append "nama", adVarWChar, 20
    tblNew.Columns.Append "login", adVarWChar, 10
    tblNew.Columns.Append "alamat", adVarWChar, 25
    tblNew.Columns.Append "telp", adVarWChar, 15
    tblNew.Columns.Append "password", adVarWChar, 15
    catDB.Tables.Append tblNew
    Set tblNew = Nothing
    
    Set tblNew = New ADOX.Table
    tblNew.Name = "admin"
    Set tblNew.ParentCatalog = catDB
    tblNew.Columns.Append "password_admin", adVarWChar, 15
    catDB.Tables.Append tblNew
    Set tblNew = Nothing
    
    Set tblNew = New ADOX.Table
    tblNew.Name = "aplikasi"
    Set tblNew.ParentCatalog = catDB
    tblNew.Columns.Append "ketik1", adVarWChar, 25
    tblNew.Columns.Append "ketik2", adVarWChar, 25
    tblNew.Columns.Append "ketik3", adVarWChar, 25
    tblNew.Columns.Append "game1", adVarWChar, 25
    tblNew.Columns.Append "game2", adVarWChar, 25
    tblNew.Columns.Append "game3", adVarWChar, 25
    catDB.Tables.Append tblNew
    Set tblNew = Nothing
    
    Set tblNew = New ADOX.Table
    tblNew.Name = "billing"
    Set tblNew.ParentCatalog = catDB
    tblNew.Columns.Append "no_nota", adInteger
    tblNew.Columns.Item("no_nota").Properties("AutoIncrement") = True
    tblNew.Columns.Append "tanggal", adVarWChar, 10
    tblNew.Columns.Append "no_client", adInteger
    tblNew.Columns.Append "user", adVarWChar, 25
    tblNew.Columns.Append "mulai", adVarWChar, 8
    tblNew.Columns.Append "selesai", adVarWChar, 8
    tblNew.Columns.Append "durasi", adVarWChar, 8
    tblNew.Columns.Append "total", adSingle
    tblNew.Columns.Append "bulan", adVarWChar, 2
    tblNew.Columns.Append "tahun", adVarWChar, 4
    tblNew.Columns.Append "jenis", adVarWChar, 15
    tblNew.Columns.Append "operator", adVarWChar, 25
    catDB.Tables.Append tblNew
    Set tblNew = Nothing
       
    Set tblNew = New ADOX.Table
    tblNew.Name = "cache"
    Set tblNew.ParentCatalog = catDB
    tblNew.Columns.Append "id", adInteger
    tblNew.Columns.Item("id").Properties("AutoIncrement") = True
    tblNew.Columns.Append "path", adVarWChar, 15
    tblNew.Columns.Append "user", adVarWChar, 25
    tblNew.Columns.Append "nama", adVarWChar, 50
    tblNew.Columns.Append "login", adVarWChar, 15
    catDB.Tables.Append tblNew
    Set tblNew = Nothing
        
    Set tblNew = New ADOX.Table
    tblNew.Name = "discount"
    Set tblNew.ParentCatalog = catDB
    tblNew.Columns.Append "disc", adInteger
    tblNew.Columns.Append "jam_awal", adVarWChar, 50
    tblNew.Columns.Append "jam_akhir", adVarWChar, 50
    catDB.Tables.Append tblNew
    Set tblNew = Nothing
    
    Set tblNew = New ADOX.Table
    tblNew.Name = "harga"
    Set tblNew.ParentCatalog = catDB
    tblNew.Columns.Append "admin", adSingle
    tblNew.Columns.Append "personal", adSingle
    tblNew.Columns.Append "member", adSingle
    tblNew.Columns.Append "game", adSingle
    tblNew.Columns.Append "ketik", adSingle
    tblNew.Columns.Append "interval", adInteger
    tblNew.Columns.Append "deposit_min", adSingle
    catDB.Tables.Append tblNew
    Set tblNew = Nothing
    
    Set tblNew = New ADOX.Table
    tblNew.Name = "ip"
    Set tblNew.ParentCatalog = catDB
    tblNew.Columns.Append "no_client", adInteger
    tblNew.Columns.Append "ip", adVarWChar, 50
    catDB.Tables.Append tblNew
    Set tblNew = Nothing
    
    Set tblNew = New ADOX.Table
    tblNew.Name = "member"
    Set tblNew.ParentCatalog = catDB
    tblNew.Columns.Append "user", adVarWChar, 15
    tblNew.Columns.Append "nama", adVarWChar, 20
    tblNew.Columns.Append "alamat", adVarWChar, 25
    tblNew.Columns.Append "telp", adVarWChar, 15
    tblNew.Columns.Append "password", adVarWChar, 15
    tblNew.Columns.Append "deposit", adSingle
    catDB.Tables.Append tblNew
    Set tblNew = Nothing
    
    Set tblNew = New ADOX.Table
    tblNew.Name = "proses"
    Set tblNew.ParentCatalog = catDB
    tblNew.Columns.Append "no_client", adInteger
    tblNew.Columns.Append "status", adVarWChar, 10
    tblNew.Columns.Append "jenis_login", adVarWChar, 15
    tblNew.Columns.Append "user", adVarWChar, 25
    tblNew.Columns.Append "mulai", adVarWChar, 30
    catDB.Tables.Append tblNew
    Set tblNew = Nothing
    
    Set tblNew = New ADOX.Table
    tblNew.Name = "seting"
    Set tblNew.ParentCatalog = catDB
    tblNew.Columns.Append "nama_warnet", adVarWChar, 35
    tblNew.Columns.Append "alamat1", adVarWChar, 35
    tblNew.Columns.Append "alamat2", adVarWChar, 35
    tblNew.Columns.Append "port", adInteger
    tblNew.Columns.Append "currency", adVarWChar, 2
    tblNew.Columns.Append "harga_awal", adSingle
    catDB.Tables.Append tblNew
    Set tblNew = Nothing
    
    Set tblNew = New ADOX.Table
    tblNew.Name = "socket"
    Set tblNew.ParentCatalog = catDB
    tblNew.Columns.Append "ip", adVarWChar, 50
    tblNew.Columns.Append "socket", adInteger
    catDB.Tables.Append tblNew
    Set tblNew = Nothing
    
    Set catDB = Nothing
End Sub

Sub SetDbPassword(strDBPath As String)
Dim db As DAO.Database
    Set db = DBEngine.OpenDatabase(strDBPath, True)
    db.NewPassword "", DefaultDbPass
    db.Close
End Sub

Sub GantiDbPass(strDBPath As String, oldpass As String, newpass As String)
Dim db As DAO.Database
    Set db = DBEngine.OpenDatabase(strDBPath, True, False, ";pwd=" & oldpass)
    db.NewPassword oldpass, newpass
    db.Close
End Sub

Sub CekLogin()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open "SELECT cache.* FROM cache WHERE cache.path = 'server'", cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        FrmMain.StatusBar.Panels(1) = LangUser
        FrmMain.StatusBar.Panels(2) = rs(2)
        FrmMain.StatusBar.Panels(3) = LangNama
        FrmMain.StatusBar.Panels(4) = rs(3)
        FrmMain.StatusBar.Panels(5) = LangLogin
        FrmMain.StatusBar.Panels(6) = rs(4)
        If rs(4) = "sysadmin" Then
            Call MenuAdmin
        ElseIf rs(4) = "admin" Then
            Call MenuAdmin
        ElseIf rs(4) = "user" Then
            Call MenuUser
        End If
        FrmMain.StatusBar.Panels(7) = Format(Date, "dddd, dd mmmm yyyy")
        FrmMain.StatusBar.Panels(8) = Format(Time, "hh:nn:ss")
    Else
        MsgBox LangGagalCache, vbCritical, "Sistem cache error!"
        End
    End If
    cn.Close
End Sub

Sub MenuAdmin()
    FrmMain.mnuconfig.Enabled = True
    FrmMain.mnuuser.Enabled = True
    FrmMain.mnubackup.Enabled = True
    FrmMain.mnumember.Enabled = True
    FrmMain.mnuregclient.Enabled = True
    FrmMain.mnudbpass.Enabled = True
    FrmMain.mnugantibahasa.Enabled = True
    FrmMain.Toolbar.Buttons(9).Enabled = True
    FrmMain.Toolbar.Buttons(10).Enabled = True
    FrmMain.Toolbar.Buttons(15).Enabled = True
    FrmMain.Toolbar.Buttons(17).Enabled = True
End Sub

Sub MenuUser()
    FrmMain.mnuconfig.Enabled = False
    FrmMain.mnuuser.Enabled = False
    FrmMain.mnubackup.Enabled = False
    FrmMain.mnumember.Enabled = False
    FrmMain.mnuregclient.Enabled = False
    FrmMain.mnudbpass.Enabled = False
    FrmMain.mnugantibahasa.Enabled = False
    FrmMain.Toolbar.Buttons(9).Enabled = False
    FrmMain.Toolbar.Buttons(10).Enabled = False
    FrmMain.Toolbar.Buttons(15).Enabled = False
    FrmMain.Toolbar.Buttons(17).Enabled = False
End Sub

Sub AturGrid()
    FrmMain.MSFlexGrid.Rows = ClientMax + 1
    FrmMain.MSFlexGrid.ColWidth(0) = 300
    For i = 1 To ClientMax
        FrmMain.MSFlexGrid.TextMatrix(i, 0) = i
    Next i
    FrmMain.MSFlexGrid.ColWidth(1) = 800
    FrmMain.MSFlexGrid.TextMatrix(0, 1) = LangClient
    FrmMain.MSFlexGrid.ColWidth(2) = 1500
    FrmMain.MSFlexGrid.TextMatrix(0, 2) = LangNamaUser
    FrmMain.MSFlexGrid.ColWidth(3) = 700
    FrmMain.MSFlexGrid.TextMatrix(0, 3) = LangStatus
    FrmMain.MSFlexGrid.ColWidth(4) = 900
    FrmMain.MSFlexGrid.TextMatrix(0, 4) = LangMulai
    FrmMain.MSFlexGrid.ColWidth(5) = 900
    FrmMain.MSFlexGrid.TextMatrix(0, 5) = LangSelesai
    FrmMain.MSFlexGrid.ColWidth(6) = 900
    FrmMain.MSFlexGrid.TextMatrix(0, 6) = LangDurasi
    FrmMain.MSFlexGrid.ColWidth(7) = 900
    FrmMain.MSFlexGrid.TextMatrix(0, 7) = LangBiaya
    FrmMain.MSFlexGrid.ColWidth(8) = 1000
    FrmMain.MSFlexGrid.TextMatrix(0, 8) = LangJenis
    FrmMain.MSFlexGrid.TextMatrix(0, 9) = LangDiscount
    FrmMain.MSFlexGrid.TextMatrix(0, 10) = LangTotal
End Sub
