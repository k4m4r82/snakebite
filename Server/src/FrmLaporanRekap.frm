VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmLaporanRekap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rekapitulasi Billing"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10140
   Icon            =   "FrmLaporanRekap.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   10140
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdHapus 
      Caption         =   "Hapus"
      Height          =   375
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton CmdCetakLaporan 
      Caption         =   "Cetak Lap."
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   " Jenis "
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   3000
      TabIndex        =   13
      Top             =   4560
      Width           =   3495
      Begin VB.OptionButton OpKetik 
         Appearance      =   0  'Flat
         Caption         =   "Ketik"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2520
         TabIndex        =   18
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OpGame 
         Appearance      =   0  'Flat
         Caption         =   "Game"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1320
         TabIndex        =   17
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton OpMember 
         Appearance      =   0  'Flat
         Caption         =   "Member"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1320
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton OpPersonal 
         Appearance      =   0  'Flat
         Caption         =   "Personal"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton OpSemua 
         Appearance      =   0  'Flat
         Caption         =   "Semua"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "Tutup"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton CmdCetakNota 
      Caption         =   "Cetak Nota"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton CmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame FrameJenisRekap 
      Appearance      =   0  'Flat
      Caption         =   " Periode "
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   4560
      Width           =   2655
      Begin VB.OptionButton OpTahun 
         Appearance      =   0  'Flat
         Caption         =   "Per Tahun"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton OpBulan 
         Appearance      =   0  'Flat
         Caption         =   "Per Bulan"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton OpHari 
         Appearance      =   0  'Flat
         Caption         =   "Per Hari"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.ComboBox TxtTahun 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "FrmLaporanRekap.frx":030A
      Left            =   9000
      List            =   "FrmLaporanRekap.frx":032F
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   4680
      Width           =   975
   End
   Begin VB.ComboBox TxtBulan 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "FrmLaporanRekap.frx":0375
      Left            =   7560
      List            =   "FrmLaporanRekap.frx":039D
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4680
      Width           =   1335
   End
   Begin VB.ComboBox TxtTgl 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "FrmLaporanRekap.frx":0404
      Left            =   6720
      List            =   "FrmLaporanRekap.frx":0406
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4680
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid FlexDataGrid 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   6800
      _Version        =   393216
      Cols            =   9
      ScrollBars      =   2
      Appearance      =   0
   End
   Begin VB.Label LabelTerbilang 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1440
      TabIndex        =   22
      Top             =   6120
      Width           =   8535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Terbilang"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total  :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7560
      TabIndex        =   2
      Top             =   5640
      Width           =   735
   End
   Begin VB.Label LabelTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8400
      TabIndex        =   1
      Top             =   5520
      Width           =   1575
   End
End
Attribute VB_Name = "FrmLaporanRekap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim jenis_login As String

Private Sub Init()
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim total As Single
    
    If Format(Date, "mm") = "01" Then
        TxtBulan.Text = LangJanuari
    ElseIf Format(Date, "mm") = "02" Then
        TxtBulan.Text = LangFebruari
    ElseIf Format(Date, "mm") = "03" Then
        TxtBulan.Text = LangMaret
    ElseIf Format(Date, "mm") = "04" Then
        TxtBulan.Text = LangApril
    ElseIf Format(Date, "mm") = "05" Then
        TxtBulan.Text = LangMei
    ElseIf Format(Date, "mm") = "06" Then
        TxtBulan.Text = LangJuni
    ElseIf Format(Date, "mm") = "07" Then
        TxtBulan.Text = LangJuli
    ElseIf Format(Date, "mm") = "08" Then
        TxtBulan.Text = LangAgustus
    ElseIf Format(Date, "mm") = "09" Then
        TxtBulan.Text = LangSeptember
    ElseIf Format(Date, "mm") = "10" Then
        TxtBulan.Text = LangOktober
    ElseIf Format(Date, "mm") = "11" Then
        TxtBulan.Text = LangNovember
    ElseIf Format(Date, "mm") = "12" Then
        TxtBulan.Text = LangDesember
    End If
    TxtBulan.ListIndex = Month(Now) - 1
    
    Call IsiTahun
    TxtTahun.Text = Format(Date, "yyyy")
    FlexDataGrid.ColWidth(0) = 500
    FlexDataGrid.ColWidth(1) = 1000
    FlexDataGrid.ColWidth(2) = 1000
    FlexDataGrid.ColWidth(3) = 1700
    FlexDataGrid.ColWidth(4) = 1000
    FlexDataGrid.ColWidth(5) = 1000
    FlexDataGrid.ColWidth(6) = 1000
    FlexDataGrid.ColWidth(7) = 1200
    FlexDataGrid.ColWidth(8) = 1000
    FlexDataGrid.TextMatrix(0, 0) = LangNo
    FlexDataGrid.TextMatrix(0, 1) = LangTanggal
    FlexDataGrid.TextMatrix(0, 2) = LangClient
    FlexDataGrid.TextMatrix(0, 3) = LangNamaUser
    FlexDataGrid.TextMatrix(0, 4) = LangMulai
    FlexDataGrid.TextMatrix(0, 5) = LangSelesai
    FlexDataGrid.TextMatrix(0, 6) = LangDurasi
    FlexDataGrid.TextMatrix(0, 7) = LangTotal
    FlexDataGrid.TextMatrix(0, 8) = LangOperator
    Call IsiTanggal
    TxtTgl.Text = Format(Date, "d")
    total = 0
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open "SELECT billing.* FROM billing WHERE billing.tanggal = '" & Format(Date, "dd-mm-yyyy") & "'", cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        FlexDataGrid.Rows = rs.RecordCount + 1
        rs.MoveFirst
        For i = 1 To rs.RecordCount
            FlexDataGrid.TextMatrix(i, 0) = rs(0) 'no nota
            FlexDataGrid.TextMatrix(i, 1) = rs(1) 'tanggal
            FlexDataGrid.TextMatrix(i, 2) = LangClient & " " & rs(2) 'no client
            FlexDataGrid.TextMatrix(i, 3) = rs(3) 'nama user
            FlexDataGrid.TextMatrix(i, 4) = rs(4) 'mulai
            FlexDataGrid.TextMatrix(i, 5) = rs(5) 'selesai
            FlexDataGrid.TextMatrix(i, 6) = rs(6) 'durasi
            FlexDataGrid.TextMatrix(i, 7) = Format(rs(7), "##,##0.00") 'total
            FlexDataGrid.TextMatrix(i, 8) = rs(11)
            total = total + rs(7)
            rs.MoveNext
        Next i
        LabelTotal.Caption = Format(total, "##,##0.00")
    Else
        FlexDataGrid.Rows = 1
    End If
    cn.Close
End Sub

Private Sub CmdCetakLaporan_Click()
    Printer.PrintQuality = 300
    Printer.FontSize = 10
    Printer.ScaleLeft = 10
    Printer.ScaleTop = 10
    Printer.Print Tab(10); FrmMain.Txtnama_warnet
    Printer.Print Tab(10); FrmMain.Txtalamat1
    Printer.Print Tab(10); FrmMain.Txtalamat2
    If OpHari.value = True Then
        Printer.Print Tab(10); LangLaporanRekapHarian & Format(CDate(TxtTgl.Text & "/" & TxtBulan.ListIndex + 1 & "/" & TxtTahun.Text), "dd dddd mmmm yyyy")
    ElseIf OpBulan.value = True Then
        Printer.Print Tab(10); LangLaporanRekapBulanan & Format(CDate(TxtTgl.Text & "/" & TxtBulan.ListIndex + 1 & "/" & TxtTahun.Text), "mmmm yyyy")
    ElseIf OpTahun.value = True Then
        Printer.Print Tab(10); LangLaporanRekapTahunan & Format(CDate(TxtTgl.Text & "/" & TxtBulan.ListIndex + 1 & "/" & TxtTahun.Text), "yyyy")
    Else
        Printer.Print Tab(10); LangLaporanRekapHarian & Format(Now, "dd dddd mmmm yyyy")
    End If
    Printer.Print Tab(10); "========================================================================================"
    Printer.Print Tab(10); LangNo & LangNota; Tab(20); LangTanggal; Tab(35); LangClient; Tab(45); "User ID"; Tab(60); LangMulai; Tab(70); LangSelesai; Tab(80); LangDurasi; Tab(90); LangTotal; Tab(110); "Operator"
    Printer.Print Tab(10); "========================================================================================"
    For i = 1 To FlexDataGrid.Rows - 1
        Printer.Print Tab(10); FlexDataGrid.TextMatrix(i, 0); Tab(20); FlexDataGrid.TextMatrix(i, 1); Tab(35); FlexDataGrid.TextMatrix(i, 2); Tab(45); FlexDataGrid.TextMatrix(i, 3); Tab(60); FlexDataGrid.TextMatrix(i, 4); Tab(70); FlexDataGrid.TextMatrix(i, 5); Tab(80); FlexDataGrid.TextMatrix(i, 6); Tab(90); FlexDataGrid.TextMatrix(i, 7); Tab(110); FlexDataGrid.TextMatrix(i, 8)
    Next i
    Printer.Print Tab(10); "========================================================================================"
    Printer.Print Tab(80); LangTotal; Tab(90); LabelTotal.Caption
    Printer.Print Tab(10); "========================================================================================"
    Printer.Print Tab(10); LangTerbilang; Tab(20); LabelTerbilang.Caption
    Printer.Print Tab(10); "========================================================================================"
    Printer.EndDoc
End Sub

Private Sub CmdCetakLaporan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdCetakLaporan, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdCetakNota_Click()
    Printer.PrintQuality = 300
    Printer.FontSize = 10
    Printer.ScaleLeft = 10
    Printer.ScaleTop = 10
    Printer.Print Tab(10); FrmMain.Txtnama_warnet
    Printer.Print Tab(10); FrmMain.Txtalamat1
    Printer.Print Tab(10); FrmMain.Txtalamat2
    Printer.Print Tab(10); "==============================="
    Printer.Print Tab(10); LangTanggal; Tab(30); ":"; Tab(35); FlexDataGrid.TextMatrix(FlexDataGrid.Row, 1)
    Printer.Print Tab(10); LangNo & LangNota; Tab(30); ":"; Tab(35); FlexDataGrid.TextMatrix(FlexDataGrid.Row, 0)
    Printer.Print Tab(10); LangNamaUser; Tab(30); ":"; Tab(35); FlexDataGrid.TextMatrix(FlexDataGrid.Row, 3)
    Printer.Print Tab(10); LangClient; Tab(30); ":"; Tab(35); FlexDataGrid.TextMatrix(FlexDataGrid.Row, 2)
    Printer.Print Tab(10); LangMulai; Tab(30); ":"; Tab(35); FlexDataGrid.TextMatrix(FlexDataGrid.Row, 4)
    Printer.Print Tab(10); LangSelesai; Tab(30); ":"; Tab(35); FlexDataGrid.TextMatrix(FlexDataGrid.Row, 5)
    Printer.Print Tab(10); LangDurasi; Tab(30); ":"; Tab(35); FlexDataGrid.TextMatrix(FlexDataGrid.Row, 6)
    Printer.Print Tab(10); LangTotal; Tab(30); ":"; Tab(35); mata_uang & ". " & FlexDataGrid.TextMatrix(FlexDataGrid.Row, 7)
    Printer.Print Tab(10); "==============================="
    Printer.Print Tab(10); LangOperator & " : " & FrmMain.StatusBar.Panels(4).Text
    Printer.Print Tab(10); LangGreeting
    Printer.EndDoc
End Sub

Private Sub CmdCetakNota_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdCetakNota, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdHapus_Click()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    cn.Open cnstr
    cmd.ActiveConnection = cn
    cmd.CommandType = adCmdText
    cmd.CommandText = "DELETE FROM Billing WHERE no_nota=" & Val(FlexDataGrid.TextMatrix(FlexDataGrid.Row, 0))
    cmd.Execute
    cn.Close
    Call CekRekap
End Sub

Private Sub CmdHapus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdHapus, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdRefresh_Click()
    Call CekRekap
End Sub

Private Sub CmdRefresh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdRefresh, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdTutup_Click()
    Unload Me
End Sub

Private Sub CmdTutup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdTutup, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub Form_Load()
    jenis_login = FrmMain.StatusBar.Panels(6).Text
    CmdHapus.Caption = LangHapus
    If jenis_login = "admin" Or jenis_login = "sysadmin" Then
        CmdHapus.Visible = True
    Else
        CmdHapus.Visible = False
    End If
    TxtBulan.Clear
    TxtBulan.AddItem LangJanuari
    TxtBulan.AddItem LangFebruari
    TxtBulan.AddItem LangMaret
    TxtBulan.AddItem LangApril
    TxtBulan.AddItem LangMei
    TxtBulan.AddItem LangJuni
    TxtBulan.AddItem LangJuli
    TxtBulan.AddItem LangAgustus
    TxtBulan.AddItem LangSeptember
    TxtBulan.AddItem LangOktober
    TxtBulan.AddItem LangNovember
    TxtBulan.AddItem LangDesember
    Call Init
    CmdRefresh.Caption = LangRefresh
    CmdCetakLaporan.Caption = LangCetakLap
    CmdCetakNota.Caption = LangCetakNota
    CmdTutup.Caption = LangTutup
    FrameJenisRekap.Caption = LangPeriode
    Frame1.Caption = LangJenis
    OpHari.Caption = LangPerHari
    OpBulan.Caption = LangPerBulan
    OpTahun.Caption = LangPerTahun
    OpSemua.Caption = LangSemua
    OpPersonal.Caption = LangJenisPersonal
    OpMember.Caption = LangJenisMember
    OpGame.Caption = LangJenisGame
    OpKetik.Caption = LangJenisKetik
    Label1.Caption = LangTotal
    FrmLaporanRekap.Caption = LangRekapBilling
    Label2.Caption = LangTerbilang
End Sub

Private Sub CekRekap()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    Dim total As Single
    Dim tgl As String
    Dim bln As String
    Dim thn As String
    Dim sql As String
    If TxtBulan.Text = LangJanuari Then
        bln = "01"
    ElseIf TxtBulan.Text = LangFebruari Then
       bln = "02"
    ElseIf TxtBulan.Text = LangMaret Then
        bln = "03"
    ElseIf TxtBulan.Text = LangApril Then
        bln = "04"
    ElseIf TxtBulan.Text = LangMei Then
        bln = "05"
    ElseIf TxtBulan.Text = LangJuni Then
        bln = "06"
    ElseIf TxtBulan.Text = LangJuli Then
        bln = "07"
    ElseIf TxtBulan.Text = LangAgustus Then
        bln = "08"
    ElseIf TxtBulan.Text = LangSeptember Then
        bln = "09"
    ElseIf TxtBulan.Text = LangOktober Then
        bln = "10"
    ElseIf TxtBulan.Text = LangNovember Then
        bln = "11"
    ElseIf TxtBulan.Text = LangDesember Then
        bln = "12"
    End If
    tgl = TxtTgl.Text
    thn = TxtTahun.Text
    total = 0
    If OpHari.value = True Then
        If OpSemua.value = True Then
            sql = "SELECT billing.* FROM billing WHERE billing.tanggal = '" & tgl & "-" & bln & "-" & thn & "'"
        ElseIf OpPersonal.value = True Then
            sql = "SELECT billing.* FROM billing WHERE billing.tanggal = '" & tgl & "-" & bln & "-" & thn & "' AND billing.jenis = '" & LangPersonal & "'"
        ElseIf OpMember.value = True Then
            sql = "SELECT billing.* FROM billing WHERE billing.tanggal = '" & tgl & "-" & bln & "-" & thn & "' AND billing.jenis = '" & LangMember & "'"
        ElseIf OpGame.value = True Then
            sql = "SELECT billing.* FROM billing WHERE billing.tanggal = '" & tgl & "-" & bln & "-" & thn & "' AND billing.jenis = '" & LangGame & "'"
        ElseIf OpKetik.value = True Then
            sql = "SELECT billing.* FROM billing WHERE billing.tanggal = '" & tgl & "-" & bln & "-" & thn & "' AND billing.jenis = '" & LangKetik & "'"
        Else
            sql = "SELECT billing.* FROM billing WHERE billing.tanggal = '" & tgl & "-" & bln & "-" & thn & "'"
        End If
    ElseIf OpBulan.value = True Then
        If OpSemua.value = True Then
            sql = "SELECT billing.* FROM billing WHERE billing.bulan = '" & bln & "' AND billing.tahun = '" & thn & "'"
        ElseIf OpPersonal.value = True Then
            sql = "SELECT billing.* FROM billing WHERE billing.bulan = '" & bln & "' AND billing.tahun = '" & thn & "' AND billing.jenis = '" & LangPersonal & "'"
        ElseIf OpMember.value = True Then
            sql = "SELECT billing.* FROM billing WHERE billing.bulan = '" & bln & "' AND billing.tahun = '" & thn & "' AND billing.jenis = '" & LangMember & "'"
        ElseIf OpGame.value = True Then
            sql = "SELECT billing.* FROM billing WHERE billing.bulan = '" & bln & "' AND billing.tahun = '" & thn & "' AND billing.jenis = '" & LangGame & "'"
        ElseIf OpKetik.value = True Then
            sql = "SELECT billing.* FROM billing WHERE billing.bulan = '" & bln & "' AND billing.tahun = '" & thn & "' AND billing.jenis = '" & LangKetik & "'"
        Else
            sql = "SELECT billing.* FROM billing WHERE billing.bulan = '" & bln & "' AND billing.tahun = '" & thn & "'"
        End If
    ElseIf OpTahun.value = True Then
        If OpSemua.value = True Then
            sql = "SELECT billing.* FROM billing WHERE billing.tahun = '" & thn & "'"
        ElseIf OpPersonal.value = True Then
            sql = "SELECT billing.* FROM billing WHERE billing.tahun = '" & thn & "' AND billing.jenis = '" & LangPersonal & "'"
        ElseIf OpMember.value = True Then
            sql = "SELECT billing.* FROM billing WHERE billing.tahun = '" & thn & "' AND billing.jenis = '" & LangMember & "'"
        ElseIf OpGame.value = True Then
            sql = "SELECT billing.* FROM billing WHERE billing.tahun = '" & thn & "' AND billing.jenis = '" & LangGame & "'"
        ElseIf OpKetik.value = True Then
            sql = "SELECT billing.* FROM billing WHERE billing.tahun = '" & thn & "' AND billing.jenis = '" & LangKetik & "'"
        Else
            sql = "SELECT billing.* FROM billing WHERE billing.tahun = '" & thn & "'"
        End If
    Else
        sql = "SELECT billing.* FROM billing WHERE billing.tanggal = '" & tgl & "-" & bln & "-" & thn & "'"
    End If
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open sql, cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        FlexDataGrid.Rows = rs.RecordCount + 1
        rs.MoveFirst
        For i = 1 To rs.RecordCount
            FlexDataGrid.TextMatrix(i, 0) = rs(0) 'no nota
            FlexDataGrid.TextMatrix(i, 1) = rs(1) ' tanggal
            FlexDataGrid.TextMatrix(i, 2) = LangClient & " " & rs(2) ' no client
            FlexDataGrid.TextMatrix(i, 3) = rs(3) 'nama user
            FlexDataGrid.TextMatrix(i, 4) = rs(4) 'mulai
            FlexDataGrid.TextMatrix(i, 5) = rs(5) 'selesai
            FlexDataGrid.TextMatrix(i, 6) = rs(6) ' durasi
            FlexDataGrid.TextMatrix(i, 7) = Format(rs(7), "##,##0.00") ' total
            FlexDataGrid.TextMatrix(i, 8) = rs(11)
            total = total + rs(7)
            rs.MoveNext
        Next i
        LabelTotal.Caption = Format(total, "##,##0.00")
        LabelTerbilang.Caption = Terbilang(total, lang)
    Else
        FlexDataGrid.Rows = 1
        LabelTotal.Caption = ""
    End If
    cn.Close
End Sub

Private Sub IsiTanggal()
    TxtTgl.Clear
    For i = 1 To JmlHari(CDate(Val(TxtTahun.Text) & "/" & TxtBulan.ListIndex + 1 & "/1"))
        TxtTgl.AddItem i
    Next i
    
    If TxtTgl.ListCount > 0 Then TxtTgl.ListIndex = Month(Now) - 1
End Sub

Private Sub IsiTahun()
    TxtTahun.Clear
    
    For i = 2018 To Year(Now)
        TxtTahun.AddItem i
    Next i
    
    TxtTahun.ListIndex = TxtTahun.ListCount - 1
End Sub

Private Sub OpBulan_Click()
    Call CekRekap
End Sub

Private Sub OpGame_Click()
    Call CekRekap
End Sub

Private Sub OpHari_Click()
    Call CekRekap
End Sub

Private Sub OpKetik_Click()
    Call CekRekap
End Sub

Private Sub OpMember_Click()
    Call CekRekap
End Sub

Private Sub OpPersonal_Click()
    Call CekRekap
End Sub

Private Sub OpSemua_Click()
    Call CekRekap
End Sub

Private Sub OpTahun_Click()
    Call CekRekap
End Sub

Private Sub TxtBulan_Click()
    Call IsiTanggal
    Call CekRekap
End Sub

Private Sub TxtTahun_Click()
    Call IsiTanggal
    Call CekRekap
End Sub

Private Sub TxtTgl_Click()
    Call CekRekap
End Sub
