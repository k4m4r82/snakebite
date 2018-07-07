VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Konfigurasi"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   Icon            =   "FrmConfig.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdTutup 
      Caption         =   "Tutup"
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4680
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7858
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Setingan Umum"
      TabPicture(0)   =   "FrmConfig.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Txtnama_warnet"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Txtalamat1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Txtalamat2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Txtport"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Txtcurrency"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Txtharga_awal"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "CmdUpdateUmum"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Setingan Harga"
      TabPicture(1)   =   "FrmConfig.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Setingan Aplikasi"
      TabPicture(2)   =   "FrmConfig.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label15"
      Tab(2).Control(1)=   "Label18"
      Tab(2).Control(2)=   "Label19"
      Tab(2).Control(3)=   "Label20"
      Tab(2).Control(4)=   "Label21"
      Tab(2).Control(5)=   "Label22"
      Tab(2).Control(6)=   "TxtKetik1"
      Tab(2).Control(7)=   "TxtKetik2"
      Tab(2).Control(8)=   "TxtKetik3"
      Tab(2).Control(9)=   "TxtGame1"
      Tab(2).Control(10)=   "TxtGame2"
      Tab(2).Control(11)=   "TxtGame3"
      Tab(2).Control(12)=   "CmdUpdateSetApp"
      Tab(2).ControlCount=   13
      Begin VB.CommandButton CmdUpdateSetApp 
         Caption         =   "Update"
         Height          =   375
         Left            =   -73440
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox TxtGame3 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -73440
         MaxLength       =   20
         TabIndex        =   43
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox TxtGame2 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -73440
         MaxLength       =   20
         TabIndex        =   42
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox TxtGame1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -73440
         MaxLength       =   20
         TabIndex        =   41
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox TxtKetik3 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -73440
         MaxLength       =   20
         TabIndex        =   40
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox TxtKetik2 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -73440
         MaxLength       =   20
         TabIndex        =   39
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox TxtKetik1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -73440
         MaxLength       =   20
         TabIndex        =   38
         Top             =   720
         Width           =   1695
      End
      Begin VB.Frame Frame2 
         Caption         =   " Discount "
         Height          =   2175
         Left            =   -72000
         TabIndex        =   14
         Top             =   600
         Width           =   2775
         Begin VB.CommandButton CmdUpdateDisc 
            Caption         =   "Update"
            Height          =   375
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   1680
            Width           =   975
         End
         Begin MSMask.MaskEdBox Txtjam_awal 
            Height          =   375
            Left            =   1560
            TabIndex        =   19
            Top             =   720
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   8
            Mask            =   "##:##:##"
            PromptChar      =   "_"
         End
         Begin VB.TextBox Txtdisc 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1560
            TabIndex        =   18
            Top             =   240
            Width           =   375
         End
         Begin MSMask.MaskEdBox Txtjam_akhir 
            Height          =   375
            Left            =   1560
            TabIndex        =   20
            Top             =   1200
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   8
            Mask            =   "##:##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label11 
            Caption         =   "%"
            Height          =   255
            Left            =   2040
            TabIndex        =   28
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label10 
            Caption         =   "Selesai Jam"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "Mulai Jam"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "Discount"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " Kelompok Harga "
         Height          =   3615
         Left            =   -74880
         TabIndex        =   13
         Top             =   600
         Width           =   2775
         Begin VB.TextBox TxtDepositMin 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   36
            Top             =   2640
            Width           =   1095
         End
         Begin VB.TextBox TxtInterval 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1200
            MaxLength       =   2
            TabIndex        =   33
            Top             =   2160
            Width           =   375
         End
         Begin VB.TextBox Txtketik 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1200
            MaxLength       =   5
            TabIndex        =   30
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox Txtgame 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1200
            MaxLength       =   5
            TabIndex        =   29
            Top             =   1200
            Width           =   855
         End
         Begin VB.CommandButton CmdupdateHarga 
            Caption         =   "Update"
            Height          =   375
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   3120
            Width           =   975
         End
         Begin VB.TextBox Txtmember 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1200
            MaxLength       =   5
            TabIndex        =   16
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox Txtpersonal 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1200
            MaxLength       =   5
            TabIndex        =   15
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label14 
            Caption         =   "Min Deposit"
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   2760
            Width           =   855
         End
         Begin VB.Label Label13 
            Caption         =   "Menit"
            Height          =   255
            Left            =   1680
            TabIndex        =   35
            Top             =   2280
            Width           =   495
         End
         Begin VB.Label Label12 
            Caption         =   "Interval"
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   2280
            Width           =   855
         End
         Begin VB.Label Label17 
            Caption         =   "Mengetik"
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label Label16 
            Caption         =   "Game"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "Member"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Personal"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.CommandButton CmdUpdateUmum 
         Caption         =   "Update"
         Height          =   375
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3780
         Width           =   975
      End
      Begin VB.TextBox Txtharga_awal 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2040
         MaxLength       =   5
         TabIndex        =   6
         Top             =   3180
         Width           =   735
      End
      Begin VB.TextBox Txtcurrency 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   5
         Top             =   2700
         Width           =   375
      End
      Begin VB.TextBox Txtport 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2040
         MaxLength       =   5
         TabIndex        =   4
         Top             =   2220
         Width           =   735
      End
      Begin VB.TextBox Txtalamat2 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2040
         MaxLength       =   35
         TabIndex        =   3
         Top             =   1740
         Width           =   3375
      End
      Begin VB.TextBox Txtalamat1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2040
         MaxLength       =   35
         TabIndex        =   2
         Top             =   1260
         Width           =   3375
      End
      Begin VB.TextBox Txtnama_warnet 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2040
         MaxLength       =   35
         TabIndex        =   1
         Top             =   780
         Width           =   3375
      End
      Begin VB.Label Label22 
         Caption         =   "Aplikasi Game"
         Height          =   255
         Left            =   -74760
         TabIndex        =   49
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label21 
         Caption         =   "Aplikasi Game"
         Height          =   255
         Left            =   -74760
         TabIndex        =   48
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label20 
         Caption         =   "Aplikasi Game"
         Height          =   255
         Left            =   -74760
         TabIndex        =   47
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "Aplikasi Ketik"
         Height          =   255
         Left            =   -74760
         TabIndex        =   46
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label18 
         Caption         =   "Aplikasi Ketik"
         Height          =   255
         Left            =   -74760
         TabIndex        =   45
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label15 
         Caption         =   "Aplikasi Ketik"
         Height          =   255
         Left            =   -74760
         TabIndex        =   44
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Harga Awal"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   3300
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Currency"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   2820
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Server Port"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   2340
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Alamat"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   1380
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Nama Warnet"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   900
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdTutup_Click()
    Unload Me
End Sub

Private Sub CekSetingUmum()
    If LihatDbSeting = 1 Then
        Txtnama_warnet.Text = TDbSeting.nama_warnet
        Txtalamat1.Text = TDbSeting.alamat1
        Txtalamat2.Text = TDbSeting.alamat2
        Txtport.Text = TDbSeting.port
        Txtcurrency.Text = TDbSeting.curr
        Txtharga_awal.Text = TDbSeting.harga_awal
    Else
        Txtnama_warnet.Text = TDbSeting.nama_warnet
        Txtalamat1.Text = TDbSeting.alamat1
        Txtalamat2.Text = TDbSeting.alamat2
        Txtport.Text = TDbSeting.port
        Txtcurrency.Text = TDbSeting.curr
        Txtharga_awal.Text = TDbSeting.harga_awal
    End If
End Sub

Private Sub CmdTutup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdTutup, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdUpdateDisc_Click()
    UpdateSetingDiscount
End Sub

Private Sub CmdUpdateDisc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdUpdateDisc, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdupdateHarga_Click()
    UpdateSetingHarga
End Sub

Private Sub CmdupdateHarga_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdupdateHarga, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdUpdateSetApp_Click()
    UpdateSetingApp
End Sub

Private Sub CmdUpdateSetApp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdUpdateSetApp, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdUpdateUmum_Click()
    UpdateSetingUmum
End Sub

Private Sub UpdateSetingUmum()
    If UpdateDbSeting(Txtnama_warnet.Text, Txtalamat1.Text, Txtalamat2.Text, Int(Val(Txtport.Text)), Txtcurrency.Text, Val(Txtharga_awal.Text)) = 1 Then
        FrmMain.Txtnama_warnet.Caption = Txtnama_warnet.Text
        FrmMain.Txtalamat1.Caption = Txtalamat1.Text
        FrmMain.Txtalamat2.Caption = Txtalamat2.Text
        MsgBox LangUpdateSukses, vbInformation, LangInformasi
    End If
End Sub

Private Sub CekSetingHarga()
    If LihatDbHarga = 1 Then
        Txtpersonal.Text = TDbHarga.personal
        Txtmember.Text = TDbHarga.member
        Txtgame = TDbHarga.game
        Txtketik = TDbHarga.ketik
        TxtInterval.Text = TDbHarga.interval
        TxtDepositMin.Text = TDbHarga.depositmin
    Else
        Txtpersonal.Text = TDbHarga.personal
        Txtmember.Text = TDbHarga.member
        Txtgame = TDbHarga.game
        Txtketik = TDbHarga.ketik
        TxtInterval.Text = TDbHarga.interval
        TxtDepositMin.Text = TDbHarga.depositmin
    End If
End Sub

Private Sub UpdateSetingHarga()
    If UpdateDbHarga(Val(Txtpersonal.Text), Val(Txtmember.Text), Val(Txtgame.Text), Val(Txtketik.Text), Int(Val(TxtInterval.Text)), Val(TxtDepositMin.Text)) = 1 Then
        MsgBox LangUpdateSukses, vbInformation, LangInformasi
    End If
End Sub

Private Sub CekSetingDiscount()
    If LihatDbDiskon = 1 Then
        Txtdisc.Text = TDbDiskon.disc
        Txtjam_awal.Text = TDbDiskon.jam_awal
        Txtjam_akhir.Text = TDbDiskon.jam_akhir
    Else
        Txtdisc.Text = TDbDiskon.disc
        Txtjam_awal.Text = TDbDiskon.jam_awal
        Txtjam_akhir.Text = TDbDiskon.jam_akhir
    End If
End Sub

Private Sub UpdateSetingDiscount()
    If UpdateDbDiskon(Int(Val(Txtdisc)), Txtjam_awal.Text, Txtjam_akhir.Text) = 1 Then
        MsgBox LangUpdateSukses, vbInformation, LangInformasi
    End If
End Sub

Private Sub CekSetingApp()
    If AmbilAplikasi = 1 Then
        TxtKetik1.Text = ketik1
        TxtKetik2.Text = ketik2
        TxtKetik3.Text = ketik3
        TxtGame1.Text = game1
        TxtGame2.Text = game2
        TxtGame3.Text = game3
    Else
        TxtKetik1.Text = ketik1
        TxtKetik2.Text = ketik2
        TxtKetik3.Text = ketik3
        TxtGame1.Text = game1
        TxtGame2.Text = game2
        TxtGame3.Text = game3
    End If
End Sub

Private Sub UpdateSetingApp()
    If UpdateDbApp(TxtKetik1.Text, TxtKetik2.Text, TxtKetik3.Text, TxtGame1.Text, TxtGame2.Text, TxtGame3.Text) = 1 Then
        MsgBox LangUpdateSukses, vbInformation, LangInformasi
    End If
End Sub

Private Sub CmdUpdateUmum_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdUpdateUmum, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub Form_Load()
    CekSetingUmum
    CekSetingHarga
    CekSetingDiscount
    CekSetingApp
    FrmConfig.Caption = LangKonfigurasi
    SSTab.TabCaption(0) = LangSetinganUmum
    SSTab.TabCaption(1) = LangSetinganHarga
    SSTab.TabCaption(2) = LangSetinganAplikasi
    Label1.Caption = LangNamaWarnet
    Label2.Caption = LangAlamat
    Label4.Caption = LangServerPort
    Label5.Caption = LangCurrency
    Label6.Caption = LangHargaAwal
    CmdUpdateUmum.Caption = LangUpdate
    CmdTutup.Caption = LangTutup
    CmdupdateHarga.Caption = LangUpdate
    CmdUpdateDisc.Caption = LangUpdate
    Frame1.Caption = " " & LangKelompokHarga & " "
    Frame2.Caption = " " & LangDiscount & " "
    Label3.Caption = LangHargaPersonal
    Label7.Caption = LangHargaMember
    Label16.Caption = LangHargaGame
    Label17.Caption = LangHargaMengetik
    Label12.Caption = LangInterval
    Label14.Caption = LangMinDeposit
    Label8.Caption = LangDiscount
    Label9.Caption = LangMulaiJam
    Label10.Caption = LangSelesaiJam
    CmdUpdateSetApp.Caption = LangUpdate
    Label15.Caption = LangApKetik
    Label18.Caption = LangApKetik
    Label19.Caption = LangApKetik
    Label20.Caption = LangApGame
    Label21.Caption = LangApGame
    Label22.Caption = LangApGame
End Sub

Private Sub Txtalamat1_GotFocus()
    Txtalamat1.Backcolor = &HFFC0C0
End Sub

Private Sub Txtalamat1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txtalamat2.SetFocus
    End If
End Sub

Private Sub Txtalamat1_LostFocus()
    Txtalamat1.Backcolor = &H80000005
End Sub

Private Sub Txtalamat2_GotFocus()
    Txtalamat2.Backcolor = &HFFC0C0
End Sub

Private Sub Txtalamat2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txtport.SetFocus
    End If
End Sub

Private Sub Txtalamat2_LostFocus()
    Txtalamat2.Backcolor = &H80000005
End Sub

Private Sub Txtcurrency_GotFocus()
    Txtcurrency.Backcolor = &HFFC0C0
End Sub

Private Sub Txtcurrency_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txtharga_awal.SetFocus
    End If
End Sub

Private Sub Txtcurrency_LostFocus()
    Txtcurrency.Backcolor = &H80000005
End Sub

Private Sub TxtDepositMin_GotFocus()
    TxtDepositMin.Backcolor = &HFFC0C0
End Sub

Private Sub TxtDepositMin_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 13 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        CmdupdateHarga.SetFocus
    End If
End Sub

Private Sub TxtDepositMin_LostFocus()
    TxtDepositMin.Backcolor = &H80000005
End Sub

Private Sub Txtdisc_GotFocus()
    Txtdisc.Backcolor = &HFFC0C0
End Sub

Private Sub Txtdisc_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 13 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        Txtjam_awal.SetFocus
    End If
End Sub

Private Sub Txtdisc_LostFocus()
    Txtdisc.Backcolor = &H80000005
End Sub

Private Sub Txtgame_GotFocus()
    Txtgame.Backcolor = &HFFC0C0
End Sub

Private Sub Txtgame_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 13 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        Txtketik.SetFocus
    End If
End Sub

Private Sub Txtgame_LostFocus()
    Txtgame.Backcolor = &H80000005
End Sub

Private Sub TxtGame1_GotFocus()
    TxtGame1.Backcolor = &HFFC0C0
End Sub

Private Sub TxtGame1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtGame2.SetFocus
    End If
End Sub

Private Sub TxtGame1_LostFocus()
    TxtGame1.Backcolor = &H80000005
End Sub

Private Sub TxtGame2_GotFocus()
    TxtGame2.Backcolor = &HFFC0C0
End Sub

Private Sub TxtGame2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtGame3.SetFocus
    End If
End Sub

Private Sub TxtGame2_LostFocus()
    TxtGame2.Backcolor = &H80000005
End Sub

Private Sub TxtGame3_GotFocus()
    TxtGame3.Backcolor = &HFFC0C0
End Sub

Private Sub TxtGame3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdUpdateSetApp.SetFocus
    End If
End Sub

Private Sub TxtGame3_LostFocus()
    TxtGame3.Backcolor = &H80000005
End Sub

Private Sub Txtharga_awal_GotFocus()
    Txtharga_awal.Backcolor = &HFFC0C0
End Sub

Private Sub Txtharga_awal_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 13 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        CmdUpdateUmum.SetFocus
    End If
End Sub

Private Sub Txtharga_awal_LostFocus()
    Txtharga_awal.Backcolor = &H80000005
End Sub

Private Sub TxtInterval_GotFocus()
    TxtInterval.Backcolor = &HFFC0C0
End Sub

Private Sub TxtInterval_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 13 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        TxtDepositMin.SetFocus
    End If
End Sub

Private Sub TxtInterval_LostFocus()
    TxtInterval.Backcolor = &H80000005
End Sub

Private Sub Txtjam_akhir_GotFocus()
    Txtjam_akhir.Backcolor = &HFFC0C0
End Sub

Private Sub Txtjam_akhir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdUpdateDisc.SetFocus
    End If
End Sub

Private Sub Txtjam_akhir_LostFocus()
    Txtjam_akhir.Backcolor = &H80000005
End Sub

Private Sub Txtjam_awal_GotFocus()
    Txtjam_awal.Backcolor = &HFFC0C0
End Sub

Private Sub Txtjam_awal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txtjam_akhir.SetFocus
    End If
End Sub

Private Sub Txtjam_awal_LostFocus()
    Txtjam_awal.Backcolor = &H80000005
End Sub

Private Sub Txtketik_GotFocus()
    Txtketik.Backcolor = &HFFC0C0
End Sub

Private Sub Txtketik_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 13 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        TxtInterval.SetFocus
    End If
End Sub

Private Sub Txtketik_LostFocus()
    Txtketik.Backcolor = &H80000005
End Sub

Private Sub TxtKetik1_GotFocus()
    TxtKetik1.Backcolor = &HFFC0C0
End Sub

Private Sub TxtKetik1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtKetik2.SetFocus
    End If
End Sub

Private Sub TxtKetik1_LostFocus()
    TxtKetik1.Backcolor = &H80000005
End Sub

Private Sub TxtKetik2_GotFocus()
    TxtKetik2.Backcolor = &HFFC0C0
End Sub

Private Sub TxtKetik2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtKetik3.SetFocus
    End If
End Sub

Private Sub TxtKetik2_LostFocus()
    TxtKetik2.Backcolor = &H80000005
End Sub

Private Sub TxtKetik3_GotFocus()
    TxtKetik3.Backcolor = &HFFC0C0
End Sub

Private Sub TxtKetik3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtGame1.SetFocus
    End If
End Sub

Private Sub TxtKetik3_LostFocus()
    TxtKetik3.Backcolor = &H80000005
End Sub

Private Sub Txtmember_GotFocus()
    Txtmember.Backcolor = &HFFC0C0
End Sub

Private Sub Txtmember_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 13 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        Txtgame.SetFocus
    End If
End Sub

Private Sub Txtmember_LostFocus()
    Txtmember.Backcolor = &H80000005
End Sub

Private Sub Txtnama_warnet_GotFocus()
    Txtnama_warnet.Backcolor = &HFFC0C0
End Sub

Private Sub Txtnama_warnet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txtalamat1.SetFocus
    End If
End Sub

Private Sub Txtnama_warnet_LostFocus()
    Txtnama_warnet.Backcolor = &H80000005
End Sub

Private Sub Txtpersonal_GotFocus()
    Txtpersonal.Backcolor = &HFFC0C0
End Sub

Private Sub Txtpersonal_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 13 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        Txtmember.SetFocus
    End If
End Sub

Private Sub Txtpersonal_LostFocus()
    Txtpersonal.Backcolor = &H80000005
End Sub

Private Sub Txtport_GotFocus()
    Txtport.Backcolor = &HFFC0C0
End Sub

Private Sub Txtport_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 13 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        Txtcurrency.SetFocus
    End If
End Sub

Private Sub Txtport_LostFocus()
    Txtport.Backcolor = &H80000005
End Sub
