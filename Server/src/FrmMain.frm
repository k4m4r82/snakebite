VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmMain 
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10455
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   10455
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   19
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            Object.Width           =   1e-4
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "mati"
            Object.ToolTipText     =   "Matikan Server"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "stop"
            Object.ToolTipText     =   "Stop Client"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "pindah"
            Object.ToolTipText     =   "Pindahkan Client"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "shutdown"
            Object.ToolTipText     =   "Shutdown Client"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "setpassword"
            Object.ToolTipText     =   "Set Password Client"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "setkonfig"
            Object.ToolTipText     =   "Set Konfigurasi"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "user"
            Object.ToolTipText     =   "Manajemen User"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "laporanbilling"
            Object.ToolTipText     =   "Laporan Billing"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "laporandeposit"
            Object.ToolTipText     =   "Laporan Deposit"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "member"
            Object.ToolTipText     =   "Member"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "pesan"
            Object.ToolTipText     =   "Pesan"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "backup"
            Object.ToolTipText     =   "Backup Database"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "about"
            Object.ToolTipText     =   "About"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Timer TimerPindah 
      Interval        =   1000
      Left            =   7200
      Top             =   480
   End
   Begin VB.Timer TimerStop 
      Interval        =   1000
      Left            =   6720
      Top             =   480
   End
   Begin VB.Timer TimerAktif 
      Interval        =   1000
      Left            =   8160
      Top             =   480
   End
   Begin VB.Timer TimerStart 
      Interval        =   1000
      Left            =   6240
      Top             =   480
   End
   Begin VB.Timer TimerUtama 
      Interval        =   1000
      Left            =   8640
      Top             =   480
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
      Height          =   4200
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   7408
      _Version        =   393216
      Rows            =   17
      Cols            =   11
      ScrollBars      =   2
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   5880
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   9
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   2
            Object.Width           =   1411
            MinWidth        =   1411
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   2
            Object.Width           =   1411
            MinWidth        =   1411
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   2
            Object.Width           =   1411
            MinWidth        =   1411
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   2
            Object.Width           =   3175
            MinWidth        =   3175
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel8 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   2
            Object.Width           =   1588
            MinWidth        =   1588
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel9 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock Winsock 
      Index           =   0
      Left            =   7680
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Txtalamat2 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Left            =   120
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Txtalamat1 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Txtnama_warnet 
      Alignment       =   2  'Center
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
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   2775
   End
   Begin ComctlLib.ImageList ImageList 
      Left            =   9120
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   13
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":04E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":06BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":0898
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":0A72
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":0C4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":0E26
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":1000
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":11DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":13B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":158E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":1768
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":1942
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Height          =   975
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   3135
   End
   Begin VB.Menu mnuserver 
      Caption         =   "&Server"
      Begin VB.Menu mnustatus 
         Caption         =   "Status"
      End
      Begin VB.Menu mnumati 
         Caption         =   "Matikan"
      End
      Begin VB.Menu mnugantiop 
         Caption         =   "Ganti Operator"
      End
      Begin VB.Menu mnuprinter 
         Caption         =   "Print Struk"
      End
   End
   Begin VB.Menu mnuclient 
      Caption         =   "&Client"
      Begin VB.Menu mnustop 
         Caption         =   "Stop Client"
      End
      Begin VB.Menu mnupindah 
         Caption         =   "Pindah Client"
      End
      Begin VB.Menu mnushutdown 
         Caption         =   "Shutdown Client"
      End
      Begin VB.Menu mnukunci 
         Caption         =   "Kunci Client"
      End
      Begin VB.Menu mnubuka 
         Caption         =   "Buka Client"
      End
      Begin VB.Menu mnupassadmin 
         Caption         =   "Set Password"
      End
      Begin VB.Menu mnuremote 
         Caption         =   "Remote Client"
      End
   End
   Begin VB.Menu mnutrans 
      Caption         =   "&Transaksi"
      Begin VB.Menu mnuadeposit 
         Caption         =   "Deposit"
      End
   End
   Begin VB.Menu mnusetting 
      Caption         =   "&Setting"
      Begin VB.Menu mnuconfig 
         Caption         =   "Konfigurasi"
      End
      Begin VB.Menu mnuuser 
         Caption         =   "Manajemen User"
      End
      Begin VB.Menu mnumember 
         Caption         =   "Member"
      End
      Begin VB.Menu mnuregclient 
         Caption         =   "Registrasi Client"
      End
      Begin VB.Menu mnudbpass 
         Caption         =   "Ganti Password Database"
      End
      Begin VB.Menu mnugantibahasa 
         Caption         =   "Ganti Bahasa"
      End
   End
   Begin VB.Menu mnulaporan 
      Caption         =   "&Laporan"
      Begin VB.Menu mnurekap 
         Caption         =   "Rekap Billing"
      End
      Begin VB.Menu mnudeposit 
         Caption         =   "Informasi Deposit"
      End
   End
   Begin VB.Menu mnuutil 
      Caption         =   "&Utility"
      Begin VB.Menu mnupesan 
         Caption         =   "Pesan"
      End
      Begin VB.Menu mnubackup 
         Caption         =   "Backup Database"
      End
      Begin VB.Menu mnukal 
         Caption         =   "Kalkulator"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iSockets As Integer
Dim iSocketsFree(SockMax) As Boolean
Dim port As Long
Dim durasi(1 To ClientMax) As Date
Dim durasi_disc(1 To ClientMax) As Date
Dim biaya(1 To ClientMax) As Single
Dim biaya_disc(1 To ClientMax) As Single
Dim total(1 To ClientMax) As Single
Dim mulai(1 To ClientMax) As Date
Dim disc_mulai(1 To ClientMax) As Date
Dim interval As Long
Dim atas(1 To ClientMax) As Long
Dim bawah(1 To ClientMax) As Long
Dim mulai_disc As Date
Dim akhir_disc As Date
Dim nama_warnet As String
Dim alamat1 As String
Dim alamat2 As String

Private Sub Form_Load()
    FrmMain.Caption = NamaAplikasi & VersiAplikasi
    Call LangMenu
    Call LangToolBar
    Call HapusdbSocket
    Call CekSetingUmum
    Call CekLogin
    Call AturGrid
    Call CekSetingHarga
    Call CekSetingDiscount
    Call CekPasswordAdmin
    Call CekTemp
    Call FreeSock
    Call AmbilAplikasi
    Call BukaKunci
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Load FrmLogOut
    FrmLogOut.Show
End Sub

Private Sub mnuabout_Click()
    Load FrmAbout
    FrmAbout.Show 1
End Sub

Private Sub mnuadeposit_Click()
    Load FrmDeposit
    FrmDeposit.Show 1
End Sub

Private Sub mnubackup_Click()
    Load FrmBackup
    FrmBackup.Show 1
End Sub

Private Sub mnubuka_Click()
    Load FrmUnLock
    FrmUnLock.TxtNo_Client.Text = clientpopup
    FrmUnLock.Show 1
End Sub

Private Sub mnuconfig_Click()
    Load FrmConfig
    FrmConfig.Show 1
End Sub

Private Sub mnudbpass_Click()
    Load FrmDbPass
    FrmDbPass.Show 1
End Sub

Private Sub mnudeposit_Click()
    Load FrmLaporanDeposit
    FrmLaporanDeposit.Show 1
End Sub

Private Sub mnugantibahasa_Click()
    Load FrmBahasa
    FrmBahasa.Show 1
End Sub

Private Sub mnugantiop_Click()
    Load FrmGantiOP
    FrmGantiOP.Show 1
End Sub

Private Sub mnuhapusclient_Click()
    Call HapusdbIP
End Sub

Private Sub mnukal_Click()
    Load FrmKal
    FrmKal.Show 1
End Sub

Private Sub mnukunci_Click()
    Load FrmLock
    FrmLock.TxtNo_Client.Text = clientpopup
    FrmLock.Show 1
End Sub

Private Sub mnumati_Click()
    Unload Me
End Sub

Private Sub mnumember_Click()
    Load FrmMember
    FrmMember.Show 1
End Sub

Private Sub mnupassadmin_Click()
    Load FrmPasswordAdmin
    FrmPasswordAdmin.Show 1
End Sub

Private Sub mnupesan_Click()
    Load FrmPesan
    FrmPesan.Show 1
End Sub

Private Sub mnupindah_Click()
    Load FrmPindahClient
    FrmPindahClient.TxtNoDari.Text = clientpopup
    FrmPindahClient.Show 1
End Sub

Private Sub mnuprinter_Click()
    If mnuprinter.Checked = True Then
        mnuprinter.Checked = False
    Else
        mnuprinter.Checked = True
    End If
End Sub

Private Sub mnuregclient_Click()
    Load FrmRegister
    FrmRegister.Show 1
End Sub

Private Sub mnurekap_Click()
    Load FrmLaporanRekap
    FrmLaporanRekap.Show 1
End Sub

Private Sub mnuremote_Click()
    Load FrmRemote
    FrmRemote.TxtNo_Client.Text = clientpopup
    FrmRemote.Show 1
End Sub

Private Sub mnushutdown_Click()
    Load FrmShutdownClient
    FrmShutdownClient.TxtNo_Client.Text = clientpopup
    FrmShutdownClient.Show 1
End Sub

Private Sub mnustatus_Click()
    Load FrmStatus
    FrmStatus.Show 1
End Sub

Private Sub mnustop_Click()
    Load FrmStopClient
    FrmStopClient.TxtNo_Client.Text = clientpopup
    FrmStopClient.Show 1
End Sub

Private Sub mnuuser_Click()
    Load FrmUser
    FrmUser.Show 1
End Sub

Private Sub MSFlexGrid_Click()
    clientpopup = MSFlexGrid.Row
End Sub

Private Sub MSFlexGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
      If clientpopup = 0 Then
        clientpopup = 1
      End If
      PopupMenu mnuclient
   End If
End Sub

Private Sub TimerAktif_Timer()
    If (Now > mulai_disc And Now < akhir_disc) Then
        For i = 1 To ClientMax
            If MSFlexGrid.TextMatrix(i, 3) = LangAktif Then
                durasi_disc(i) = Now - disc_mulai(i)
                durasi(i) = Now - mulai(i)
                atas(i) = Int(Val(Format(durasi(i), "nn")) / interval)
                bawah(i) = atas(i) + 1
                If MSFlexGrid.TextMatrix(i, 8) = LangPersonal Then
                    If Val(Format(durasi(i), "nn")) / interval >= atas(i) And Val(Format(durasi(i), "nn")) / interval <= bawah(i) Then
                        biaya(i) = Round((Val(Mid(Format(durasi(i), "hh:nn:ss"), 1, 2)) * harga_personal * 60 / interval) + (bawah(i) * harga_personal), 2)
                        biaya_disc(i) = Round(((Val(Mid(Format(durasi_disc(i), "hh:nn:ss"), 1, 2)) * harga_personal * 60 / interval) + (bawah(i) * harga_personal)) * discount / 100, 2)
                    End If
                    If biaya(i) <= harga_awal Then
                        biaya(i) = harga_awal
                    End If
                ElseIf MSFlexGrid.TextMatrix(i, 8) = LangMember Then
                    If Val(Format(durasi(i), "nn")) / interval >= atas(i) And Val(Format(durasi(i), "nn")) / interval <= bawah(i) Then
                        biaya(i) = Round((Val(Mid(Format(durasi(i), "hh:nn:ss"), 1, 2)) * harga_member * 60 / interval) + (bawah(i) * harga_member), 2)
                        biaya_disc(i) = Round(((Val(Mid(Format(durasi_disc(i), "hh:nn:ss"), 1, 2)) * harga_member * 60 / interval) + (bawah(i) * harga_member)) * discount / 100, 2)
                    End If
                    If biaya(i) <= harga_awal Then
                        biaya(i) = harga_awal
                    End If
                ElseIf MSFlexGrid.TextMatrix(i, 8) = "admin" Then
                    biaya(i) = 0
                    biaya_disc(i) = 0
                ElseIf MSFlexGrid.TextMatrix(i, 8) = LangGame Then
                    If Val(Format(durasi(i), "nn")) / interval >= atas(i) And Val(Format(durasi(i), "nn")) / interval <= bawah(i) Then
                        biaya(i) = Round((Val(Mid(Format(durasi(i), "hh:nn:ss"), 1, 2)) * harga_game * 60 / interval) + (bawah(i) * harga_game), 2)
                        biaya_disc(i) = Round(((Val(Mid(Format(durasi_disc(i), "hh:nn:ss"), 1, 2)) * harga_game * 60 / interval) + (bawah(i) * harga_game)) * discount / 100, 2)
                    End If
                    If biaya(i) <= harga_awal Then
                        biaya(i) = harga_awal
                    End If
                ElseIf MSFlexGrid.TextMatrix(i, 8) = LangKetik Then
                    If Val(Format(durasi(i), "nn")) / interval >= atas(i) And Val(Format(durasi(i), "nn")) / interval <= bawah(i) Then
                        biaya(i) = Round((Val(Mid(Format(durasi(i), "hh:nn:ss"), 1, 2)) * harga_ketik * 60 / interval) + (bawah(i) * harga_ketik), 2)
                        biaya_disc(i) = Round(((Val(Mid(Format(durasi_disc(i), "hh:nn:ss"), 1, 2)) * harga_ketik * 60 / interval) + (bawah(i) * harga_ketik)) * discount / 100, 2)
                    End If
                    If biaya(i) <= harga_awal Then
                        biaya(i) = harga_awal
                    End If
                End If
                total(i) = biaya(i) - biaya_disc(i)
                MSFlexGrid.TextMatrix(i, 6) = Format(durasi(i), "hh:nn:ss")
                MSFlexGrid.TextMatrix(i, 7) = Format(biaya(i), "##,##0.00")
                MSFlexGrid.TextMatrix(i, 9) = Format(biaya_disc(i), "##,##0.00")
                MSFlexGrid.TextMatrix(i, 10) = Format(total(i), "##,##0.00")
            End If
        Next i
    Else
        For i = 1 To ClientMax
            If MSFlexGrid.TextMatrix(i, 3) = LangAktif Then
                durasi(i) = Now - mulai(i)
                atas(i) = Int(Val(Format(durasi(i), "nn")) / interval)
                bawah(i) = atas(i) + 1
                If MSFlexGrid.TextMatrix(i, 8) = LangPersonal Then
                    If Val(Format(durasi(i), "nn")) / interval >= atas(i) And Val(Format(durasi(i), "nn")) / interval <= bawah(i) Then
                        biaya(i) = Round((Val(Mid(Format(durasi(i), "hh:nn:ss"), 1, 2)) * harga_personal * 60 / interval) + (bawah(i) * harga_personal), 2)
                        biaya_disc(i) = Round(((Val(Mid(Format(durasi_disc(i), "hh:nn:ss"), 1, 2)) * harga_personal * 60 / interval) + (bawah(i) * harga_personal)) * discount / 100, 2)
                    End If
                    If biaya(i) <= harga_awal Then
                        biaya(i) = harga_awal
                    End If
                ElseIf MSFlexGrid.TextMatrix(i, 8) = LangMember Then
                    If Val(Format(durasi(i), "nn")) / interval >= atas(i) And Val(Format(durasi(i), "nn")) / interval <= bawah(i) Then
                        biaya(i) = Round((Val(Mid(Format(durasi(i), "hh:nn:ss"), 1, 2)) * harga_member * 60 / interval) + (bawah(i) * harga_member), 2)
                        biaya_disc(i) = Round(((Val(Mid(Format(durasi_disc(i), "hh:nn:ss"), 1, 2)) * harga_member * 60 / interval) + (bawah(i) * harga_member)) * discount / 100, 2)
                    End If
                    If biaya(i) <= harga_awal Then
                        biaya(i) = harga_awal
                    End If
                ElseIf MSFlexGrid.TextMatrix(i, 8) = "admin" Then
                    biaya(i) = 0
                    biaya_disc(i) = 0
                ElseIf MSFlexGrid.TextMatrix(i, 8) = LangGame Then
                    If Val(Format(durasi(i), "nn")) / interval >= atas(i) And Val(Format(durasi(i), "nn")) / interval <= bawah(i) Then
                        biaya(i) = Round((Val(Mid(Format(durasi(i), "hh:nn:ss"), 1, 2)) * harga_game * 60 / interval) + (bawah(i) * harga_game), 2)
                        biaya_disc(i) = Round(((Val(Mid(Format(durasi_disc(i), "hh:nn:ss"), 1, 2)) * harga_game * 60 / interval) + (bawah(i) * harga_game)) * discount / 100, 2)
                    End If
                    If biaya(i) <= harga_awal Then
                        biaya(i) = harga_awal
                    End If
                ElseIf MSFlexGrid.TextMatrix(i, 8) = LangKetik Then
                    If Val(Format(durasi(i), "nn")) / interval >= atas(i) And Val(Format(durasi(i), "nn")) / interval <= bawah(i) Then
                        biaya(i) = Round((Val(Mid(Format(durasi(i), "hh:nn:ss"), 1, 2)) * harga_ketik * 60 / interval) + (bawah(i) * harga_ketik), 2)
                        biaya_disc(i) = Round(((Val(Mid(Format(durasi_disc(i), "hh:nn:ss"), 1, 2)) * bawah(i) * harga_ketik * 60 / interval) + (bawah(i) * harga_ketik)) * discount / 100, 2)
                    End If
                    If biaya(i) <= harga_awal Then
                        biaya(i) = harga_awal
                    End If
                End If
                If Format(durasi_disc(i), "hh:nn:ss") = "00:00:00" Then
                    biaya_disc(i) = 0
                End If
                total(i) = biaya(i) - biaya_disc(i)
                MSFlexGrid.TextMatrix(i, 6) = Format(durasi(i), "hh:nn:ss")
                MSFlexGrid.TextMatrix(i, 7) = Format(biaya(i), "##,##0.00")
                MSFlexGrid.TextMatrix(i, 9) = Format(biaya_disc(i), "##,##0.00")
                MSFlexGrid.TextMatrix(i, 10) = Format(total(i), "##,##0.00")
            End If
        Next i
    End If
End Sub

Private Sub TimerPindah_Timer()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    For i = 1 To ClientMax
        If MSFlexGrid.TextMatrix(i, 3) = LangPindah Then
            cn.Open cnstr
            rs.ActiveConnection = cn
            rs.Open "SELECT proses.* FROM proses WHERE proses.no_client = " & i, cn, adOpenStatic, adLockOptimistic
            If Not rs.RecordCount = 0 Then
                mulai(i) = rs(4)
                If mulai(i) >= mulai_disc Then
                    disc_mulai(i) = mulai(i)
                Else
                    disc_mulai(i) = mulai_disc
                End If
            End If
            cn.Close
            MSFlexGrid.TextMatrix(i, 4) = Format(mulai(i), "hh:nn:ss")
            MSFlexGrid.TextMatrix(i, 3) = LangAktif
        End If
    Next i
End Sub

Private Sub TimerStart_Timer()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    For i = 1 To ClientMax
        If MSFlexGrid.TextMatrix(i, 3) = "start" Then
            MSFlexGrid.TextMatrix(i, 4) = Format(Now, "hh:nn:ss")
            durasi(i) = CDate("00:00:00")
            durasi_disc(i) = CDate("00:00:00")
            mulai(i) = Now
            If Now >= mulai_disc Then
                disc_mulai(i) = Now
            Else
                disc_mulai(i) = mulai_disc
            End If
            'tulis proses dalam db
            cn.Open cnstr
            rs.ActiveConnection = cn
            rs.Open "SELECT proses.* FROM proses WHERE proses.no_client = " & i, cn, adOpenStatic, adLockOptimistic
            If Not rs.RecordCount = 0 Then
                rs(0) = i 'no client
                rs(3) = MSFlexGrid.TextMatrix(i, 2) 'user
                rs(1) = LangAktif 'status
                If MSFlexGrid.TextMatrix(i, 8) = LangPersonal Then
                    rs(2) = LangPersonal
                ElseIf MSFlexGrid.TextMatrix(i, 8) = LangMember Then
                    rs(2) = LangMember
                ElseIf MSFlexGrid.TextMatrix(i, 8) = LangGame Then
                    rs(2) = LangGame
                ElseIf MSFlexGrid.TextMatrix(i, 8) = LangKetik Then
                    rs(2) = LangKetik
                End If
                rs(4) = Now
                rs.Update
            Else
                rs.AddNew
                rs(0) = i 'no client
                rs(3) = MSFlexGrid.TextMatrix(i, 2) 'user
                rs(1) = LangAktif 'status
                If MSFlexGrid.TextMatrix(i, 8) = LangPersonal Then
                    rs(2) = LangPersonal
                ElseIf MSFlexGrid.TextMatrix(i, 8) = LangMember Then
                    rs(2) = LangPersonal
                ElseIf MSFlexGrid.TextMatrix(i, 8) = LangGame Then
                    rs(2) = LangGame
                ElseIf MSFlexGrid.TextMatrix(i, 8) = LangKetik Then
                    rs(2) = LangKetik
                End If
                rs(4) = Now
                rs.Update
            End If
            cn.Close
            MSFlexGrid.TextMatrix(i, 3) = LangAktif
            MSFlexGrid.TextMatrix(i, 5) = ""
        End If
    Next i
End Sub

Private Sub TimerStop_Timer()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    For i = 1 To ClientMax
        If MSFlexGrid.TextMatrix(i, 3) = "stop" Then
            'cetak struk
            If mnuprinter.Checked = True Then
                Printer.PrintQuality = 300
                Printer.FontSize = 10
                Printer.ScaleLeft = 10
                Printer.ScaleTop = 10
                Printer.Print Tab(10); Txtnama_warnet
                Printer.Print Tab(10); Txtalamat1
                Printer.Print Tab(10); Txtalamat2
                Printer.Print Tab(10); "==============================="
                Printer.Print Tab(10); LangTanggal; Tab(30); ":"; Tab(35); Format(Now, "dd-mm-yyyy")
                Printer.Print Tab(10); LangNamaUser; Tab(30); ":"; Tab(35); MSFlexGrid.TextMatrix(i, 2)
                Printer.Print Tab(10); LangClient; Tab(30); ":"; Tab(35); "Client " & i
                Printer.Print Tab(10); LangMulai; Tab(30); ":"; Tab(35); Format(mulai(i), "hh:nn:ss")
                Printer.Print Tab(10); LangSelesai; Tab(30); ":"; Tab(35); Format(Now, "hh:nn:ss")
                Printer.Print Tab(10); LangDurasi; Tab(30); ":"; Tab(35); Format(durasi(i), "hh:nn:ss")
                Printer.Print Tab(10); LangTotal; Tab(30); ":"; Tab(35); mata_uang & ". " & Format(total(i), "##,##0.00")
                Printer.Print Tab(10); "==============================="
                Printer.Print Tab(10); "Operator : " & StatusBar.Panels(4).Text
                Printer.Print Tab(10); LangGreeting
                Printer.EndDoc
            End If
            'tulis billing dalam db
            cn.Open cnstr
            rs.ActiveConnection = cn
            rs.Open "SELECT billing.* FROM billing", cn, adOpenStatic, adLockOptimistic
            rs.AddNew
            rs(1) = Format(Date, "dd-mm-yyyy") 'tanggal
            rs(2) = i 'no client
            rs(3) = MSFlexGrid.TextMatrix(i, 2) 'user
            rs(4) = Format(mulai(i), "hh:nn:ss") ' mulai
            rs(5) = Format(Now, "hh:nn:ss") ' selesai
            rs(6) = Format(durasi(i), "hh:nn:ss") ' durasi
            rs(7) = total(i) 'total
            rs(8) = Format(Date, "mm") 'bulan
            rs(9) = Format(Date, "yyyy") 'tahun
            MSFlexGrid.TextMatrix(i, 5) = Format(Now, "hh:nn:ss")
            If MSFlexGrid.TextMatrix(i, 8) = LangPersonal Then 'jenis
                rs(10) = LangPersonal
            ElseIf MSFlexGrid.TextMatrix(i, 8) = LangMember Then
                rs(10) = LangMember
            ElseIf MSFlexGrid.TextMatrix(i, 8) = LangGame Then
                rs(10) = LangGame
            ElseIf MSFlexGrid.TextMatrix(i, 8) = LangKetik Then
                rs(10) = LangKetik
            End If
            rs(11) = StatusBar.Panels(4).Text
            rs.Update
            cn.Close
            'tulis proses dalam db
            cn.Open cnstr
            rs.ActiveConnection = cn
            rs.Open "SELECT proses.* FROM proses WHERE proses.no_client = " & i, cn, adOpenStatic, adLockOptimistic
            If Not rs.RecordCount = 0 Then
                rs(0) = i 'no client
                rs(3) = MSFlexGrid.TextMatrix(i, 2) 'user
                rs(1) = LangBebas
                rs(4) = Now
                If MSFlexGrid.TextMatrix(i, 8) = LangPersonal Then 'jenis
                    rs(2) = LangPersonal
                ElseIf MSFlexGrid.TextMatrix(i, 8) = LangMember Then
                    rs(2) = LangMember
                ElseIf MSFlexGrid.TextMatrix(i, 8) = LangGame Then
                    rs(2) = LangGame
                ElseIf MSFlexGrid.TextMatrix(i, 8) = LangKetik Then
                    rs(2) = LangKetik
                End If
                rs.Update
            Else
                rs.AddNew
                rs(0) = i 'no client
                rs(3) = MSFlexGrid.TextMatrix(i, 2) 'user
                rs(1) = LangBebas
                rs(4) = Now
                If MSFlexGrid.TextMatrix(i, 8) = LangPersonal Then 'jenis
                    rs(2) = LangPersonal
                ElseIf MSFlexGrid.TextMatrix(i, 8) = LangMember Then
                    rs(2) = LangMember
                ElseIf MSFlexGrid.TextMatrix(i, 8) = LangGame Then
                    rs(2) = LangGame
                ElseIf MSFlexGrid.TextMatrix(i, 8) = LangKetik Then
                    rs(2) = LangKetik
                End If
                rs.Update
            End If
            cn.Close
            If MSFlexGrid.TextMatrix(i, 8) = LangMember Then
                'update deposit member
                cn.Open cnstr
                rs.ActiveConnection = cn
                rs.Open "SELECT member.* FROM member WHERE member.user = '" & MSFlexGrid.TextMatrix(i, 2) & "'", cn, adOpenStatic, adLockOptimistic
                If Not rs.RecordCount = 0 Then
                    rs(5) = rs(5) - total(i)
                    rs.Update
                End If
                cn.Close
            End If
            MSFlexGrid.TextMatrix(i, 3) = LangBebas
        End If
    Next i
End Sub

Private Sub TimerUtama_Timer()
    StatusBar.Panels(7) = Format(Date, "dddd, dd mmmm yyyy")
    StatusBar.Panels(8) = Format(Time, "hh:nn:ss")
End Sub

Private Sub CekTemp()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    mulai_disc = CDate(Format(Date, "dd/mm/yyyy") & " " & jam_awal)
    If CDate(jam_akhir) < CDate(jam_awal) Then
        akhir_disc = CDate(Format(Date + 1, "dd/mm/yyyy") & " " & jam_akhir)
    Else
        akhir_disc = CDate(Format(Date, "dd/mm/yyyy") & " " & jam_akhir)
    End If
    For i = 1 To ClientMax
        cn.Open cnstr
        rs.ActiveConnection = cn
        rs.Open "SELECT proses.* FROM proses WHERE proses.no_client = " & i & " AND proses.status = '" & LangAktif & "'", cn, adOpenStatic, adLockOptimistic
        If Not rs.RecordCount = 0 Then
            mulai(i) = CDate(rs(4))
            If mulai(i) >= mulai_disc Then
                disc_mulai(i) = mulai(i)
            Else
                disc_mulai(i) = mulai_disc
            End If
            MSFlexGrid.TextMatrix(i, 1) = LangClient & rs(0) ' no client
            MSFlexGrid.TextMatrix(i, 2) = rs(3) 'user
            If rs(1) = LangAktif Then
                MSFlexGrid.TextMatrix(i, 3) = LangAktif 'status
            ElseIf rs(1) = LangBebas Then
                MSFlexGrid.TextMatrix(i, 3) = LangBebas 'status
            End If
            MSFlexGrid.Col = 3
            MSFlexGrid.Row = i
            MSFlexGrid.CellBackColor = &HC0FFC0       'hijau
            MSFlexGrid.TextMatrix(i, 4) = Format(CDate(rs(4)), "hh:nn:ss") 'mulai
            If rs(2) = LangPersonal Then
                MSFlexGrid.TextMatrix(i, 8) = LangPersonal 'jenis
            ElseIf rs(2) = LangMember Then
                MSFlexGrid.TextMatrix(i, 8) = LangMember 'jenis
            ElseIf rs(2) = LangGame Then
                MSFlexGrid.TextMatrix(i, 8) = LangGame 'jenis
            ElseIf rs(2) = LangKetik Then
                MSFlexGrid.TextMatrix(i, 8) = LangKetik 'jenis
            End If
        Else
            MSFlexGrid.TextMatrix(i, 1) = LangClient & i 'no client
            MSFlexGrid.TextMatrix(i, 2) = "" ' user
            MSFlexGrid.TextMatrix(i, 3) = LangBebas 'status
            MSFlexGrid.Col = 3
            MSFlexGrid.Row = i
            MSFlexGrid.CellBackColor = &HFFC0C0         'ungu
            MSFlexGrid.TextMatrix(i, 4) = "" 'mulai
            MSFlexGrid.TextMatrix(i, 5) = "" 'mulai
            MSFlexGrid.TextMatrix(i, 6) = "" 'durasi
            MSFlexGrid.TextMatrix(i, 7) = "" 'biaya
            MSFlexGrid.TextMatrix(i, 8) = "" 'jenis
            MSFlexGrid.TextMatrix(i, 9) = "" 'discount
            MSFlexGrid.TextMatrix(i, 10) = "" 'total
        End If
        cn.Close
    Next i
End Sub

Private Sub CekSetingUmum()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open "SELECT seting.* FROM seting", cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        rs.MoveFirst
        Txtnama_warnet.Caption = rs(0)
        nama_warnet = rs(0)
        Txtalamat1.Caption = rs(1)
        alamat1 = rs(1)
        Txtalamat2.Caption = rs(2)
        alamat2 = rs(2)
        mata_uang = rs(4)
        harga_awal = rs(5)
        port = rs(3)
        Winsock(0).LocalPort = port
        Winsock(0).Listen
        iSockets = 1
        iSocketsFree(iSockets) = True
        cn.Close
    Else
        cn.Close
        Txtnama_warnet.Caption = "Warnet Citra Mandiri"
        Txtalamat1.Caption = "Jl. Dewi Sartika No. 4"
        Txtalamat2.Caption = "Abepura - Jayapura"
        nama_warnet = Txtnama_warnet.Caption
        alamat1 = Txtalamat1.Caption
        alamat2 = Txtalamat2.Caption
        harga_awal = 1000
        mata_uang = "Rp"
        port = 1500
        Winsock(0).LocalPort = port
        Winsock(0).Listen
        iSockets = 1
        iSocketsFree(iSockets) = True
    End If
End Sub

Private Sub CekSetingHarga()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open "SELECT harga.* FROM harga", cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        rs.MoveFirst
        harga_personal = rs(1)
        harga_member = rs(2)
        harga_game = rs(3)
        harga_ketik = rs(4)
        interval = rs(5)
        deposit_min = rs(6)
    Else
        harga_personal = 2500
        harga_member = 2000
        harga_game = 1500
        harga_ketik = 1000
        interval = 15
        deposit_min = 25000
    End If
    cn.Close
End Sub

Private Sub CekSetingDiscount()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open "SELECT discount.* FROM discount", cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        rs.MoveFirst
        discount = rs(0)
        jam_awal = rs(1)
        jam_akhir = rs(2)
    Else
        discount = 30
        jam_awal = "22:00:00"
        jam_akhir = "04:00:00"
    End If
    cn.Close
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
    Case Is = "mati"
        Unload Me
    Case Is = "stop"
        Load FrmStopClient
        FrmStopClient.TxtNo_Client.Text = clientpopup
        FrmStopClient.Show 1
    Case Is = "pindah"
        Load FrmPindahClient
        FrmPindahClient.TxtNoDari.Text = clientpopup
        FrmPindahClient.Show 1
    Case Is = "shutdown"
        Load FrmShutdownClient
        FrmShutdownClient.TxtNo_Client.Text = clientpopup
        FrmShutdownClient.Show 1
    Case Is = "setpassword"
        Load FrmPasswordAdmin
        FrmPasswordAdmin.Show 1
    Case Is = "setkonfig"
        Load FrmConfig
        FrmConfig.Show 1
    Case Is = "user"
        Load FrmUser
        FrmUser.Show 1
    Case Is = "laporanbilling"
            Load FrmLaporanRekap
            FrmLaporanRekap.Show 1
    Case Is = "laporandeposit"
            Load FrmLaporanDeposit
            FrmLaporanDeposit.Show 1
    Case Is = "member"
            Load FrmMember
            FrmMember.Show 1
    Case Is = "pesan"
            Load FrmPesan
            FrmPesan.Show 1
    Case Is = "backup"
            Load FrmBackup
            FrmBackup.Show 1
    Case Is = "about"
            Load FrmAbout
            FrmAbout.Show 1
    End Select
End Sub

Private Sub Winsock_Close(Index As Integer)
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    Winsock(Index).Close
    Unload Winsock(Index)
    'tulis socket dalam db
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open "SELECT socket.* FROM socket WHERE socket.socket = " & Index, cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
       rs.Delete
    End If
    cn.Close
    iSocketsFree(Index) = True
    'cek socket bebas dalam db
    For i = 1 To SockMax
        cn.Open cnstr
        rs.ActiveConnection = cn
        rs.Open "SELECT socket.* FROM socket WHERE socket.socket = " & i, cn, adOpenStatic, adLockOptimistic
        If rs.RecordCount = 0 Then
            iSockets = i
            iSocketsFree(iSockets) = True
            cn.Close
            Exit For
        Else
            cn.Close
        End If
    Next i
End Sub

Private Sub Winsock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim ip_client As String
        If iSocketsFree(iSockets) = True Then
            Load Winsock(iSockets)
            Winsock(iSockets).LocalPort = port
            Winsock(iSockets).Accept requestID
            ip_client = Winsock(iSockets).RemoteHostIP
            'tulis socket dalam db
            cn.Open cnstr
            rs.ActiveConnection = cn
            rs.Open "SELECT socket.* FROM socket WHERE socket.ip = '" & ip_client & "'", cn, adOpenStatic, adLockOptimistic
            If Not rs.RecordCount = 0 Then
                rs(0) = ip_client
                rs(1) = iSockets 'socket
                rs.Update
            Else
                rs.AddNew
                rs(0) = ip_client
                rs(1) = iSockets 'socket
                rs.Update
            End If
            cn.Close
            iSocketsFree(iSockets) = False
            sockload = True
            'update iSockets
            For i = 1 To SockMax
                cn.Open cnstr
                rs.ActiveConnection = cn
                rs.Open "SELECT socket.* FROM socket WHERE socket.socket = " & i, cn, adOpenStatic, adLockOptimistic
                If rs.RecordCount = 0 Then
                    iSockets = i
                    iSocketsFree(iSockets) = True
                    cn.Close
                    Exit For
                Else
                    cn.Close
                End If
            Next i
        Else
            For i = 1 To SockMax
                cn.Open cnstr
                rs.ActiveConnection = cn
                rs.Open "SELECT socket.* FROM socket WHERE socket.socket = " & i, cn, adOpenStatic, adLockOptimistic
                If rs.RecordCount = 0 Then
                    iSockets = i
                    iSocketsFree(iSockets) = True
                    cn.Close
                    Exit For
                Else
                    cn.Close
                End If
            Next i
        End If
End Sub

Private Sub Winsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    '**********************************************************************************************************************
    '* available command to server                                                                                        *
    '* syntax = perintah|ip|op1|op2|op3|                                                                                  *
    '* register client = register|ip|no_client|_|_|                                                                       *
    '* connect = connect|ip|_|_|_|                                                                                        *
    '* set = set|ip|_|_|_|                                                                                                *
    '* status = status|ip|_|_|_|                                                                                          *
    '* start = start|ip|user|jenis|                                                                                       *
    '* stop = stop|ip|_|_|_|                                                                                              *
    '* password = password|ip|user|_|_|                                                                                   *
    '* pesan = pesan|ip|pesannya|_|_|                                                                                     *
    '*====================================================================================================================*
    '* available command to client                                                                                        *
    '* time = time|time|date|no_client|interval|nama_warnet|alamat1|alamat2|lang|                                         *
    '* set = set|discount|jam_awal|jam_akhir|harga_personal|harga_member|harga_awal|harga_game|harga_ketik|password_admin *
    '* pindahke = pindahke|_|_|_|_|_|_|                                                                                   *
    '* pindahdari = pindahdari|mulai|jenis|_|_|_|_|                                                                       *
    '* aktif = aktif|mulai|jenis|_|_|_|_|                                                                                 *
    '* stop = stop|_|_|_|_|_|_|                                                                                           *
    '* shutdown = shutdown|_|_|_|_|_|_|                                                                                   *
    '* password = password|password|member_ok|_|_|_|_|                                                                    *
    '* pesan = pesan|pesannya|_|_|_|_|_|                                                                                  *
    '* aplikasi = app|ketik1|ketik2|ketik3|game1|game2|game3|                                                             *
    '* snakebites@users.sourceforge.net                                                                                   *
    '**********************************************************************************************************************
    Dim str As String
    Dim sub_str() As String
    Dim perintah As String
    Dim ip As String
    Dim op1 As String
    Dim op2 As String
    Dim op3 As String
    Dim no_client As Integer
    Winsock(Index).GetData str
    sub_str() = Split(str, "|")
    If UBound(sub_str) >= 0 Then
        perintah = sub_str(0)
    End If
    If UBound(sub_str) >= 1 Then
        ip = sub_str(1)
    End If
    If UBound(sub_str) >= 2 Then
        op1 = sub_str(2)
    End If
    If UBound(sub_str) >= 3 Then
        op2 = sub_str(3)
    End If
    If UBound(sub_str) >= 4 Then
        op3 = sub_str(4)
    End If
    
    If perintah = "register" Then
        cn.Open cnstr
        rs.ActiveConnection = cn
        rs.Open "SELECT ip.* FROM ip WHERE ip.ip = '" & ip & "' OR ip.no_client = " & Val(op1), cn, adOpenStatic, adLockOptimistic
        If Not rs.RecordCount = 0 Then
            Winsock(Index).SendData LangNoClientTerpakai & "|_|_|_|_|_|_|"
        Else
            Winsock(Index).SendData "OK|_|_|_|_|_|_|"
            rs.AddNew
            rs(0) = Val(op1)
            rs(1) = ip
            rs.Update
        End If
        cn.Close
    ElseIf perintah = "status" Then
        cn.Open cnstr
        rs.ActiveConnection = cn
        rs.Open "SELECT ip.* FROM ip WHERE ip.ip = '" & ip & "'", cn, adOpenStatic, adLockOptimistic
        If Not rs.RecordCount = 0 Then
            no_client = rs(0)
            If MSFlexGrid.TextMatrix(rs(0), 3) = LangBebas Then
                Winsock(Index).SendData LangBebas & "|_|_|_|_|_|_"
                cn.Close
            ElseIf MSFlexGrid.TextMatrix(rs(0), 3) = LangAktif Then
                cn.Close
                cn.Open cnstr
                rs.ActiveConnection = cn
                rs.Open "SELECT proses.* FROM proses WHERE proses.no_client = " & no_client, cn, adOpenStatic, adLockOptimistic
                If Not rs.RecordCount = 0 Then
                    Winsock(Index).SendData LangAktif & "|" & rs(4) & "|" & rs(2) & "|_|_|_|_|"
                Else
                    MSFlexGrid.TextMatrix(rs(0), 3) = LangBebas
                End If
                cn.Close
            End If
        Else
            Winsock(Index).SendData LangClientBelumTerdaftar & "|_|_|_|_|_|_|"
            cn.Close
        End If
    ElseIf perintah = "connect" Then
        cn.Open cnstr
        rs.ActiveConnection = cn
        rs.Open "SELECT ip.* FROM ip WHERE ip.ip = '" & ip & "'", cn, adOpenStatic, adLockOptimistic
        If Not rs.RecordCount = 0 Then
            no_client = rs(0)
            Winsock(Index).SendData "time|" & Format(Time, "hh:nn:ss") & "|" & Date & "|" & no_client & "|" & interval & "|" & nama_warnet & "|" & alamat1 & "|" & alamat2 & "|" & lang & "|"
        Else
            Winsock(Index).SendData LangClientBelumTerdaftar & "|_|_|_|_|_|_|"
        End If
        cn.Close
    ElseIf perintah = "set" Then
        cn.Open cnstr
        rs.ActiveConnection = cn
        rs.Open "SELECT ip.* FROM ip WHERE ip.ip = '" & ip & "'", cn, adOpenStatic, adLockOptimistic
        If Not rs.RecordCount = 0 Then
            Winsock(Index).SendData "set|" & discount & "|" & jam_awal & "|" & jam_akhir & "|" & harga_personal & "|" & harga_member & "|" & harga_awal & "|" & harga_game & "|" & harga_ketik & "|" & password_admin
        Else
            Winsock(Index).SendData LangClientBelumTerdaftar & "|_|_|_|_|_|_|"
        End If
        cn.Close
    ElseIf perintah = "start" Then
        cn.Open cnstr
        rs.ActiveConnection = cn
        rs.Open "SELECT ip.* FROM ip WHERE ip.ip = '" & ip & "'", cn, adOpenStatic, adLockOptimistic
        If Not rs.RecordCount = 0 Then
            no_client = rs(0)
            If kunci(no_client) = False Then
                MSFlexGrid.TextMatrix(rs(0), 3) = "start"
                MSFlexGrid.TextMatrix(rs(0), 2) = op1
                If op2 = LangPersonal Then
                    MSFlexGrid.TextMatrix(rs(0), 8) = LangPersonal
                ElseIf op2 = LangMember Then
                    MSFlexGrid.TextMatrix(rs(0), 8) = LangMember
                ElseIf op2 = LangGame Then
                    MSFlexGrid.TextMatrix(rs(0), 8) = LangGame
                ElseIf op2 = LangKetik Then
                    MSFlexGrid.TextMatrix(rs(0), 8) = LangKetik
                End If
                mulai(no_client) = Now
                MSFlexGrid.Col = 3
                MSFlexGrid.Row = rs(0)
                MSFlexGrid.CellBackColor = &HC0FFC0 'hijau
                Winsock(Index).SendData "OK|_|_|_|_|_|_|"
            Else
                Winsock(Index).SendData "LOCK|_|_|_|_|_|_|"
            End If
        Else
            Winsock(Index).SendData LangClientBelumTerdaftar & "|_|_|_|_|_|_|"
        End If
        cn.Close
    ElseIf perintah = "stop" Then
        cn.Open cnstr
        rs.ActiveConnection = cn
        rs.Open "SELECT ip.* FROM ip WHERE ip.ip = '" & ip & "'", cn, adOpenStatic, adLockOptimistic
        If Not rs.RecordCount = 0 Then
            If MSFlexGrid.TextMatrix(Val(rs(0)), 3) = LangAktif Then
                no_client = rs(0)
                MSFlexGrid.TextMatrix(rs(0), 3) = "stop"
                MSFlexGrid.Col = 3
                MSFlexGrid.Row = rs(0)
                MSFlexGrid.CellBackColor = &HFFC0C0 'ungu
                Winsock(Index).SendData "OK|_|_|_|_|_|_|"
            End If
        Else
            Winsock(Index).SendData LangClientBelumTerdaftar & "|_|_|_|_|_|_|"
        End If
        cn.Close
    ElseIf perintah = "password" Then
        cn.Open cnstr
        rs.ActiveConnection = cn
        rs.Open "SELECT member.* FROM member WHERE member.user = '" & op1 & "'", cn, adOpenStatic, adLockOptimistic
        If Not rs.RecordCount = 0 Then
            If rs(5) >= deposit_min Then
                Winsock(Index).SendData "password|" & DecryptText(rs(4), "password") & "|ok|_|_|_|_|"
            Else
                Winsock(Index).SendData "password|" & rs(4) & "|no|_|_|_|_|"
            End If
        Else
            Winsock(Index).SendData LangMemberBelumTerdaftar & "|_|_|_|_|_|_|"
        End If
        cn.Close
    ElseIf perintah = "pesan" Then
        cn.Open cnstr
        rs.ActiveConnection = cn
        rs.Open "SELECT ip.* FROM ip WHERE ip.ip = '" & ip & "'", cn, adOpenStatic, adLockOptimistic
        If Not rs.RecordCount = 0 Then
            no_client = rs(0)
            cn.Close
            If FrmPesanLoad = False Then
                Load FrmPesan
                If FExists(App.path & "\" & MsgFilename) Then
                    FrmPesan.TxtPesanDiterima.LoadFile (App.path & "\" & MsgFilename)
                    FrmPesan.TxtPesanDiterima.Text = FrmPesan.TxtPesanDiterima.Text & Chr(13) & LangClient & " " & no_client & ": " & op1
                Else
                    FrmPesan.TxtPesanDiterima.Text = LangClient & " " & no_client & ": " & op1
                End If
                FrmPesan.Show 1
            End If
            If FrmPesan.TxtPesanDiterima.Text = "" Then
                FrmPesan.TxtPesanDiterima.SaveFile (App.path & "\" & MsgFilename)
            Else
                FrmPesan.TxtPesanDiterima.Text = FrmPesan.TxtPesanDiterima.Text & Chr(13) & LangClient & " " & no_client & ": " & op1
                FrmPesan.TxtPesanDiterima.SaveFile (App.path & "\" & MsgFilename)
            End If
        Else
            Winsock(Index).SendData LangClientBelumTerdaftar & "|_|_|_|_|_|_|"
            cn.Close
        End If
    ElseIf perintah = "app" Then
        cn.Open cnstr
        rs.ActiveConnection = cn
        rs.Open "SELECT aplikasi.* FROM aplikasi", cn, adOpenStatic, adLockOptimistic
        If Not rs.RecordCount = 0 Then
            Winsock(Index).SendData "app|" & ketik1 & "|" & ketik2 & "|" & ketik3 & "|" & game1 & "|" & game2 & "|" & game3 & "|"
        Else
            Winsock(Index).SendData LangClientBelumTerdaftar & "|_|_|_|_|_|_|"
        End If
        cn.Close
    ElseIf perintah = "remote" Then
        cn.Open cnstr
        rs.ActiveConnection = cn
        rs.Open "SELECT ip.* FROM ip WHERE ip.ip = '" & ip & "'", cn, adOpenStatic, adLockOptimistic
        If Not rs.RecordCount = 0 Then
            no_client = rs(0)
            remotestr1(no_client) = op1
            remotestr2(no_client) = op2
            Call TampilRemote
        Else
            Winsock(Index).SendData LangClientBelumTerdaftar & "|_|_|_|_|_|_|"
        End If
        cn.Close
    End If
End Sub

Private Sub CekPasswordAdmin()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open "SELECT admin.* FROM admin", cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        rs.MoveFirst
        password_admin = DecryptText(rs(0), "password")
    Else
        password_admin = "snake"
    End If
    cn.Close
End Sub

Private Sub HapusdbSocket()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open "SELECT socket.* FROM socket", cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        rs.MoveFirst
        Do While Not rs.EOF
            rs.Delete
            If rs.EOF Then
                Exit Do
            Else
                rs.MoveFirst
            End If
        Loop
    End If
    cn.Close
End Sub

Private Sub HapusdbIP()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open "SELECT ip.* FROM ip", cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        rs.MoveFirst
        Do While Not rs.EOF
            rs.Delete
            If rs.EOF Then
                Exit Do
            Else
                rs.MoveFirst
            End If
        Loop
    End If
    cn.Close
End Sub

Private Sub FreeSock()
    For i = 1 To SockMax
        iSocketsFree(i) = True
    Next i
End Sub

Private Sub BukaKunci()
    For i = 1 To ClientMax
        kunci(i) = False
        StatusBar.Panels(9).Text = LangBuka
    Next i
End Sub

Private Sub TampilRemote()
Dim substr_remote1() As String
Dim substr_remote2() As String
    substr_remote1() = Split(remotestr1(Val(FrmRemote.TxtNo_Client.Text)), ":")
    substr_remote2() = Split(remotestr2(Val(FrmRemote.TxtNo_Client.Text)), ":")
    If UBound(substr_remote1) > 1 Then
        FrmRemote.Grid.Rows = UBound(substr_remote1) + 1
        FrmRemote.Grid.TextMatrix(0, 0) = LangAplikasi
        FrmRemote.Grid.TextMatrix(0, 1) = LangPID
        FrmRemote.Grid.ColWidth(0) = 2500
        FrmRemote.Grid.ColWidth(1) = 1000
        For i = 0 To UBound(substr_remote1) - 1
            FrmRemote.Grid.TextMatrix(i + 1, 0) = substr_remote1(i)
            FrmRemote.Grid.TextMatrix(i + 1, 1) = substr_remote2(i)
        Next i
        FrmRemote.LabelTotal.Caption = LangTotalProc & UBound(substr_remote1)
    Else
        FrmRemote.Grid.Rows = 1
        FrmRemote.Grid.TextMatrix(0, 0) = LangAplikasi
        FrmRemote.Grid.TextMatrix(0, 1) = LangPID
        FrmRemote.Grid.ColWidth(0) = 2500
        FrmRemote.Grid.ColWidth(1) = 1000
    End If
End Sub
