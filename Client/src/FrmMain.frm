VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "snakebite cybercafe billing system v3.5.6"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerAnim 
      Interval        =   500
      Left            =   6720
      Top             =   1800
   End
   Begin VB.Frame FrameStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   8040
      TabIndex        =   71
      Top             =   7440
      Width           =   3615
      Begin VB.Shape Shape5 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   120
         Width           =   375
      End
      Begin VB.Shape Shape4 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   1680
         Shape           =   3  'Circle
         Top             =   120
         Width           =   375
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   1200
         Shape           =   3  'Circle
         Top             =   120
         Width           =   375
      End
      Begin VB.Label LabelStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   840
         Width           =   3375
      End
      Begin VB.Shape Shape2 
         Height          =   1215
         Left            =   0
         Top             =   0
         Width           =   3615
      End
   End
   Begin VB.Frame FrameMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   7800
      TabIndex        =   68
      Top             =   1920
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CommandButton CmdTutupMsg 
         Caption         =   "Tutup"
         Height          =   375
         Left            =   1440
         TabIndex        =   70
         Top             =   960
         Width           =   975
      End
      Begin VB.Label LabelMsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   480
         Width           =   3615
      End
      Begin VB.Shape Shape1 
         Height          =   1575
         Left            =   0
         Top             =   0
         Width           =   3855
      End
   End
   Begin VB.Timer TimerMonitorGame 
      Interval        =   1000
      Left            =   6240
      Top             =   1800
   End
   Begin VB.Frame FrameKetik 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " Login Rental "
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   3960
      TabIndex        =   31
      Top             =   5520
      Visible         =   0   'False
      Width           =   3615
      Begin VB.TextBox TxtUserKetik 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   1320
         MaxLength       =   25
         TabIndex        =   34
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton CmdStartKetik 
         Caption         =   "Start"
         Height          =   375
         Left            =   1320
         TabIndex        =   33
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton CmdBackKetik 
         Caption         =   "Kembali"
         Height          =   375
         Left            =   2400
         TabIndex        =   32
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama User"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   480
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView ProsesMati 
      Height          =   2055
      Left            =   3840
      TabIndex        =   62
      Top             =   5280
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   3625
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Proses"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "prosesID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Thread"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Parent"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Priority"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Exe Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView listProses 
      Height          =   2055
      Left            =   3840
      TabIndex        =   61
      Top             =   4680
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   3625
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Proses"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "prosesID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Thread"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Parent"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Priority"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Exe Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Timer TimerMonitorKetik 
      Interval        =   1000
      Left            =   5760
      Top             =   1800
   End
   Begin VB.Timer TimerFirewall 
      Interval        =   1000
      Left            =   5280
      Top             =   1800
   End
   Begin VB.Frame FrameProfile 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   3840
      TabIndex        =   55
      Top             =   120
      Width           =   7815
      Begin VB.Label Txtalamat2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   240
         TabIndex        =   58
         Top             =   960
         Width           =   7335
      End
      Begin VB.Label Txtalamat1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   240
         TabIndex        =   57
         Top             =   600
         Width           =   7335
      End
      Begin VB.Label TxtNama_Warnet 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   56
         Top             =   240
         Width           =   7335
      End
   End
   Begin VB.Frame FrameBill 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   43
      Top             =   7920
      Visible         =   0   'False
      Width           =   9650
      Begin VB.CommandButton CmdPesan 
         Caption         =   "Pesan"
         Height          =   375
         Left            =   8640
         TabIndex        =   53
         Top             =   105
         Width           =   975
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "Stop"
         Height          =   375
         Left            =   10
         TabIndex        =   44
         Top             =   105
         Width           =   840
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Total :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6600
         TabIndex        =   52
         Top             =   150
         Width           =   615
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Discount :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4440
         TabIndex        =   51
         Top             =   150
         Width           =   855
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Biaya :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   50
         Top             =   150
         Width           =   600
      End
      Begin VB.Label LabelTotal 
         Alignment       =   2  'Center
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
         Height          =   375
         Left            =   7320
         TabIndex        =   49
         Top             =   105
         Width           =   1215
      End
      Begin VB.Label LabelDiscount 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5400
         TabIndex        =   48
         Top             =   105
         Width           =   1095
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Durasi :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   47
         Top             =   150
         Width           =   615
      End
      Begin VB.Label LabelBiaya 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3240
         TabIndex        =   46
         Top             =   105
         Width           =   1095
      End
      Begin VB.Label LabelDurasi 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1560
         TabIndex        =   45
         Top             =   105
         Width           =   975
      End
   End
   Begin VB.Frame FrameMember 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " Login Member "
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   7560
      TabIndex        =   36
      Top             =   5280
      Visible         =   0   'False
      Width           =   3615
      Begin VB.CommandButton CmdBackMember 
         Caption         =   "Kembali"
         Height          =   375
         Left            =   2400
         TabIndex        =   40
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton CmdStartMember 
         Caption         =   "Login"
         Height          =   375
         Left            =   1320
         TabIndex        =   39
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox TxtPasswordMember 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1320
         MaxLength       =   25
         PasswordChar    =   "*"
         TabIndex        =   38
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox TxtUserMember 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   1320
         MaxLength       =   25
         TabIndex        =   37
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama User"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame FrameGame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " Login Game "
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      TabIndex        =   26
      Top             =   5160
      Visible         =   0   'False
      Width           =   3615
      Begin VB.TextBox TxtUserGame 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   1320
         MaxLength       =   25
         TabIndex        =   29
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton CmdStartGame 
         Caption         =   "Start"
         Height          =   375
         Left            =   1320
         TabIndex        =   28
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton CmdBackGame 
         Caption         =   "Kembali"
         Height          =   375
         Left            =   2400
         TabIndex        =   27
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama User"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame FramePersonal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " Login Personal "
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      TabIndex        =   21
      Top             =   3480
      Visible         =   0   'False
      Width           =   3615
      Begin VB.CommandButton CmdBackPers 
         Caption         =   "Kembali"
         Height          =   375
         Left            =   2400
         TabIndex        =   24
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton CmdStartPersonal 
         Caption         =   "Start"
         Height          =   375
         Left            =   1320
         TabIndex        =   23
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox TxtUserPersonal 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   1320
         MaxLength       =   25
         TabIndex        =   22
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama User"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame FrameAdmin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " Login Admin "
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   3615
      Begin VB.CommandButton CmdBatal 
         Caption         =   "Kembali"
         Height          =   375
         Left            =   2400
         TabIndex        =   19
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton CmdSeting 
         Caption         =   "Seting"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton CmdTutup 
         Caption         =   "Tutup"
         Height          =   375
         Left            =   1320
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox TxtPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1320
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   15
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame FrameLogin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3615
      Begin VB.CommandButton CmdAdmin 
         Caption         =   "Admin"
         Height          =   375
         Left            =   2400
         TabIndex        =   13
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton CmdKetik 
         Caption         =   "Mengetik"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1320
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton CmdGame 
         Caption         =   "Game"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1320
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton CmdMember 
         Caption         =   "Member"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton CmdPersonal 
         Caption         =   "Personal"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame FrameSeting 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " Client Seting "
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   7440
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   3615
      Begin VB.CheckBox optketik 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ketik"
         Height          =   375
         Left            =   2400
         TabIndex        =   66
         Top             =   2280
         Width           =   975
      End
      Begin VB.CheckBox optgame 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Game"
         Height          =   375
         Left            =   2400
         TabIndex        =   65
         Top             =   1800
         Width           =   975
      End
      Begin VB.CheckBox optmember 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Member"
         Height          =   375
         Left            =   1320
         TabIndex        =   64
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CheckBox optpersonal 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Personal"
         Height          =   375
         Left            =   1320
         TabIndex        =   63
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton CmdCanSet 
         Caption         =   "Kembali"
         Height          =   375
         Left            =   2400
         TabIndex        =   20
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton CmdSetClient 
         Caption         =   "Update"
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox TxtNo_Client 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   4
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox TxtPort 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   3
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox TxtIPServer 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   1320
         MaxLength       =   25
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Akses Login"
         Height          =   255
         Left            =   240
         TabIndex        =   67
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Client"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Port"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
      Begin VB.Label ServerIP 
         BackStyle       =   0  'Transparent
         Caption         =   "IP Server"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Timer TimerBill 
      Interval        =   1000
      Left            =   4800
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4320
      Top             =   1800
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   3840
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListProc 
      Height          =   1215
      Left            =   3840
      TabIndex        =   59
      Top             =   2280
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2143
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Aplikasi"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "pID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Local Adress"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Local Port"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Remote Adress"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Remote Port"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView ListTerminate 
      Height          =   1935
      Left            =   3840
      TabIndex        =   60
      Top             =   3600
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   3413
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Aplikasi"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "pID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Local Adress"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Local Port"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Remote Adress"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Remote Port"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label LabelAdv 
      BackStyle       =   0  'Transparent
      Caption         =   "snakebite cybercafe billing system 3.6.0  home page : http://snakebite.sourceforge.net  email : snakebites@users.sourceforge.net"
      Height          =   255
      Left            =   120
      TabIndex        =   54
      Top             =   8520
      Width           =   9615
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   120
      Top             =   6840
      Width           =   2175
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
'*********************************************************************************************************************
'* syntax = perintah|op1|op2|op3|op4|op5|op6|op7|op8|op9|                                                            *
'* time = time|time|date|no_client|interval|nama_warnet|alamat1|alamat2|                                             *
'* set = set|discount|jam_awal|jam_akhir|harga_personal|harga_member|harga_awal|harga_game|harga_ketik|password_admin*
'* pindahke = pindahke|_|                                                                                            *
'* pindahdari = pindahdari|mulai|jenis|                                                                              *
'* aktif = aktif|mulai|jenis|                                                                                        *
'* stop = stop|_|                                                                                                    *
'* shutdown = shutdown|_|                                                                                            *
'* pesan = pesan|pesannya|                                                                                           *
'* *******************************************************************************************************************
'def firewall
Dim Firewall As Boolean
'def monitor
Dim MonitorKetik As Boolean
Dim MonitorGame As Boolean
'def status client
Dim status As String
Dim jenis As String
'def seting client
Dim port As String
Dim password_admin As String
Dim no_client As Long
Dim ipserver As String
Dim nama_warnet As String
Dim alamat1 As String
Dim alamat2 As String
' def member
Dim member As String
Dim member_pass As String
'def billing
Dim mulai As Date
Dim durasi As Date
Dim durasi_disc As Date
Dim biaya As Single
Dim biaya_disc As Single
Dim mulai_disc As Date
Dim akhir_disc As Date
Dim disc_mulai As Date
Dim interval As Long
Dim atas As Long
Dim bawah As Long
'def seting server
Dim discount As Long
Dim jam_awal As String
Dim jam_akhir As String
Dim harga_personal As Single
Dim harga_member As Single
Dim harga_game As Single
Dim harga_ketik As Single
Dim harga_awal As Single
'akses login
Dim lpersonal As String
Dim lmember As String
Dim lgame As String
Dim lketik As String

Private Sub setBahasa()
    CmdPersonal.Caption = LangJenisPersonal
    CmdGame.Caption = LangJenisGame
    CmdMember.Caption = LangJenisMember
    CmdKetik.Caption = LangJenisKetik
    CmdBatal.Caption = LangBatal
    CmdBackPers.Caption = LangBatal
    CmdBackMember.Caption = LangBatal
    CmdBackGame.Caption = LangBatal
    CmdBackKetik.Caption = LangBatal
    FramePersonal.Caption = LangLogin & " " & LangJenisPersonal
    FrameMember.Caption = LangLogin & " " & LangJenisMember
    FrameGame.Caption = LangLogin & " " & LangJenisGame
    FrameKetik.Caption = LangLogin & " " & LangJenisKetik
    CmdPesan.Caption = LangmnuPesan
    Label10.Caption = LangDurasi
    Label13.Caption = LangBiaya
    Label14.Caption = LangDiscount
    Label15.Caption = LangTotal
    optpersonal.Caption = LangJenisPersonal
    optmember.Caption = LangJenisMember
    optgame.Caption = LangJenisGame
    optketik.Caption = LangJenisKetik
    Label9.Caption = LangLogin
    CmdTutup.Caption = LangTutup
    CmdTutupMsg.Caption = LangTutup
    FrameMsg.Caption = LangInformasi
    CmdCanSet.Caption = LangBatal
End Sub
Private Sub CmdAdmin_Click()
    FrameLogin.Visible = False
    FrameAdmin.Left = 120
    FrameAdmin.Top = 120
    FrameAdmin.Visible = True
    TxtPassword.SetFocus
End Sub

Private Sub CmdBackGame_Click()
    FrameGame.Visible = False
    TxtUserGame.Text = ""
    FrameLogin.Visible = True
    If CmdPersonal.Enabled = True Then
        CmdPersonal.SetFocus
    End If
End Sub

Private Sub CmdBackKetik_Click()
    FrameKetik.Visible = False
    TxtUserKetik.Text = ""
    FrameLogin.Visible = True
    If CmdPersonal.Enabled = True Then
        CmdPersonal.SetFocus
    End If
End Sub

Private Sub CmdBackMember_Click()
    FrameMember.Visible = False
    TxtUserMember.Text = ""
    TxtPasswordMember.Text = ""
    CmdStartMember.Caption = LangLogin
    FrameLogin.Visible = True
    If CmdPersonal.Enabled = True Then
        CmdPersonal.SetFocus
    End If
End Sub

Private Sub CmdBackPers_Click()
    FramePersonal.Visible = False
    TxtUserPersonal.Text = ""
    FrameLogin.Visible = True
    If CmdPersonal.Enabled = True Then
        CmdPersonal.SetFocus
    End If
End Sub

Private Sub CmdBatal_Click()
    FrameAdmin.Visible = False
    FrameLogin.Visible = True
End Sub

Private Sub CmdCanSet_Click()
    FrameSeting.Visible = False
    FrameLogin.Visible = True
    If CmdPersonal.Enabled = True Then
        CmdPersonal.SetFocus
    Else
        CmdAdmin.SetFocus
    End If
End Sub

Private Sub CmdGame_Click()
    FrameGame.Left = 120
    FrameGame.Top = 120
    FrameLogin.Visible = False
    FrameGame.Visible = True
    TxtUserGame.Text = ""
    TxtUserGame.SetFocus
End Sub

Private Sub CmdKetik_Click()
    FrameKetik.Left = 120
    FrameKetik.Top = 120
    FrameLogin.Visible = False
    FrameKetik.Visible = True
    TxtUserKetik.Text = ""
    TxtUserKetik.SetFocus
End Sub

Private Sub CmdMember_Click()
    FrameMember.Left = 120
    FrameMember.Top = 120
    FrameLogin.Visible = False
    FrameMember.Visible = True
    TxtUserMember.SetFocus
    CmdStartMember.Caption = LangLogin
    TxtUserMember.Text = ""
    TxtPasswordMember.Text = ""
End Sub

Private Sub CmdPersonal_Click()
    FramePersonal.Left = 120
    FramePersonal.Top = 120
    FrameLogin.Visible = False
    FramePersonal.Visible = True
    TxtUserPersonal.Text = ""
    TxtUserPersonal.SetFocus
End Sub

Private Sub CmdPesan_Click()
    Load FrmPesan
    FrmPesan.Left = 0
    FrmPesan.Top = FrmMain.Height + 10
    FrmPesan.Show 1
End Sub

Private Sub CmdSetClient_Click()
    ipserver = TxtIPServer.Text
    port = TxtPort.Text
    no_client = Val(TxtNo_Client.Text)
    lpersonal = optpersonal.value
    lmember = optmember.value
    lgame = optgame.value
    lketik = optketik.value
    status = "waitregister"
    FrameSeting.Visible = False
    FrameLogin.Visible = True
    If CmdPersonal.Enabled = True Then
        CmdPersonal.SetFocus
    Else
        CmdAdmin.SetFocus
    End If
    Call TulisKey
    Call InitAksesLogin
End Sub

Private Sub CmdSeting_Click()
    If password_admin <> "" Then
        If TxtPassword.Text = password_admin Then
            FrameAdmin.Visible = False
            TxtPassword.Text = ""
            FrameSeting.Top = 120
            FrameSeting.Left = 120
            FrameSeting.Visible = True
            TxtIPServer.SetFocus
        Else
            If FrameMsg.Visible = False Then
                FrameMsg.Visible = True
                LabelMsg.Caption = LangPasswordSalah
            Else
                LabelMsg.Caption = LangPasswordSalah
            End If
            TxtPassword.Text = ""
            TxtPassword.SetFocus
        End If
    Else
        If TxtPassword.Text = "snake" Then
            FrameAdmin.Visible = False
            TxtPassword.Text = ""
            FrameSeting.Top = 120
            FrameSeting.Left = 120
            FrameSeting.Visible = True
            TxtIPServer.SetFocus
        End If
    End If
End Sub

Private Sub CmdStartGame_Click()
    If Not TxtUserGame.Text = "" Then
        jenis = LangGame
        Winsock.SendData "start|" & Winsock.LocalIP & "|" & TxtUserGame.Text & "|" & jenis & "|"
        status = "waitstart"
    Else
        If FrameMsg.Visible = False Then
            FrameMsg.Visible = True
            LabelMsg.Caption = LangUserKosong
        Else
            LabelMsg.Caption = LangUserKosong
        End If
        TxtUserGame.SetFocus
    End If
End Sub

Private Sub CmdStartKetik_Click()
    If Not TxtUserKetik.Text = "" Then
        jenis = LangKetik
        Winsock.SendData "start|" & Winsock.LocalIP & "|" & TxtUserKetik.Text & "|" & jenis & "|"
        status = "waitstart"
    Else
        If FrameMsg.Visible = False Then
            FrameMsg.Visible = True
            LabelMsg.Caption = LangUserKosong
        Else
            LabelMsg.Caption = LangUserKosong
        End If
        TxtUserKetik.SetFocus
    End If
End Sub

Private Sub CmdStartMember_Click()
    If Not TxtUserMember.Text = "" Then
        LoginMember
    Else
        If FrameMsg.Visible = False Then
            FrameMsg.Visible = True
            LabelMsg.Caption = LangUserKosong
        Else
            LabelMsg.Caption = LangUserKosong
        End If
        TxtUserMember.SetFocus
    End If
End Sub
Private Sub CmdStartPersonal_Click()
    If Not TxtUserPersonal.Text = "" Then
        jenis = LangPersonal
        Winsock.SendData "start|" & Winsock.LocalIP & "|" & TxtUserPersonal.Text & "|" & jenis & "|"
        status = "waitstart"
    Else
        If FrameMsg.Visible = False Then
            FrameMsg.Visible = True
            LabelMsg.Caption = LangUserKosong
        Else
            LabelMsg.Caption = LangUserKosong
        End If
        TxtUserPersonal.SetFocus
    End If
End Sub

Private Sub CmdStop_Click()
    If Winsock.State = sckConnected Then
        status = LangBebas
        Firewall = False
        MonitorKetik = False
        MonitorGame = False
        Winsock.SendData "stop|" & Winsock.LocalIP & "|" & member & "|" & jenis & "|"
        FrameBill.Visible = False
        SetFormMain
        CmdPersonal.Enabled = True
        CmdMember.Enabled = True
        CmdGame.Enabled = True
        CmdKetik.Enabled = True
        FrameLogin.Visible = True
        If CmdPersonal.Enabled = True Then
            CmdPersonal.SetFocus
        End If
        If frmpesanload = True Then
            Unload FrmPesan
        End If
    End If
    Call InitAksesLogin
End Sub

Private Sub CmdTutup_Click()
    If password_admin <> "" Then
        If TxtPassword.Text = password_admin Then
            Call MatikanCtrlAltDel(False)
            End
        Else
            If FrameMsg.Visible = False Then
                FrameMsg.Visible = True
                LabelMsg.Caption = LangPasswordSalah
            Else
                LabelMsg.Caption = LangPasswordSalah
            End If
            TxtPassword.Text = ""
            TxtPassword.SetFocus
        End If
    Else
        If TxtPassword.Text = "snake" Then
            Call MatikanCtrlAltDel(False)
            End
        End If
    End If
End Sub

Private Sub CmdTutupMsg_Click()
    FrameMsg.Visible = False
End Sub

Private Sub Form_Load()
    Dim strset As String
    Dim substrset() As String
    If FExists(App.Path & "\bg.jpg") Then
        Image1.Picture = LoadPicture(App.Path & "\bg.jpg")
        Image1.Top = (Screen.Height - Image1.Height) / 2
        Image1.Left = (Screen.Width - Image1.Width) / 2
    End If
    Call MatikanCtrlAltDel(True)
    Call SetFormMain
    FrameStatus.Left = FrmMain.Width - FrameStatus.Width - 300
    FrameStatus.Top = FrmMain.Height - FrameStatus.Height - 300
    LabelAdv.Top = FrmMain.Height - LabelAdv.Height - 200
    LabelAdv.Caption = "snakebite cybercafe billing system " & VersiAplikasi & " home page : http://snakebite.sourceforge.net  email : snakebites@users.sourceforge.net"
    FrameProfile.Left = (Screen.Width - FrameProfile.Width) / 2
    FrameMsg.Left = (Screen.Width - FrameMsg.Width) / 2
    FrameMsg.Top = (Screen.Height - FrameMsg.Height) / 2
    Call CekKey
    Firewall = False
    MonitorKetik = False
    MonitorGame = False
    Call Bahasa
    Call InitAksesLogin
    If App.PrevInstance = True Then
        End
    End If
End Sub
Private Sub CekKey()
Dim x As Long
    x = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\snakebite\client\" & VersiAplikasi, 0, KEY_READ, hregkey)
    If x <> 0 Then
        status = "register"
        CmdPersonal.Enabled = False
        CmdMember.Enabled = False
        CmdGame.Enabled = False
        CmdKetik.Enabled = False
    Else
        Call AmbilKey
        status = "connect"
        CmdPersonal.Enabled = False
        CmdMember.Enabled = False
        CmdGame.Enabled = False
        CmdKetik.Enabled = False
    End If
    x = RegCloseKey(hregkey)
End Sub
Private Sub AmbilKey()
Dim x As Long
    x = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\snakebite\client\" & VersiAplikasi, 0, KEY_READ, hregkey)
    ipserver = GetRegistryValue(hregkey, "ipserver")
    TxtIPServer = ipserver
    port = GetRegistryValue(hregkey, "port")
    TxtPort = port
    TxtNo_Client = GetRegistryValue(hregkey, "noclient")
    lpersonal = GetRegistryValue(hregkey, "lpersonal")
    lmember = GetRegistryValue(hregkey, "lmember")
    lgame = GetRegistryValue(hregkey, "lgame")
    lketik = GetRegistryValue(hregkey, "lketik")
    x = RegCloseKey(hregkey)
End Sub
Private Sub TulisKey()
    Call RegSaveString(HKEY_LOCAL_MACHINE, "Software\snakebite\client\" & VersiAplikasi & "\", "ipserver", TxtIPServer)
    Call RegSaveString(HKEY_LOCAL_MACHINE, "Software\snakebite\client\" & VersiAplikasi & "\", "port", TxtPort)
    Call RegSaveString(HKEY_LOCAL_MACHINE, "Software\snakebite\client\" & VersiAplikasi & "\", "noclient", TxtNo_Client)
    Call RegSaveString(HKEY_LOCAL_MACHINE, "Software\snakebite\client\" & VersiAplikasi & "\", "lpersonal", optpersonal.value)
    Call RegSaveString(HKEY_LOCAL_MACHINE, "Software\snakebite\client\" & VersiAplikasi & "\", "lmember", optmember.value)
    Call RegSaveString(HKEY_LOCAL_MACHINE, "Software\snakebite\client\" & VersiAplikasi & "\", "lgame", optgame.value)
    Call RegSaveString(HKEY_LOCAL_MACHINE, "Software\snakebite\client\" & VersiAplikasi & "\", "lketik", optketik.value)
End Sub

Private Sub Timer1_Timer()
    If Winsock.State = 7 Then
        If status = "connect" Then
            Winsock.SendData "connect|" & Winsock.LocalIP & "|_|_|_|"
            status = "waitconnect"
        ElseIf status = "set" Then
            Winsock.SendData "set|" & Winsock.LocalIP & "|_|_|_|"
            status = "waitset"
        ElseIf status = "apset" Then
            Winsock.SendData "app|" & Winsock.LocalIP & "|_|_|_|"
            status = "waitapset"
        ElseIf status = "cekstatus" Then
            Winsock.SendData "status|" & Winsock.LocalIP & "|_|_|_|"
            status = "waitcekstatus"
            mulai_disc = CDate(Format(Date, "dd/mm/yyyy") & " " & jam_awal)
            If CDate(jam_akhir) < CDate(jam_awal) Then
                akhir_disc = CDate(Format(Date + 1, "dd/mm/yyyy") & " " & jam_akhir)
            Else
                akhir_disc = CDate(Format(Date, "dd/mm/yyyy") & " " & jam_akhir)
            End If
        ElseIf status = "waitregister" Then
            Winsock.SendData "register|" & Winsock.LocalIP & "|" & no_client
            status = "okregister"
        End If
        LabelStatus.Caption = LangTerkonek
        TimerAnim.Enabled = False
        Shape3.FillColor = vbWhite
        Shape4.FillColor = vbWhite
        Shape5.FillColor = vbGreen
    ElseIf Winsock.State = sckClosed Then
        If Not status = "register" Then
            Winsock.Connect ipserver, port
            LabelStatus.Caption = LangStatusClientKonek & " " & ipserver
        Else
            LabelStatus.Caption = LangStatusClientSet
            TimerAnim.Enabled = False
            Shape3.FillColor = vbRed
            Shape4.FillColor = vbWhite
            Shape5.FillColor = vbWhite
        End If
    Else
        Winsock.Close
        LabelStatus.Caption = LangStatusClientKonek & " " & ipserver
    End If
End Sub


Private Sub TimerAnim_Timer()
    If Shape3.FillColor = vbRed Then
        Shape3.FillColor = vbWhite
    Else
        Shape3.FillColor = vbRed
    End If
    If Shape4.FillColor = vbYellow Then
        Shape4.FillColor = vbWhite
    Else
        Shape4.FillColor = vbYellow
    End If
    If Shape5.FillColor = vbGreen Then
        Shape5.FillColor = vbWhite
    Else
        Shape5.FillColor = vbGreen
    End If
End Sub

Private Sub TimerBill_Timer()
    If (Now > mulai_disc And Now < akhir_disc) Then
            If status = LangAktif Then
                durasi_disc = Now - disc_mulai
                durasi = Now - mulai
                atas = Int(Val(Format(durasi, "nn")) / interval)
                bawah = atas + 1
                If jenis = LangPersonal Then
                    If Val(Format(durasi, "nn")) / interval >= atas And Val(Format(durasi, "nn")) / interval <= bawah Then
                        biaya = Round((Val(Mid(Format(durasi, "hh:nn:ss"), 1, 2)) * harga_personal * 60 / interval) + (bawah * harga_personal), 2)
                        biaya_disc = Round(((Val(Mid(Format(durasi_disc, "hh:nn:ss"), 1, 2)) * harga_personal * 60 / interval) + (bawah * harga_personal)) * discount / 100, 2)
                    End If
                    If biaya <= harga_awal Then
                        biaya = harga_awal
                    End If
                ElseIf jenis = LangMember Then
                    If Val(Format(durasi, "nn")) / interval >= atas And Val(Format(durasi, "nn")) / interval <= bawah Then
                        biaya = Round((Val(Mid(Format(durasi, "hh:nn:ss"), 1, 2)) * harga_member * 60 / interval) + (bawah * harga_member), 2)
                        biaya_disc = Round(((Val(Mid(Format(durasi_disc, "hh:nn:ss"), 1, 2)) * harga_member * 60 / interval) + (bawah * harga_member)) * discount / 100, 2)
                    End If
                    If biaya <= harga_awal Then
                        biaya = harga_awal
                    End If
                ElseIf jenis = "admin" Then
                    biaya = 0
                    biaya_disc = 0
                ElseIf jenis = LangGame Then
                    If Val(Format(durasi, "nn")) / interval >= atas And Val(Format(durasi, "nn")) / interval <= bawah Then
                        biaya = Round((Val(Mid(Format(durasi, "hh:nn:ss"), 1, 2)) * harga_game * 60 / interval) + (bawah * harga_game), 2)
                        biaya_disc = Round(((Val(Mid(Format(durasi_disc, "hh:nn:ss"), 1, 2)) * harga_game * 60 / interval) + (bawah * harga_game)) * discount / 100, 2)
                    End If
                    If biaya <= harga_awal Then
                        biaya = harga_awal
                    End If
                ElseIf jenis = LangKetik Then
                    If Val(Format(durasi, "nn")) / interval >= atas And Val(Format(durasi, "nn")) / interval <= bawah Then
                        biaya = Round((Val(Mid(Format(durasi, "hh:nn:ss"), 1, 2)) * harga_ketik * 60 / interval) + (bawah * harga_ketik), 2)
                        biaya_disc = Round(((Val(Mid(Format(durasi_disc, "hh:nn:ss"), 1, 2)) * harga_ketik * 60 / interval) + (bawah * harga_ketik)) * discount / 100, 2)
                    End If
                    If biaya <= harga_awal Then
                        biaya = harga_awal
                    End If
                End If
                LabelTotal.Caption = Format(biaya - biaya_disc, "##,##0.00")
                LabelDurasi.Caption = Format(durasi, "hh:nn:ss")
                LabelBiaya.Caption = Format(biaya, "##,##0.00")
                LabelDiscount.Caption = Format(biaya_disc, "##,##0.00")
            End If
    Else
            If status = LangAktif Then
                durasi = Now - mulai
                atas = Int(Val(Format(durasi, "nn")) / interval)
                bawah = atas + 1
                If jenis = LangPersonal Then
                    If Val(Format(durasi, "nn")) / interval >= atas And Val(Format(durasi, "nn")) / interval <= bawah Then
                        biaya = Round((Val(Mid(Format(durasi, "hh:nn:ss"), 1, 2)) * harga_personal * 60 / interval) + (bawah * harga_personal), 2)
                        biaya_disc = Round(((Val(Mid(Format(durasi_disc, "hh:nn:ss"), 1, 2)) * harga_personal * 60 / interval) + (bawah * harga_personal)) * discount / 100, 2)
                    End If
                    If biaya <= harga_awal Then
                        biaya = harga_awal
                    End If
                ElseIf jenis = LangMember Then
                    If Val(Format(durasi, "nn")) / interval >= atas And Val(Format(durasi, "nn")) / interval <= bawah Then
                        biaya = Round((Val(Mid(Format(durasi, "hh:nn:ss"), 1, 2)) * harga_member * 60 / interval) + (bawah * harga_member), 2)
                        biaya_disc = Round(((Val(Mid(Format(durasi_disc, "hh:nn:ss"), 1, 2)) * harga_member * 60 / interval) + (bawah * harga_member)) * discount / 100, 2)
                    End If
                    If biaya <= harga_awal Then
                        biaya = harga_awal
                    End If
                ElseIf jenis = "admin" Then
                    biaya = 0
                    biaya_disc = 0
                ElseIf jenis = LangGame Then
                    If Val(Format(durasi, "nn")) / interval >= atas And Val(Format(durasi, "nn")) / interval <= bawah Then
                        biaya = Round((Val(Mid(Format(durasi, "hh:nn:ss"), 1, 2)) * harga_game * 60 / interval) + (bawah * harga_game), 2)
                        biaya_disc = Round(((Val(Mid(Format(durasi_disc, "hh:nn:ss"), 1, 2)) * harga_game * 60 / interval) + (bawah * harga_game)) * discount / 100, 2)
                    End If
                    If biaya <= harga_awal Then
                        biaya = harga_awal
                    End If
                ElseIf jenis = LangKetik Then
                    If Val(Format(durasi, "nn")) / interval >= atas And Val(Format(durasi, "nn")) / interval <= bawah Then
                        biaya = Round((Val(Mid(Format(durasi, "hh:nn:ss"), 1, 2)) * harga_ketik * 60 / interval) + (bawah * harga_ketik), 2)
                        biaya_disc = Round(((Val(Mid(Format(durasi_disc, "hh:nn:ss"), 1, 2)) * harga_ketik * 60 / interval) + (bawah * harga_ketik)) * discount / 100, 2)
                    End If
                    If biaya <= harga_awal Then
                        biaya = harga_awal
                    End If
                End If
                If Format(durasi_disc, "hh:nn:ss") = "00:00:00" Then
                    biaya_disc = 0
                    LabelTotal.Caption = Format(biaya - biaya_disc, "##,##0.00")
                    LabelDiscount.Caption = Format(biaya_disc, "##,##0.00")
                Else
                    LabelTotal.Caption = Format(biaya - biaya_disc, "##,##0.00")
                    LabelDiscount.Caption = Format(biaya_disc, "##,##0.00")
                End If
                LabelDurasi.Caption = Format(durasi, "hh:nn:ss")
                LabelBiaya.Caption = Format(biaya, "##,##0.00")
            End If
    End If
End Sub

Private Sub TimerFirewall_Timer()
    If Firewall = True Then
        If isWinXp = True Then
            ListProc.ListItems.Clear
            Call RefreshStack
            Call EnumEntries
            Call TerminateList
            Call TerminateProc
        End If
    End If
End Sub

Private Sub TimerMonitorGame_Timer()
    If MonitorGame = True Then
        listProses.ListItems.Clear
        Call enumProc
        Call ListGame
        Call MatikanGame
    End If
End Sub

Private Sub TimerMonitorKetik_Timer()
    If MonitorKetik = True Then
        listProses.ListItems.Clear
        Call enumProc
        Call ListKetik
        Call MatikanKetik
    End If
End Sub

Private Sub TxtIPServer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtPort.SetFocus
    End If
End Sub

Private Sub TxtNo_Client_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 13 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPasswordMember_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdStartMember_Click
    End If
End Sub

Private Sub TxtPort_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 13 Or KeyAscii = 8) Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        TxtNo_Client.SetFocus
    End If
End Sub

Private Sub TxtUserGame_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdStartGame_Click
    End If
End Sub

Private Sub TxtUserKetik_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdStartKetik_Click
    End If
End Sub

Private Sub TxtUserMember_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtPasswordMember.SetFocus
    End If
End Sub

Private Sub TxtUserPersonal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdStartPersonal_Click
    End If
End Sub

Private Sub Winsock_Close()
    If Not status = LangAktif Then
        status = "connect"
    End If
End Sub

Private Sub Winsock_ConnectionRequest(ByVal requestID As Long)
    If Winsock.State <> sckClosed Then Winsock.Close
    Winsock.Accept requestID
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
    Dim str As String
    Dim sub_str() As String
    Dim perintah As String
    Dim op1 As String
    Dim op2 As String
    Dim op3 As String
    Dim op4 As String
    Dim op5 As String
    Dim op6 As String
    Dim op7 As String
    Dim op8 As String
    Dim op9 As String
    Winsock.GetData str
    sub_str() = Split(str, "|")
    If UBound(sub_str) >= 0 Then
        perintah = sub_str(0)
    End If
    If UBound(sub_str) >= 1 Then
        op1 = sub_str(1)
    End If
    If UBound(sub_str) >= 2 Then
        op2 = sub_str(2)
    End If
    If UBound(sub_str) >= 3 Then
        op3 = sub_str(3)
    End If
    If UBound(sub_str) >= 4 Then
        op4 = sub_str(4)
    End If
    If UBound(sub_str) >= 5 Then
        op5 = sub_str(5)
    End If
    If UBound(sub_str) >= 6 Then
        op6 = sub_str(6)
    End If
    If UBound(sub_str) >= 7 Then
        op7 = sub_str(7)
    End If
    If UBound(sub_str) >= 8 Then
        op8 = sub_str(8)
    End If
    If UBound(sub_str) >= 9 Then
        op9 = sub_str(9)
    End If
    If status = "waitconnect" Then
        If perintah = "time" Then
            Time = CDate(op1)
            Date = CDate(op2)
            no_client = Val(op3)
            interval = Val(op4)
            nama_warnet = op5
            alamat1 = op6
            alamat2 = op7
            lang = op8
            Call Bahasa
            Call setBahasa
            FrameLogin.Caption = LangClient & " " & no_client & " "
            TxtNama_Warnet.Caption = nama_warnet
            Txtalamat1.Caption = alamat1
            Txtalamat2.Caption = alamat2
            status = "set"
        End If
    ElseIf status = "waitset" Then
        If perintah = "set" Then
            discount = Val(op1)
            jam_awal = op2
            jam_akhir = op3
            harga_personal = Val(op4)
            harga_member = Val(op5)
            harga_awal = Val(op6)
            harga_game = Val(op7)
            harga_ketik = Val(op8)
            password_admin = op9
            status = "apset"
        End If
    ElseIf status = "waitapset" Then
        If perintah = "app" Then
            ketik1 = op1
            ketik2 = op2
            ketik3 = op3
            game1 = op4
            game2 = op5
            game3 = op6
            status = "cekstatus"
        End If
    ElseIf status = "waitcekstatus" Then
        If perintah = LangBebas Then
            status = LangBebas
            CmdPersonal.Enabled = True
            CmdMember.Enabled = True
            CmdGame.Enabled = True
            CmdKetik.Enabled = True
            If FrameLogin.Visible = True Then
                If CmdPersonal.Enabled = True Then
                    CmdPersonal.SetFocus
                End If
            End If
        ElseIf perintah = LangAktif Then
            mulai = CDate(op1)
            If mulai >= mulai_disc Then
                disc_mulai = mulai
            Else
                disc_mulai = mulai_disc
            End If
            jenis = op2
            FrameBill.Top = -100
            FrameBill.Left = 0
            FrameLogin.Visible = False
            FrmMain.Width = FrameBill.Width
            FrmMain.Height = FrameBill.Height - 100
            FrameBill.Visible = True
            FrameProfile.Visible = False
            status = LangAktif
            If jenis = LangKetik Then
                Firewall = True
                MonitorKetik = False
                MonitorGame = True
            ElseIf jenis = LangGame Then
                Firewall = True
                MonitorGame = False
                MonitorKetik = True
            Else
                Firewall = False
                MonitorGame = False
                MonitorKetik = False
            End If
        End If
    ElseIf status = "waitstart" Then
        If perintah = "OK" Then
            If jenis = LangPersonal Then
                LoginPersonal
            ElseIf jenis = LangGame Then
                LoginGame
            ElseIf jenis = LangKetik Then
                LoginKetik
            ElseIf jenis = LangMember Then
                StartMember
            End If
        ElseIf perintah = "LOCK" Then
            If jenis = LangPersonal Then
                TxtUserPersonal.Text = ""
                TxtUserPersonal.SetFocus
                status = LangBebas
                If FrameMsg.Visible = False Then
                    FrameMsg.Visible = True
                    LabelMsg.Caption = LangClient & LangTerkunci
                Else
                    LabelMsg.Caption = LangClient & LangTerkunci
                End If
            ElseIf jenis = LangGame Then
                TxtUserGame.Text = ""
                TxtUserGame.SetFocus
                status = LangBebas
                If FrameMsg.Visible = False Then
                    FrameMsg.Visible = True
                    LabelMsg.Caption = LangClient & LangTerkunci
                Else
                    LabelMsg.Caption = LangClient & LangTerkunci
                End If
            ElseIf jenis = LangKetik Then
                TxtUserKetik.Text = ""
                TxtUserKetik.SetFocus
                status = LangBebas
                If FrameMsg.Visible = False Then
                    FrameMsg.Visible = True
                    LabelMsg.Caption = LangClient & LangTerkunci
                Else
                    LabelMsg.Caption = LangClient & LangTerkunci
                End If
            ElseIf jenis = LangMember Then
                TxtUserMember.Text = ""
                TxtUserMember.SetFocus
                TxtPasswordMember.Text = ""
                CmdStartMember.Caption = LangLogin
                status = LangBebas
                If FrameMsg.Visible = False Then
                    FrameMsg.Visible = True
                    LabelMsg.Caption = LangClient & LangTerkunci
                Else
                    LabelMsg.Caption = LangClient & LangTerkunci
                End If
            End If
        End If
    ElseIf status = "okregister" Then
        If perintah = "OK" Then
            status = "connect"
        Else
            LabelStatus.Caption = perintah
            status = "register"
        End If
    ElseIf status = LangAktif Then
        'distop
        If perintah = "stop" Then
            status = LangBebas
            Firewall = False
            MonitorKetik = False
            MonitorGame = False
            Winsock.SendData "stop|" & Winsock.LocalIP & "|" & member & "|" & jenis & "|"
            FrameBill.Visible = False
            SetFormMain
            CmdPersonal.Enabled = True
            CmdMember.Enabled = True
            CmdGame.Enabled = True
            CmdKetik.Enabled = True
            FrameLogin.Visible = True
            If CmdPersonal.Enabled = True Then
                CmdPersonal.SetFocus
            End If
        'dipindah
        ElseIf perintah = "pindahke" Then
            status = LangBebas
            Firewall = False
            MonitorKetik = False
            MonitorGame = False
            FrameBill.Visible = False
            SetFormMain
            CmdPersonal.Enabled = True
            CmdMember.Enabled = True
            CmdGame.Enabled = True
            CmdKetik.Enabled = True
            FrameLogin.Visible = True
            If CmdPersonal.Enabled = True Then
                CmdPersonal.SetFocus
            End If
        ElseIf perintah = "pesan" Then
            If frmpesanload = False Then
                Load FrmPesan
                FrmPesan.Show 1
            End If
            If FrmPesan.TxtTerimaPesan.Text = "" Then
                FrmPesan.TxtTerimaPesan.Text = "Server :" & op1
            Else
                FrmPesan.TxtTerimaPesan.Text = FrmPesan.TxtTerimaPesan.Text & Chr(13) & "Server :" & op1
            End If
        End If
    ElseIf status = LangBebas Then
        'login member
        If perintah = "password" Then
            If op2 = "ok" Then
                If op1 = member_pass Then
                    If FrameMember.Visible = True Then
                        CmdStartMember.Caption = "Start"
                    End If
                Else
                    If FrameMember.Visible = True Then
                        TxtUserMember.Text = ""
                        TxtUserMember.SetFocus
                        TxtPasswordMember.Text = ""
                    End If
                    If FrameMsg.Visible = False Then
                        FrameMsg.Visible = True
                        LabelMsg.Caption = LangPasswordSalah
                    Else
                        LabelMsg.Caption = LangPasswordSalah
                    End If
                End If
            ElseIf op2 = "no" Then
                If FrameMsg.Visible = False Then
                    FrameMsg.Visible = True
                    LabelMsg.Caption = LangDepositTidakCukup
                Else
                    LabelMsg.Caption = LangDepositTidakCukup
                End If
            End If
        'member belumterdaftar
        ElseIf perintah = LangMemberBelumTerdaftar Then
                If FrameMember.Visible = True Then
                    TxtUserMember.Text = ""
                    TxtUserMember.SetFocus
                    TxtPasswordMember.Text = ""
                End If
                If FrameMsg.Visible = False Then
                    FrameMsg.Visible = True
                    LabelMsg.Caption = LangMemberBelumTerdaftar
                Else
                    LabelMsg.Caption = LangMemberBelumTerdaftar
                End If
        'pindahan
        ElseIf perintah = "pindahdari" Then
            mulai = op1
            jenis = op2
            FrameBill.Top = -100
            FrameBill.Left = 0
            FrameLogin.Visible = False
            FrmMain.Width = FrameBill.Width
            FrmMain.Height = FrameBill.Height - 100
            FrameBill.Visible = True
            status = LangAktif
        'shutdown
        ElseIf perintah = "shutdown" Then
            Call Down
        End If
    End If
    'remote
    If perintah = "remote" Then
        If Winsock.State = sckConnected Then
            listProses.ListItems.Clear
            Call enumProc
            If listProses.ListItems.Count <> 0 Then
                remotestr1 = ""
                remotestr2 = ""
                For i = 1 To listProses.ListItems.Count
                    remotestr1 = remotestr1 & listProses.ListItems.Item(i).SubItems(5) & ":"
                    remotestr2 = remotestr2 & listProses.ListItems.Item(i).SubItems(1) & ":"
                Next i
                Winsock.SendData "remote|" & Winsock.LocalIP & "|" & remotestr1 & "|" & remotestr2 & "|"
            End If
        End If
    End If
    'remote kill
    If perintah = "kill" Then
        pid = Val(op1)
        hWndof = OpenProcess(PROCESS_TERMINATE, False, pid)
        Mati = TerminateProcess(hWndof, 0)
        CloseHandle hWndof
    End If
    Call InitAksesLogin
End Sub
Private Sub LoginPersonal()
    If Winsock.State = sckConnected Then
        FramePersonal.Visible = False
        TxtUserPersonal.Text = ""
        mulai = Now
        If Now >= mulai_disc Then
            disc_mulai = Now
        Else
            disc_mulai = mulai_disc
        End If
        durasi_disc = CDate("00:00:00")
        FrameBill.Top = -100
        FrameBill.Left = 0
        FrmMain.Width = FrameBill.Width
        FrmMain.Height = FrameBill.Height - 100
        FrameBill.Visible = True
        FrameProfile.Visible = False
        status = LangAktif
        Firewall = False
        MonitorKetik = False
        MonitorGame = False
    End If
End Sub
Private Sub LoginGame()
    If Winsock.State = sckConnected Then
        FrameGame.Visible = False
        TxtUserGame.Text = ""
        mulai = Now
        If Now >= mulai_disc Then
            disc_mulai = Now
        Else
            disc_mulai = mulai_disc
        End If
        durasi_disc = CDate("00:00:00")
        FrameBill.Top = -100
        FrameBill.Left = 0
        FrmMain.Width = FrameBill.Width
        FrmMain.Height = FrameBill.Height - 100
        FrameBill.Visible = True
        FrameProfile.Visible = False
        status = LangAktif
        Firewall = True
        MonitorKetik = True
        MonitorGame = False
    End If
End Sub
Private Sub LoginKetik()
    If Winsock.State = sckConnected Then
        FrameKetik.Visible = False
        TxtUserKetik.Text = ""
        mulai = Now
        If Now >= mulai_disc Then
            disc_mulai = Now
        Else
            disc_mulai = mulai_disc
        End If
        durasi_disc = CDate("00:00:00")
        FrameBill.Top = -100
        FrameBill.Left = 0
        FrmMain.Width = FrameBill.Width
        FrmMain.Height = FrameBill.Height - 100
        FrameBill.Visible = True
        FrameProfile.Visible = False
        status = LangAktif
        Firewall = True
        MonitorGame = True
        MonitorKetik = False
    End If
End Sub
Private Sub LoginMember()
    If CmdStartMember.Caption = LangLogin Then
        'minta password ke server
        If Winsock.State = sckConnected Then
            member = TxtUserMember.Text
            member_pass = TxtPasswordMember.Text
            jenis = LangMember
            Winsock.SendData "password|" & Winsock.LocalIP & "|" & member & "|"
        End If
    ElseIf CmdStartMember.Caption = "Start" Then
        'login sebagai member
        If Winsock.State = sckConnected Then
            Winsock.SendData "start|" & Winsock.LocalIP & "|" & member & "|" & jenis & "|"
            status = "waitstart"
        End If
    End If
End Sub

Sub StartMember()
    FrameMember.Visible = False
    TxtUserMember.Text = ""
    TxtPasswordMember.Text = ""
    CmdStartMember.Caption = LangLogin
    mulai = Now
    If Now >= mulai_disc Then
        disc_mulai = Now
    Else
        disc_mulai = mulai_disc
    End If
    durasi_disc = CDate("00:00:00")
    FrameBill.Top = -100
    FrameBill.Left = 0
    FrmMain.Width = FrameBill.Width
    FrmMain.Height = FrameBill.Height - 100
    FrameBill.Visible = True
    FrameProfile.Visible = False
    status = LangAktif
    Firewall = False
    MonitorKetik = False
    MonitorGame = False
End Sub

Private Sub TerminateList()
Dim ada As Integer
    If ListProc.ListItems.Count <> 0 Then
        For i = 1 To ListProc.ListItems.Count
            If ListTerminate.ListItems.Count > 0 Then
                If ListProc.ListItems.Item(i).SubItems(5) = "6667" Or ListProc.ListItems.Item(i).SubItems(5) = "6668" Or ListProc.ListItems.Item(i).SubItems(5) = "6669" Or ListProc.ListItems.Item(i).SubItems(5) = "7000" Or ListProc.ListItems.Item(i).SubItems(5) = "80" Or ListProc.ListItems.Item(i).SubItems(5) = "22" Then
                    ada = 0
                    For j = 1 To ListTerminate.ListItems.Count
                        If ListTerminate.ListItems.Item(j).SubItems(1) = ListProc.ListItems.Item(i).SubItems(1) Then
                            ada = ada + 1
                        End If
                    Next j
                    If ada = 0 Then
                        Set lstItem = ListTerminate.ListItems.Add(, , ListProc.ListItems.Item(i).Text)
                        lstItem.SubItems(1) = ListProc.ListItems.Item(i).SubItems(1)
                        lstItem.SubItems(2) = ListProc.ListItems.Item(i).SubItems(2)
                        lstItem.SubItems(3) = ListProc.ListItems.Item(i).SubItems(3)
                        lstItem.SubItems(4) = ListProc.ListItems.Item(i).SubItems(4)
                        lstItem.SubItems(5) = ListProc.ListItems.Item(i).SubItems(5)
                        lstItem.SubItems(6) = ListProc.ListItems.Item(i).SubItems(6)
                    End If
                End If
            Else
                If ListProc.ListItems.Item(i).SubItems(5) = "6667" Or ListProc.ListItems.Item(i).SubItems(5) = "6668" Or ListProc.ListItems.Item(i).SubItems(5) = "6669" Or ListProc.ListItems.Item(i).SubItems(5) = "7000" Or ListProc.ListItems.Item(i).SubItems(5) = "80" Or ListProc.ListItems.Item(i).SubItems(5) = "22" Then
                    Set lstItem = ListTerminate.ListItems.Add(, , ListProc.ListItems.Item(i).Text)
                    lstItem.SubItems(1) = ListProc.ListItems.Item(i).SubItems(1)
                    lstItem.SubItems(2) = ListProc.ListItems.Item(i).SubItems(2)
                    lstItem.SubItems(3) = ListProc.ListItems.Item(i).SubItems(3)
                    lstItem.SubItems(4) = ListProc.ListItems.Item(i).SubItems(4)
                    lstItem.SubItems(5) = ListProc.ListItems.Item(i).SubItems(5)
                    lstItem.SubItems(6) = ListProc.ListItems.Item(i).SubItems(6)
                End If
            End If
        Next i
    End If
End Sub
Private Sub TerminateProc()
Dim Mati As Long
Dim pid As Long
    If ListTerminate.ListItems.Count <> 0 Then
        For i = 1 To ListTerminate.ListItems.Count
            pid = Val(ListTerminate.ListItems.Item(i).SubItems(1))
            hWndof = OpenProcess(PROCESS_TERMINATE, False, pid)
            Mati = TerminateProcess(hWndof, 0)
            CloseHandle hWndof
        Next i
    End If
End Sub

Private Sub ListKetik()
Dim ada As Integer
    If listProses.ListItems.Count <> 0 Then
        For i = 1 To listProses.ListItems.Count
            If ProsesMati.ListItems.Count > 0 Then
                If listProses.ListItems.Item(i).SubItems(5) = ketik1 Or listProses.ListItems.Item(i).SubItems(5) = ketik2 Or listProses.ListItems.Item(i).SubItems(5) = ketik3 Then
                    ada = 0
                    For j = 1 To ProsesMati.ListItems.Count
                        If ProsesMati.ListItems.Item(j).SubItems(1) = listProses.ListItems.Item(i).SubItems(1) Then
                            ada = ada + 1
                        End If
                    Next j
                    If ada = 0 Then
                        Set lstItem = ProsesMati.ListItems.Add(, , listProses.ListItems.Item(i).Text)
                        lstItem.SubItems(1) = listProses.ListItems.Item(i).SubItems(1)
                        lstItem.SubItems(2) = listProses.ListItems.Item(i).SubItems(2)
                        lstItem.SubItems(3) = listProses.ListItems.Item(i).SubItems(3)
                        lstItem.SubItems(4) = listProses.ListItems.Item(i).SubItems(4)
                        lstItem.SubItems(5) = listProses.ListItems.Item(i).SubItems(5)
                    End If
                End If
            Else
                If listProses.ListItems.Item(i).SubItems(5) = ketik1 Or listProses.ListItems.Item(i).SubItems(5) = ketik2 Or listProses.ListItems.Item(i).SubItems(5) = ketik3 Then
                    Set lstItem = ProsesMati.ListItems.Add(, , listProses.ListItems.Item(i).Text)
                    lstItem.SubItems(1) = listProses.ListItems.Item(i).SubItems(1)
                    lstItem.SubItems(2) = listProses.ListItems.Item(i).SubItems(2)
                    lstItem.SubItems(3) = listProses.ListItems.Item(i).SubItems(3)
                    lstItem.SubItems(4) = listProses.ListItems.Item(i).SubItems(4)
                    lstItem.SubItems(5) = listProses.ListItems.Item(i).SubItems(5)
                End If
            End If
        Next i
    End If
End Sub

Private Sub MatikanKetik()
Dim Mati As Long
Dim pid As Long
    If ProsesMati.ListItems.Count <> 0 Then
        For i = 1 To ProsesMati.ListItems.Count
            pid = Val(ProsesMati.ListItems.Item(i).SubItems(1))
            hWndof = OpenProcess(PROCESS_TERMINATE, False, pid)
            Mati = TerminateProcess(hWndof, 0)
            CloseHandle hWndof
        Next i
    End If
End Sub

Private Sub ListGame()
Dim ada As Integer
    If listProses.ListItems.Count <> 0 Then
        For i = 1 To listProses.ListItems.Count
            If ProsesMati.ListItems.Count > 0 Then
                If listProses.ListItems.Item(i).SubItems(5) = game1 Or listProses.ListItems.Item(i).SubItems(5) = game2 Or listProses.ListItems.Item(i).SubItems(5) = game3 Then
                    ada = 0
                    For j = 1 To ProsesMati.ListItems.Count
                        If ProsesMati.ListItems.Item(j).SubItems(1) = listProses.ListItems.Item(i).SubItems(1) Then
                            ada = ada + 1
                        End If
                    Next j
                    If ada = 0 Then
                        Set lstItem = ProsesMati.ListItems.Add(, , listProses.ListItems.Item(i).Text)
                        lstItem.SubItems(1) = listProses.ListItems.Item(i).SubItems(1)
                        lstItem.SubItems(2) = listProses.ListItems.Item(i).SubItems(2)
                        lstItem.SubItems(3) = listProses.ListItems.Item(i).SubItems(3)
                        lstItem.SubItems(4) = listProses.ListItems.Item(i).SubItems(4)
                        lstItem.SubItems(5) = listProses.ListItems.Item(i).SubItems(5)
                    End If
                End If
            Else
                If listProses.ListItems.Item(i).SubItems(5) = game1 Or listProses.ListItems.Item(i).SubItems(5) = game2 Or listProses.ListItems.Item(i).SubItems(5) = game3 Then
                    Set lstItem = ProsesMati.ListItems.Add(, , listProses.ListItems.Item(i).Text)
                    lstItem.SubItems(1) = listProses.ListItems.Item(i).SubItems(1)
                    lstItem.SubItems(2) = listProses.ListItems.Item(i).SubItems(2)
                    lstItem.SubItems(3) = listProses.ListItems.Item(i).SubItems(3)
                    lstItem.SubItems(4) = listProses.ListItems.Item(i).SubItems(4)
                    lstItem.SubItems(5) = listProses.ListItems.Item(i).SubItems(5)
                End If
            End If
        Next i
    End If
End Sub

Private Sub MatikanGame()
Dim Mati As Long
Dim pid As Long
    If ProsesMati.ListItems.Count <> 0 Then
        For i = 1 To ProsesMati.ListItems.Count
            pid = Val(ProsesMati.ListItems.Item(i).SubItems(1))
            hWndof = OpenProcess(PROCESS_TERMINATE, False, pid)
            Mati = TerminateProcess(hWndof, 0)
            CloseHandle hWndof
        Next i
    End If
End Sub

Private Sub InitAksesLogin()
    If lpersonal = "1" Then
        If Winsock.State = sckConnected Then
            CmdPersonal.Enabled = True
            optpersonal.value = 1
        Else
            CmdPersonal.Enabled = False
            optpersonal.value = 1
        End If
    Else
        CmdPersonal.Enabled = False
        optpersonal.value = 0
    End If
    If lmember = "1" Then
        If Winsock.State = sckConnected Then
            CmdMember.Enabled = True
            optmember.value = 1
        Else
            CmdMember.Enabled = False
            optmember.value = 1
        End If
    Else
        CmdMember.Enabled = False
        optmember.value = 0
    End If
    If lgame = "1" Then
        If Winsock.State = sckConnected Then
            CmdGame.Enabled = True
            optgame.value = 1
        Else
            CmdGame.Enabled = False
            optgame.value = 1
        End If
    Else
        CmdGame.Enabled = False
        optgame.value = 0
    End If
    If lketik = "1" Then
        If Winsock.State = sckConnected Then
            CmdKetik.Enabled = True
            optketik.value = 1
        Else
            CmdKetik.Enabled = False
            optketik.value = 1
        End If
    Else
        CmdKetik.Enabled = False
        optketik.value = 0
    End If
End Sub
