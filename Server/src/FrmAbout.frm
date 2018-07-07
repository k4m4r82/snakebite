VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "FrmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdTutup 
      Caption         =   "Tutup"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3360
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   120
      Picture         =   "FrmAbout.frx":030A
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "website : http://snakebite.sourceforge.net"
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   2880
      Width           =   3135
   End
   Begin VB.Label Label6 
      Caption         =   "mobile : +62 081344823917"
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Label Label5 
      Caption         =   "phone : +62 0967 587761"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "email : snakebites@users.sourceforge.net"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "contact info :"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "programmed by setia budi (snake)"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "snakebite cybercafe billing system"
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdTutup_Click()
    Unload Me
End Sub

Private Sub CmdTutup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdTutup, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub Form_Load()
    FrmAbout.Caption = "About " & NamaAplikasi
    CmdTutup.Caption = LangTutup
End Sub
