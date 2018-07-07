VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form FrmLoading 
   BorderStyle     =   0  'None
   Caption         =   "Loading..."
   ClientHeight    =   6285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8640
   Icon            =   "FrmLoading.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "FrmLoading.frx":030A
   ScaleHeight     =   6285
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5400
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   661
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
      Max             =   5000
   End
   Begin VB.Label TxtLoad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "version 3.5.6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   2
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "snakebite cybercafe billing system"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   720
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "snakeserver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
End
Attribute VB_Name = "FrmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Call SetPassword
    'sndPlaySound App.path & "\Sound\ak47.wav", SND_ASYNC
End Sub

Private Sub Form_Load()
    Call CekKey
    Call Bahasa
    If FExists(App.path & "\" & DataFileName) = True Then
        Call Main
    Else
        BuatDatabase (App.path & "\" & DataFileName)
        BuatTabel (App.path & "\" & DataFileName)
        SetDbPassword (App.path & "\" & DataFileName)
        Call Main
    End If
    PBcolor ProgressBar, vbWhite, vbBlue
    FrmLoading.Caption = NamaAplikasi & " " & VersiAplikasi
    Label3.Caption = LangVersi & " " & VersiAplikasi
    If App.PrevInstance = True Then
        End
    End If
End Sub

Private Sub SetPassword()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open "SELECT account.* FROM account", cn, adOpenStatic, adLockOptimistic
    If rs.RecordCount = 0 Then
        MsgBox "First Time Runing Program. Dumping database...", vbInformation, "First Runing Program..."
        Load FrmSetPassword
        FrmSetPassword.Show
        Unload Me
    Else
        Call Animate
    End If
    cn.Close
End Sub
Private Sub Animate()
    For i = 1 To 5000
        ProgressBar.value = i
        TxtLoad.Refresh
    Next i
    Unload Me
    Load FrmLogin
    FrmLogin.Show
End Sub
