VERSION 5.00
Begin VB.Form FrmLogOut 
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   2775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4095
   Icon            =   "FrmLogout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2775
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdBatal 
      Appearance      =   0  'Flat
      Caption         =   "&Batal"
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton CmdLogout 
      Appearance      =   0  'Flat
      Caption         =   "&Logout"
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox TxtPassword 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1680
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox TxtUser 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
   Begin VB.Shape Shape2 
      Height          =   1935
      Left            =   0
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Logout..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1200
      TabIndex        =   6
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "User Name"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   80
      Picture         =   "FrmLogout.frx":030A
      Top             =   80
      Width           =   720
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "FrmLogOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBatal_Click()
    Unload Me
    Load FrmMain
    FrmMain.Show
End Sub

Private Sub CmdBatal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdBatal, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdLogout_Click()
    Call LogOut
End Sub

Private Sub CmdLogout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdLogout, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub Form_Load()
    TxtUser.Text = username
    Label1.Caption = LangNamaUser
    Label2.Caption = LangPassword
    CmdLogout.Caption = LangGlobLogout
    CmdBatal.Caption = LangBatal
    CmdLogout.ToolTipText = LangGlobLogout
    CmdBatal.ToolTipText = LangBatal
End Sub

Private Sub TxtPassword_GotFocus()
    TxtPassword.Backcolor = &HFFC0C0
End Sub

Private Sub TxtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call LogOut
    End If
End Sub

Private Sub TxtPassword_LostFocus()
    TxtPassword.Backcolor = &H80000005
End Sub

Private Sub TxtUser_GotFocus()
    TxtUser.Backcolor = &HFFC0C0
End Sub

Private Sub TxtUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call CekUser
    End If
End Sub

Private Sub CekUser()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open "SELECT cache.* FROM cache WHERE cache.path = 'server'", cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        If TxtUser.Text = rs(2) Then
            TxtPassword.SetFocus
        Else
            MsgBox LangUserSalah & "!", vbInformation, LangInformasi
            TxtUser.Text = ""
            TxtUser.SetFocus
        End If
    Else
        MsgBox LangGagalCache, vbCritical, LangInformasi
        End
    End If
    cn.Close
End Sub

Private Sub LogOut()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open "SELECT account.* FROM account WHERE account.user = '" & TxtUser.Text & "'", cn, adOpenStatic, adLockOptimistic
    If rs.RecordCount = 0 Then
        MsgBox LangNamaUser & " '" & TxtUser.Text & "' " & LangTidakDitemukan & ".", vbInformation, LangInformasi
        TxtUser.Text = ""
        TxtPassword.Text = ""
        TxtUser.SetFocus
    Else
        If TxtPassword.Text = DecryptText(rs(5), "password") Then
            Unload Me
            End
        Else
            MsgBox LangPasswordSalah & "!", vbInformation, LangInformasi
            TxtPassword.Text = ""
            TxtPassword.SetFocus
        End If
    End If
    cn.Close
End Sub

Private Sub TxtUser_LostFocus()
    TxtUser.Backcolor = &H80000005
End Sub
