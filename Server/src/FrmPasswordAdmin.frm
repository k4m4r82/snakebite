VERSION 5.00
Begin VB.Form FrmPasswordAdmin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Password Admin Client"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   Icon            =   "FrmPasswordAdmin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4395
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdBatal 
      Caption         =   "Batal"
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton CmdSet 
      Caption         =   "Update"
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox TxtPassword1 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2280
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox TxtPassword 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2280
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Ulang Password"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Password"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "FrmPasswordAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SetPassword()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open "SELECT admin.* FROM admin", cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        rs.MoveFirst
        rs(0) = EncryptText(TxtPassword.Text, "password")
        rs.Update
    Else
        rs.AddNew
        rs(0) = EncryptText(TxtPassword.Text, "password")
        rs.Update
    End If
    cn.Close
    MsgBox LangUpdateSukses, vbInformation, LangInformasi
    password_admin = TxtPassword.Text
    Unload Me
End Sub

Private Sub CmdBatal_Click()
    Unload Me
End Sub

Private Sub CmdBatal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdBatal, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdSet_Click()
    If TxtPassword.Text <> "" Then
        If TxtPassword = TxtPassword1 Then
            Call SetPassword
        Else
            TxtPassword.Text = ""
            TxtPassword1.Text = ""
            MsgBox LangPasswordSalah, vbInformation, LangInformasi
            TxtPassword.SetFocus
        End If
    Else
        MsgBox LangPasswordSalah, vbInformation, LangInformasi
        TxtPassword.SetFocus
    End If
End Sub

Private Sub CmdSet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdSet, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub Form_Load()
    FrmPasswordAdmin.Caption = LangSetAdminClient
    Label1.Caption = LangPassword
    Label2.Caption = LangUlangPassword
    CmdSet.Caption = LangUpdate
    CmdBatal.Caption = LangBatal
End Sub

Private Sub TxtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtPassword1.SetFocus
    End If
End Sub

Private Sub TxtPassword1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdSet.SetFocus
    End If
End Sub
