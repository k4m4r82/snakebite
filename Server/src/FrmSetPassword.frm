VERSION 5.00
Begin VB.Form FrmSetPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Password Sysadmin"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   Icon            =   "FrmSetPassword.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   4425
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox TxtBahasa 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "FrmSetPassword.frx":030A
      Left            =   2040
      List            =   "FrmSetPassword.frx":0314
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox TxtUser 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton CmdOK 
      Appearance      =   0  'Flat
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox TxtPasswordLagi 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox TxtPassword 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Bahasa"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Password Lagi"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Nama User"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "FrmSetPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdOK_Click()
    Call DumpSysadmin
End Sub

Private Sub CmdOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdOK, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub Form_Activate()
    TxtUser.Text = "sysadmin"
    TxtBahasa.ListIndex = 0
End Sub

Private Sub Form_Load()
    Label1.Caption = LangNamaUser
    Label2.Caption = LangPassword
    Label3.Caption = LangPasswordLagi
    Label4.Caption = LangBahasa
    CmdOK.Caption = LangUpdate
End Sub

Private Sub TxtBahasa_Click()
    If TxtBahasa.Text = "Indonesia" Then
        lang = "id"
    Else
        lang = "en"
    End If
    Call Bahasa
    Label1.Caption = LangNamaUser
    Label2.Caption = LangPassword
    Label3.Caption = LangPasswordLagi
    Label4.Caption = LangBahasa
    CmdOK.Caption = LangUpdate
End Sub

Private Sub TxtBahasa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CmdOK.Enabled = True Then
            CmdOK.SetFocus
        End If
    End If
End Sub

Private Sub TxtPassword_Change()
    If Not (TxtPassword.Text = "" Or TxtPasswordLagi.Text = "") Then
        CmdOK.Enabled = True
    Else
        CmdOK.Enabled = False
    End If
End Sub

Private Sub TxtPassword_GotFocus()
    TxtPassword.Backcolor = &HFFC0C0
End Sub

Private Sub TxtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TxtPassword.Text <> "" Then
            TxtPasswordLagi.SetFocus
        End If
    End If
End Sub

Private Sub DumpSysadmin()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim passwd As String
    If TxtPassword.Text = TxtPasswordLagi.Text Then
        passwd = EncryptText(TxtPassword.Text, "password")
        cn.Open cnstr
        cmd.ActiveConnection = cn
        cmd.CommandType = adCmdText
        cmd.CommandText = "INSERT INTO account([user],[nama],[login],[alamat],[telp],[password]) VALUES ('sysadmin','system administrator','sysadmin','Abepura','587761','" & passwd & "')"
        cmd.Execute
        If TxtBahasa.Text = "Indonesia" Then
            Call RegSaveString(HKEY_LOCAL_MACHINE, "Software\snakebite\server\" & VersiAplikasi & "\", "lang", "id")
        ElseIf TxtBahasa.Text = "English" Then
            Call RegSaveString(HKEY_LOCAL_MACHINE, "Software\snakebite\server\" & VersiAplikasi & "\", "lang", "en")
        Else
            Call RegSaveString(HKEY_LOCAL_MACHINE, "Software\snakebite\server\" & VersiAplikasi & "\", "lang", "id")
        End If
        MsgBox LangSetPasswordSukses, vbInformation, LangInformasi
        Unload Me
        End
    Else
        MsgBox LangPasswordSalah, vbInformation, LangInformasi
        TxtPassword.Text = ""
        TxtPasswordLagi.Text = ""
        TxtPassword.SetFocus
    End If
End Sub

Private Sub TxtPassword_LostFocus()
    TxtPassword.Backcolor = &H80000005
End Sub

Private Sub TxtPasswordLagi_Change()
    If Not (TxtPassword.Text = "" Or TxtPasswordLagi.Text = "") Then
        CmdOK.Enabled = True
    Else
        CmdOK.Enabled = False
    End If
End Sub

Private Sub TxtPasswordLagi_GotFocus()
    TxtPasswordLagi.Backcolor = &HFFC0C0
End Sub

Private Sub TxtPasswordLagi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not TxtPasswordLagi.Text = "" Then
            TxtBahasa.SetFocus
        End If
    End If
End Sub

Private Sub TxtPasswordLagi_LostFocus()
    TxtPasswordLagi.Backcolor = &H80000005
End Sub
