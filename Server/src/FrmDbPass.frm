VERSION 5.00
Begin VB.Form FrmDbPass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ganti Password Database"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4170
   Icon            =   "FrmDbPass.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   4170
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdBatal 
      Appearance      =   0  'Flat
      Caption         =   "&Batal"
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton CmdOK 
      Appearance      =   0  'Flat
      Caption         =   "&OK"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox TxtPassLagi 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox TxtPass 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Password Lagi"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Password"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "FrmDbPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBatal_Click()
    Unload Me
End Sub

Private Sub CmdBatal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdBatal, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdOK_Click()
    If TxtPass.Text = TxtPassLagi.Text Then
        GantiDbPass App.path & "\" & DataFileName, dbpass, TxtPass.Text
        dbpass = TxtPass.Text
        Call Main
        Call RegSaveString(HKEY_LOCAL_MACHINE, "Software\snakebite\server\" & VersiAplikasi & "\", "dbpass", EncryptText(dbpass, "password"))
        Unload Me
    Else
        MsgBox LangPasswordSalah, vbInformation, LangInformasi
        TxtPass.Text = ""
        TxtPassLagi.Text = ""
        TxtPass.SetFocus
    End If
End Sub

Private Sub CmdOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdOK, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub Form_Load()
    Me.Caption = LangmnuGantiPass
    Label1.Caption = LangPassword
    Label2.Caption = LangPasswordLagi
    CmdOK.Caption = LangUpdate
    CmdBatal.Caption = LangBatal
End Sub

Private Sub TxtPass_GotFocus()
    TxtPass.Backcolor = &HFFC0C0
End Sub

Private Sub TxtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not TxtPass.Text = "" Then
            TxtPassLagi.SetFocus
        End If
    End If
End Sub

Private Sub TxtPass_LostFocus()
    TxtPass.Backcolor = &H80000005
End Sub

Private Sub TxtPassLagi_GotFocus()
    TxtPassLagi.Backcolor = &HFFC0C0
End Sub

Private Sub TxtPassLagi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not TxtPassLagi.Text = "" Then
            CmdOK.SetFocus
        End If
    End If
End Sub

Private Sub TxtPassLagi_LostFocus()
    TxtPassLagi.Backcolor = &H80000005
End Sub
