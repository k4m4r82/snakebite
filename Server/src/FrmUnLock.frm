VERSION 5.00
Begin VB.Form FrmUnLock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buka Client"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3765
   Icon            =   "FrmUnLock.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   3765
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdBatal 
      Caption         =   "Batal"
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton CmdUnLock 
      Caption         =   "Buka"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.OptionButton OpSemua 
      Caption         =   "Semua Client"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox TxtNo_Client 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1440
      MaxLength       =   2
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Nomor Client"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "FrmUnLock"
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

Private Sub CmdUnLock_Click()
    If OpSemua.value = True Then
        For i = 1 To ClientMax
            kunci(i) = False
        Next i
        FrmMain.StatusBar.Panels(9).Text = LangBuka
    Else
        kunci(Val(TxtNo_Client.Text)) = False
    End If
    Unload Me
End Sub

Private Sub CmdUnLock_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdUnLock, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub Form_Load()
    Me.Caption = LangmnuBuka
    Label1.Caption = LangNomor & " " & LangClient
    OpSemua.Caption = LangSemua & " " & LangClient
    CmdUnLock.Caption = LangBuka
    CmdBatal.Caption = LangBatal
    OpSemua.value = False
End Sub

Private Sub TxtNo_Client_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 13 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        CmdUnLock.SetFocus
    End If
End Sub
