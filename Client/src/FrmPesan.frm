VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmPesan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kirim Pesan"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   Icon            =   "FrmPesan.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6015
   Begin VB.CommandButton CmdKirim 
      Caption         =   "Kirim"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox TxtKirimPesan 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   4695
   End
   Begin RichTextLib.RichTextBox TxtTerimaPesan 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5530
      _Version        =   393217
      Enabled         =   0   'False
      Appearance      =   0
      TextRTF         =   $"FrmPesan.frx":030A
   End
End
Attribute VB_Name = "FrmPesan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdKirim_Click()
    If FrmMain.Winsock.State = sckConnected Then
        FrmMain.Winsock.SendData "pesan|" & FrmMain.Winsock.LocalIP & "|" & TxtKirimPesan.Text & "|_|_|"
        TxtKirimPesan.Text = ""
    End If
End Sub

Private Sub Form_Activate()
    FrmPesan.TxtKirimPesan.SetFocus
End Sub

Private Sub Form_Load()
    frmpesanload = True
    CmdKirim.Caption = LangKirim
    Me.Caption = LangKirimPesan
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmpesanload = False
End Sub

Private Sub TxtKirimPesan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdKirim.SetFocus
    End If
End Sub
