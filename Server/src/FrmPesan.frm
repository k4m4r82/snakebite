VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmPesan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kirim dan Terima Pesan"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7215
   Icon            =   "FrmPesan.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7215
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdTutup 
      Caption         =   "Tutup"
      Height          =   375
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton CmdHapusLog 
      Caption         =   "Hapus Log"
      Height          =   375
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton CmdKirim 
      Caption         =   "Kirim"
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   855
   End
   Begin VB.ComboBox TxtNoClient 
      Height          =   315
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox TxtPesanDiKirim 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   4335
   End
   Begin RichTextLib.RichTextBox TxtPesanDiterima 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   5318
      _Version        =   393217
      Enabled         =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"FrmPesan.frx":030A
   End
   Begin VB.Label Label1 
      Caption         =   "Client"
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   3360
      Width           =   495
   End
End
Attribute VB_Name = "FrmPesan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ClientSock(1 To ClientMax) As Integer

Private Sub CmdHapusLog_Click()
    TxtPesanDiterima.Text = ""
    TxtPesanDiterima.SaveFile (App.path & "\" & MsgFilename)
End Sub

Private Sub CmdHapusLog_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdHapusLog, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdKirim_Click()
    If FrmMain.Winsock(ClientSock(Val(TxtNoClient.Text))).State = sckConnected Then
        FrmMain.Winsock(ClientSock(Val(TxtNoClient.Text))).SendData "pesan|" & TxtPesanDiKirim.Text
        TxtPesanDiKirim.Text = ""
    End If
End Sub

Private Sub CmdKirim_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdKirim, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdTutup_Click()
    Unload Me
End Sub

Private Sub CmdTutup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdTutup, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub Form_Activate()
    TxtPesanDiKirim.SetFocus
    Call CekClientAktif
End Sub

Private Sub Form_Load()
    FrmPesanLoad = True
    FrmPesan.Caption = LangKirimPesan
    CmdKirim.Caption = LangKirim
    CmdHapusLog.Caption = LangHapusLog
    CmdTutup.Caption = LangTutup
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FrmPesanLoad = False
End Sub
Private Sub CekClientAktif()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    Dim sql As String
    sql = "SELECT socket.socket,ip.no_client FROM socket,ip WHERE socket.ip=ip.ip"
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open sql, cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        rs.MoveFirst
        For i = 1 To rs.RecordCount
            ClientSock(rs(1)) = rs(0)
            TxtNoClient.AddItem rs(1)
            rs.MoveNext
        Next i
        TxtNoClient.ListIndex = 0
        CmdKirim.Enabled = True
    Else
        CmdKirim.Enabled = False
    End If
    cn.Close
End Sub

Private Sub TxtPesanDiKirim_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CmdKirim.Enabled = True Then
            CmdKirim.SetFocus
        End If
    End If
End Sub
