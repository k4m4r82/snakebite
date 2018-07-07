VERSION 5.00
Begin VB.Form FrmStopClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Stop Client"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2325
   Icon            =   "FrmStopClient.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   2325
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdBatal 
      Caption         =   "Batal"
      Height          =   375
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton CmdStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox TxtNo_Client 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   0
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Nomor Client"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "FrmStopClient"
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

Private Sub CmdStop_Click()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    If FrmMain.MSFlexGrid.TextMatrix(Val(TxtNo_Client.Text), 3) = LangAktif Then
        cn.Open cnstr
        rs.ActiveConnection = cn
        rs.Open "SELECT ip.*,socket.socket FROM ip,socket WHERE ip.ip = socket.ip AND ip.no_client = " & Val(TxtNo_Client.Text), cn, adOpenStatic, adLockOptimistic
        If Not rs.RecordCount = 0 Then
            If FrmMain.Winsock(rs(2)).State = sckConnected Then
                FrmMain.Winsock(rs(2)).SendData "stop|_|_|_|_|_|_|"
            Else
                FrmMain.MSFlexGrid.Col = 3
                FrmMain.MSFlexGrid.Row = Val(TxtNo_Client.Text)
                FrmMain.MSFlexGrid.CellBackColor = &HFFC0C0 'ungu
                FrmMain.MSFlexGrid.TextMatrix(Val(TxtNo_Client.Text), 3) = "stop"
            End If
        Else
            FrmMain.MSFlexGrid.Col = 3
            FrmMain.MSFlexGrid.Row = Val(TxtNo_Client.Text)
            FrmMain.MSFlexGrid.CellBackColor = &HFFC0C0 'ungu
            FrmMain.MSFlexGrid.TextMatrix(Val(TxtNo_Client.Text), 3) = "stop"
        End If
        cn.Close
    End If
    Unload Me
End Sub

Private Sub CmdStop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdStop, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub Form_Load()
    Label1.Caption = LangNomor & " " & LangClient
    CmdStop.Caption = LangStop
    CmdBatal.Caption = LangBatal
    CmdStop.ToolTipText = LangStop
    CmdBatal.ToolTipText = LangBatal
    FrmStopClient.Caption = LangStop & " " & LangClient
End Sub

Private Sub TxtNo_Client_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 13 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        CmdStop.SetFocus
    End If
End Sub
