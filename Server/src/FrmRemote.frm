VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmRemote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remote Client"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4455
   Icon            =   "FrmRemote.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   4455
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdKill 
      Caption         =   "Kill"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "Tutup"
      Height          =   375
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3495
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   6165
      _Version        =   393216
      FixedCols       =   0
      Appearance      =   0
   End
   Begin VB.CommandButton CmdRemote 
      Caption         =   "Remote"
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "Batal"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox TxtNo_Client 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1800
      MaxLength       =   2
      TabIndex        =   0
      Top             =   240
      Width           =   375
   End
   Begin VB.Label LabelTotal 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   4560
      Width           =   1815
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
Attribute VB_Name = "FrmRemote"
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

Private Sub CmdKill_Click()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open "SELECT ip.*,socket.socket FROM ip,socket WHERE ip.ip = socket.ip AND ip.no_client = " & Val(TxtNo_Client.Text), cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        If FrmMain.Winsock(rs(2)).State = sckConnected Then
            If Grid.Rows > 1 Then
                FrmMain.Winsock(rs(2)).SendData "kill|" & Grid.TextMatrix(Grid.Row, 1) & "|_|_|_|_|_|"
            End If
        End If
    End If
    cn.Close
    CmdRemote_Click
End Sub

Private Sub CmdKill_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdKill, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdRemote_Click()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open "SELECT ip.*,socket.socket FROM ip,socket WHERE ip.ip = socket.ip AND ip.no_client = " & Val(TxtNo_Client.Text), cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        If FrmMain.Winsock(rs(2)).State = sckConnected Then
            FrmMain.Winsock(rs(2)).SendData "remote|_|_|_|_|_|_|"
        End If
    End If
    cn.Close
End Sub


Private Sub CmdRemote_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdRemote, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdTutup_Click()
    Unload Me
End Sub

Private Sub CmdTutup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdTutup, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub Form_Load()
    Me.Caption = LangmnuRemote
    Label1.Caption = LangNo & " " & LangClient
    CmdBatal.Caption = LangBatal
    CmdTutup.Caption = LangTutup
End Sub
