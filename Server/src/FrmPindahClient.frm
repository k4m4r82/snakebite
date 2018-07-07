VERSION 5.00
Begin VB.Form FrmPindahClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Pindahkan Client"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3600
   Icon            =   "FrmPindahClient.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   3600
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdBatal 
      Caption         =   "Batal"
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton CmdPindah 
      Caption         =   "Pindah"
      Height          =   375
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox TxtNoKe 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2880
      MaxLength       =   2
      TabIndex        =   1
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox TxtNoDari 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1200
      MaxLength       =   2
      TabIndex        =   0
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Ke Client"
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Dari Client"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "FrmPindahClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sockdari As Long
Public sockke As Long
Public mulai_dari As Date
Public jenis_dari As String
Public user_dari As String

Private Sub CmdBatal_Click()
    Unload Me
End Sub

Private Sub CmdBatal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdBatal, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdPindah_Click()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    If FrmMain.MSFlexGrid.TextMatrix(Val(TxtNoDari.Text), 3) = LangAktif And FrmMain.MSFlexGrid.TextMatrix(Val(TxtNoKe.Text), 3) = LangBebas Then
        'cek open sock client dipindah
        cn.Open cnstr
        rs.ActiveConnection = cn
        rs.Open "SELECT ip.*,socket.socket FROM ip,socket WHERE ip.ip = socket.ip AND ip.no_client = " & Val(TxtNoDari.Text), cn, adOpenStatic, adLockOptimistic
        If Not rs.RecordCount = 0 Then
            sockdari = rs(2)
        Else
            sockdari = 999
        End If
        cn.Close
        'cek open sock client kepindah
        cn.Open cnstr
        rs.ActiveConnection = cn
        rs.Open "SELECT ip.*,socket.socket FROM ip,socket WHERE ip.ip = socket.ip AND ip.no_client = " & Val(TxtNoKe.Text), cn, adOpenStatic, adLockOptimistic
        If Not rs.RecordCount = 0 Then
            sockke = rs(2)
        Else
            sockke = 999
        End If
        cn.Close
        'ambil mulai dan jenis
        cn.Open cnstr
        rs.ActiveConnection = cn
        rs.Open "SELECT proses.* FROM proses WHERE proses.no_client = " & Val(TxtNoDari.Text), cn, adOpenStatic, adLockOptimistic
        'buat terpindah jadi bebas
        If Not rs.RecordCount = 0 Then
            mulai_dari = rs(4)
            jenis_dari = rs(2)
            user_dari = rs(3)
            rs(1) = LangBebas
            rs.Update
        End If
        cn.Close
        'tulis proses dalam db
        cn.Open cnstr
        rs.ActiveConnection = cn
        rs.Open "SELECT proses.* FROM proses WHERE proses.no_client = " & Val(TxtNoKe.Text), cn, adOpenStatic, adLockOptimistic
        If Not rs.RecordCount = 0 Then
            rs(0) = Val(TxtNoKe.Text)
            rs(4) = mulai_dari
            rs(2) = jenis_dari
            rs(3) = user_dari
            rs(1) = LangAktif
            rs.Update
        Else
            rs.AddNew
            rs(0) = Val(TxtNoKe.Text)
            rs(4) = mulai_dari
            rs(2) = jenis_dari
            rs(3) = user_dari
            rs(1) = LangAktif
            rs.Update
        End If
        cn.Close
        'ubah grid pindahdari
        FrmMain.MSFlexGrid.TextMatrix(Val(TxtNoKe.Text), 2) = FrmMain.MSFlexGrid.TextMatrix(Val(TxtNoDari.Text), 2)
        FrmMain.MSFlexGrid.Col = 3
        FrmMain.MSFlexGrid.Row = Val(TxtNoKe)
        FrmMain.MSFlexGrid.CellBackColor = &HC0FFC0 'hijau
        FrmMain.MSFlexGrid.TextMatrix(Val(TxtNoKe.Text), 3) = LangPindah
        FrmMain.MSFlexGrid.TextMatrix(Val(TxtNoKe.Text), 4) = FrmMain.MSFlexGrid.TextMatrix(Val(TxtNoDari.Text), 4)
        FrmMain.MSFlexGrid.TextMatrix(Val(TxtNoKe.Text), 5) = FrmMain.MSFlexGrid.TextMatrix(Val(TxtNoDari.Text), 5)
        FrmMain.MSFlexGrid.TextMatrix(Val(TxtNoKe.Text), 6) = FrmMain.MSFlexGrid.TextMatrix(Val(TxtNoDari.Text), 6)
        FrmMain.MSFlexGrid.TextMatrix(Val(TxtNoKe.Text), 7) = FrmMain.MSFlexGrid.TextMatrix(Val(TxtNoDari.Text), 7)
        FrmMain.MSFlexGrid.TextMatrix(Val(TxtNoKe.Text), 8) = FrmMain.MSFlexGrid.TextMatrix(Val(TxtNoDari.Text), 8)
        FrmMain.MSFlexGrid.TextMatrix(Val(TxtNoKe.Text), 9) = FrmMain.MSFlexGrid.TextMatrix(Val(TxtNoDari.Text), 9)
        FrmMain.MSFlexGrid.TextMatrix(Val(TxtNoKe.Text), 10) = FrmMain.MSFlexGrid.TextMatrix(Val(TxtNoDari.Text), 10)
        'ubah grid pindahke
        FrmMain.MSFlexGrid.Col = 3
        FrmMain.MSFlexGrid.Row = Val(TxtNoDari)
        FrmMain.MSFlexGrid.CellBackColor = &HFFC0C0 ' ungu
        FrmMain.MSFlexGrid.TextMatrix(Val(TxtNoDari.Text), 3) = LangBebas
        FrmMain.MSFlexGrid.TextMatrix(Val(TxtNoDari.Text), 2) = ""
        FrmMain.MSFlexGrid.TextMatrix(Val(TxtNoDari.Text), 4) = ""
        FrmMain.MSFlexGrid.TextMatrix(Val(TxtNoDari.Text), 5) = ""
        FrmMain.MSFlexGrid.TextMatrix(Val(TxtNoDari.Text), 6) = ""
        FrmMain.MSFlexGrid.TextMatrix(Val(TxtNoDari.Text), 7) = ""
        FrmMain.MSFlexGrid.TextMatrix(Val(TxtNoDari.Text), 8) = ""
        FrmMain.MSFlexGrid.TextMatrix(Val(TxtNoDari.Text), 9) = ""
        FrmMain.MSFlexGrid.TextMatrix(Val(TxtNoDari.Text), 10) = ""
        If Not sockdari = 999 Then
            Call Menuju
        End If
        If Not sockke = 999 Then
            Call Tertuju
        End If
    End If
    Unload Me
End Sub

Private Sub Menuju()
    If FrmMain.Winsock(sockdari).State = sckConnected Then
        FrmMain.Winsock(sockdari).SendData "pindahke|_|_|_|_|_|_|"
    End If
End Sub

Private Sub Tertuju()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    If FrmMain.Winsock(sockke).State = sckConnected Then
        FrmMain.Winsock(sockke).SendData "pindahdari|" & mulai_dari & "|" & jenis_dari & "|_|_|_|_|"
        FrmMain.Winsock(sockke).Close
        Unload FrmMain.Winsock(sockke)
        cn.Open cnstr
        rs.ActiveConnection = cn
        rs.Open "SELECT socket.* FROM socket WHERE socket.socket = " & sockke, cn, adOpenStatic, adLockOptimistic
        If Not rs.RecordCount = 0 Then
           rs.Delete
        End If
        cn.Close
    End If
End Sub

Private Sub CmdPindah_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdPindah, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub Form_Load()
    FrmPindahClient.Caption = LangGlobPindah & " " & LangClient
    Label1.Caption = LangDari & " " & LangClient
    Label2.Caption = LangKe & " " & LangClient
    CmdPindah.Caption = LangGlobPindah
    CmdPindah.ToolTipText = LangGlobPindah
    CmdBatal.Caption = LangBatal
    CmdBatal.ToolTipText = LangBatal
End Sub

Private Sub TxtNoDari_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 13 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
    If kwyascii = 13 Then
        TxtNoKe.SetFocus
    End If
End Sub

Private Sub TxtNoKe_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 13 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub
