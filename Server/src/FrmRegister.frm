VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registrasi Client"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5775
   Icon            =   "FrmRegister.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   5775
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton CmdUbah 
      Caption         =   "&Ubah"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox TxtIP 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1320
      MaxLength       =   25
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox TxtNoClient 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   2295
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   4048
      _Version        =   393216
      FixedCols       =   0
      Appearance      =   0
   End
   Begin VB.Label Label2 
      Caption         =   "IP Adress"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "No. Client"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "FrmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdHapus_Click()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open "SELECT ip.* FROM ip WHERE ip.no_client = " & Val(TxtNoClient.Text), cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        rs.Delete
    End If
    cn.Close
    Bersih
    CmdSimpan.Enabled = False
    CmdUbah.Enabled = False
    CmdHapus.Enabled = False
    TxtNoClient.Text = ""
    TxtNoClient.SetFocus
    Call Tampil
End Sub

Private Sub CmdHapus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdHapus, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdSimpan_Click()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open "SELECT ip.* FROM ip WHERE ip.no_client =" & Val(TxtNoClient.Text), cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        rs(0) = Val(TxtNoClient.Text)
        rs(1) = TxtIP.Text
        rs.Update
    Else
        rs.AddNew
        rs(0) = Val(TxtNoClient.Text)
        rs(1) = TxtIP.Text
        rs.Update
    End If
    cn.Close
    Bersih
    CmdSimpan.Enabled = False
    CmdUbah.Enabled = False
    CmdHapus.Enabled = False
    TxtNoClient.Text = ""
    TxtNoClient.SetFocus
    Call Tampil
End Sub

Private Sub CmdSimpan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdSimpan, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdTutup_Click()
    Unload Me
End Sub

Private Sub CmdTutup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdTutup, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdUbah_Click()
    CmdUbah.Enabled = False
    CmdHapus.Enabled = False
    CmdSimpan.Enabled = True
    TxtIP.SetFocus
End Sub

Private Sub CmdUbah_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdUbah, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub Form_Load()
    Me.Caption = LangmnuRegClient
    Label1.Caption = LangNo & " " & LangClient
    Label2.Caption = LangIP
    CmdSimpan.Caption = LangSimpan
    CmdUbah.Caption = LangUbah
    CmdHapus.Caption = LangHapus
    CmdTutup.Caption = LangTutup
    Call Tampil
End Sub

Private Sub Tampil()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim sql As String
    Grid.TextMatrix(0, 0) = LangNo & " " & LangClient
    Grid.TextMatrix(0, 1) = LangIP
    Grid.ColWidth(0) = 1000
    Grid.ColWidth(1) = 2500
    sql = "SELECT ip.* FROM ip"
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open sql, cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        Grid.Rows = rs.RecordCount + 1
        rs.MoveFirst
        For i = 1 To rs.RecordCount
            Grid.TextMatrix(i, 0) = rs(0)
            Grid.TextMatrix(i, 1) = rs(1)
            rs.MoveNext
        Next i
    Else
        Grid.Rows = 1
    End If
    cn.Close
End Sub

Private Sub Grid_Click()
    If Grid.Rows > 1 Then
        TxtNoClient.Text = Grid.TextMatrix(Grid.Row, 0)
        TxtNoClient.SetFocus
        SendKeys "{ENTER}"
    Else
        TxtNoClient.SetFocus
    End If
End Sub

Private Sub TxtIP_GotFocus()
    TxtIP.Backcolor = &HFFC0C0
End Sub

Private Sub TxtIP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CmdSimpan.Enabled = True Then
            CmdSimpan.SetFocus
        End If
    End If
End Sub

Private Sub TxtIP_LostFocus()
    TxtIP.Backcolor = &H80000005
End Sub

Private Sub TxtNoClient_GotFocus()
    TxtNoClient.Backcolor = &HFFC0C0
End Sub

Private Sub TxtNoClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not TxtNoClient.Text = "" Then
            CekClient
        End If
    End If
End Sub

Private Sub CekClient()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open "SELECT ip.* FROM ip WHERE ip.no_client = " & Val(TxtNoClient.Text), cn, adOpenStatic, adLockOptimistic
    If rs.RecordCount = 0 Then
        Call Baru
    Else
        TxtIP.Text = rs(1)
        CmdUbah.Enabled = True
        CmdHapus.Enabled = True
        CmdSimpan.Enabled = False
    End If
    cn.Close
End Sub

Private Sub Bersih()
    TxtIP.Text = ""
End Sub

Private Sub Baru()
    Bersih
    TxtIP.SetFocus
    CmdSimpan.Enabled = True
    CmdUbah.Enabled = False
    CmdHapus.Enabled = False
End Sub

Private Sub TxtNoClient_LostFocus()
    TxtNoClient.Backcolor = &H80000005
End Sub
