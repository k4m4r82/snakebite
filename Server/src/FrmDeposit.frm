VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmDeposit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tambah Deposit"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6015
   Icon            =   "FrmDeposit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   6015
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Tambah"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox TxtDeposit 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox TxtUser 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      MaxLength       =   15
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   2055
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   3625
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      Appearance      =   0
   End
   Begin VB.Label Label2 
      Caption         =   "Deposit"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "User ID"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "FrmDeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim depawal As Long

Private Sub Simpan()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open "SELECT member.* FROM member WHERE member.user = '" & TxtUser.Text & "'", cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        TxtDeposit.SetFocus
        CmdSimpan.Enabled = True
        rs(5) = depawal + Val(TxtDeposit.Text)
        rs.Update
    End If
    cn.Close
    CmdSimpan.Enabled = False
    Call Tampil
End Sub

Private Sub CmdSimpan_Click()
    If Not TxtDeposit.Text = "" Then
        Call Simpan
        TxtUser.Text = ""
        TxtDeposit.Text = ""
        TxtUser.SetFocus
    Else
        TxtDeposit.SetFocus
    End If
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

Private Sub Form_Load()
    Me.Caption = LangTambah & " " & LangDeposit
    Label1.Caption = LangNamaUser
    Label2.Caption = LangDeposit
    CmdSimpan.Caption = LangTambah
    CmdTutup.Caption = LangTutup
    Call Tampil
End Sub

Private Sub Tampil()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim sql As String
    Grid.TextMatrix(0, 0) = LangNamaUser
    Grid.TextMatrix(0, 1) = LangNamaAsli
    Grid.TextMatrix(0, 2) = LangDeposit
    Grid.ColWidth(0) = 1000
    Grid.ColWidth(1) = 1500
    Grid.ColWidth(2) = 2000
    sql = "SELECT member.* FROM member"
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open sql, cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        Grid.Rows = rs.RecordCount + 1
        rs.MoveFirst
        For i = 1 To rs.RecordCount
            Grid.TextMatrix(i, 0) = rs(0)
            Grid.TextMatrix(i, 1) = rs(1)
            Grid.TextMatrix(i, 2) = Format(rs(5), "##,##0.00")
            rs.MoveNext
        Next i
    Else
        Grid.Rows = 1
    End If
    cn.Close
End Sub

Private Sub Grid_Click()
    If Grid.Rows > 1 Then
        TxtUser.Text = Grid.TextMatrix(Grid.Row, 0)
        TxtUser.SetFocus
        SendKeys "{ENTER}"
    Else
        TxtUser.SetFocus
    End If
End Sub

Private Sub TxtDeposit_GotFocus()
    TxtDeposit.Backcolor = &HFFC0C0
End Sub

Private Sub TxtDeposit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CmdSimpan.Enabled = True Then
            CmdSimpan.SetFocus
        End If
    End If
End Sub

Private Sub TxtDeposit_LostFocus()
    TxtDeposit.Backcolor = &H80000005
End Sub

Private Sub TxtUser_GotFocus()
    TxtUser.Backcolor = &HFFC0C0
End Sub

Private Sub TxtUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not TxtUser.Text = "" Then
            CekUser
        End If
    End If
End Sub

Private Sub TxtUser_LostFocus()
    TxtUser.Backcolor = &H80000005
End Sub

Private Sub CekUser()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open "SELECT member.* FROM member WHERE member.user = '" & TxtUser.Text & "'", cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        TxtDeposit.SetFocus
        CmdSimpan.Enabled = True
        depawal = rs(5)
    Else
        TxtUser.Text = ""
        TxtUser.SetFocus
    End If
    cn.Close
End Sub

