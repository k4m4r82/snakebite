VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmLaporanDeposit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Deposit Member"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7365
   Icon            =   "FrmLaporanDeposit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   7365
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdTutup 
      Caption         =   "Tutup"
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton CmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid GridTransaksi 
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   3413
      _Version        =   393216
      Cols            =   5
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   3836
      _Version        =   393216
      Cols            =   5
      Appearance      =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Detail Transaksi"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   1335
   End
End
Attribute VB_Name = "FrmLaporanDeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CekDeposit()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    Dim sql As String
    sql = "SELECT member.* FROM member"
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open sql, cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        MSFlexGrid.Rows = rs.RecordCount + 1
        rs.MoveFirst
        For i = 1 To rs.RecordCount
            MSFlexGrid.TextMatrix(i, 0) = rs(0) 'user name
            MSFlexGrid.TextMatrix(i, 1) = rs(1) 'nama
            MSFlexGrid.TextMatrix(i, 2) = rs(2) 'alamat
            MSFlexGrid.TextMatrix(i, 3) = rs(3) 'telp
            MSFlexGrid.TextMatrix(i, 4) = Format(rs(5), "##,##0.00") 'deposit
            rs.MoveNext
        Next i
    Else
        MSFlexGrid.Rows = 1
    End If
    cn.Close
End Sub

Private Sub CmdRefresh_Click()
    Call CekDeposit
End Sub

Private Sub CmdRefresh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdRefresh, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdTutup_Click()
    Unload Me
End Sub

Private Sub CmdTutup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdTutup, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub Form_Load()
    Call AturGridDeposit
    Call AturGridTransaksi
    Call CekDeposit
    FrmLaporanDeposit.Caption = LangLaporanDeposit
    CmdRefresh.Caption = LangRefresh
    CmdTutup.Caption = LangTutup
    Label1.Caption = LangDetailTransaksi
End Sub

Private Sub AturGridDeposit()
    MSFlexGrid.TextMatrix(0, 0) = LangNamaUser
    MSFlexGrid.TextMatrix(0, 1) = LangNamaMember
    MSFlexGrid.TextMatrix(0, 2) = LangAlamat
    MSFlexGrid.TextMatrix(0, 3) = LangNoTelp
    MSFlexGrid.TextMatrix(0, 4) = LangDeposit
    MSFlexGrid.ColWidth(0) = 800
    MSFlexGrid.ColWidth(1) = 1200
    MSFlexGrid.ColWidth(2) = 2300
    MSFlexGrid.ColWidth(3) = 1300
    MSFlexGrid.ColWidth(4) = 1200
End Sub

Private Sub AturGridTransaksi()
    GridTransaksi.TextMatrix(0, 0) = LangNota
    GridTransaksi.TextMatrix(0, 1) = LangTanggal
    GridTransaksi.TextMatrix(0, 2) = LangNo & "" & LangClient
    GridTransaksi.TextMatrix(0, 3) = LangDurasi
    GridTransaksi.TextMatrix(0, 4) = LangTotal
    GridTransaksi.ColWidth(0) = 800
    GridTransaksi.ColWidth(1) = 1600
    GridTransaksi.ColWidth(2) = 1200
    GridTransaksi.ColWidth(3) = 1400
    GridTransaksi.ColWidth(4) = 1800
End Sub

Private Sub CekDetail()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
    sql = "SELECT billing.* FROM billing WHERE billing.user = '" & MSFlexGrid.TextMatrix(MSFlexGrid.Row, 0) & "' AND billing.jenis = 'member'"
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open sql, cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        GridTransaksi.Rows = rs.RecordCount + 1
        rs.MoveFirst
        For i = 1 To rs.RecordCount
            GridTransaksi.TextMatrix(i, 0) = rs(0) 'no nota
            GridTransaksi.TextMatrix(i, 1) = rs(1) 'tanggal
            GridTransaksi.TextMatrix(i, 2) = LangClient & " " & rs(2) 'no client
            GridTransaksi.TextMatrix(i, 3) = rs(6) 'durasi
            GridTransaksi.TextMatrix(i, 4) = Format(rs(7), "##,##0.00") 'total
            rs.MoveNext
        Next i
    Else
        GridTransaksi.Rows = 1
    End If
    cn.Close
End Sub

Private Sub MSFlexGrid_Click()
    Call CekDetail
End Sub
