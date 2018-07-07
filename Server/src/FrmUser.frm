VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manajemen User"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   Icon            =   "FrmUser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   2295
      Left            =   360
      TabIndex        =   17
      Top             =   3840
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      Appearance      =   0
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton CmdUbah 
      Caption         =   "&Ubah"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3240
      Width           =   975
   End
   Begin VB.OptionButton TxtLoginAdmin 
      Appearance      =   0  'Flat
      Caption         =   "Administrator"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   2640
      Width           =   1335
   End
   Begin VB.OptionButton TxtLoginUser 
      Appearance      =   0  'Flat
      Caption         =   "Operator"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox TxtTelpon 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   4
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox TxtAlamat 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1440
      MaxLength       =   25
      TabIndex        =   3
      Top             =   1680
      Width           =   3135
   End
   Begin VB.TextBox TxtNama 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox TxtPassword 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox TxtUser 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis User"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Telpon"
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Asli"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "FrmUser"
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
    rs.Open "SELECT account.* FROM account WHERE account.user = '" & TxtUser.Text & "'", cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        If rs(2) = "sysadmin" Then
            MsgBox LangSysadmin, vbInformation, LangInformasi
        Else
            rs.Delete
        End If
    End If
    cn.Close
    Bersih
    CmdSimpan.Enabled = False
    CmdUbah.Enabled = False
    CmdHapus.Enabled = False
    TxtUser.Text = ""
    TxtUser.SetFocus
    Call Tampil
End Sub

Private Sub CmdHapus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdHapus, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdSimpan_Click()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim passwd As String
    passwd = EncryptText(TxtPassword.Text, "password")
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open "SELECT account.* FROM account WHERE account.user = '" & TxtUser.Text & "'", cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        If rs(2) = "sysadmin" Then
            rs(0) = TxtUser.Text
            rs(5) = passwd
            rs(1) = TxtNama.Text
            rs(3) = TxtAlamat.Text
            rs(4) = TxtTelpon.Text
            rs(2) = "sysadmin"
            rs.Update
        Else
            rs(0) = TxtUser.Text
            rs(5) = passwd
            rs(1) = TxtNama.Text
            rs(3) = TxtAlamat.Text
            rs(4) = TxtTelpon.Text
            If TxtLoginAdmin.value = True Then
                rs(2) = "admin"
            Else
                rs(2) = "user"
            End If
            rs.Update
        End If
    Else
        rs.AddNew
        rs(0) = TxtUser.Text
        rs(5) = passwd
        rs(1) = TxtNama.Text
        rs(3) = TxtAlamat.Text
        rs(4) = TxtTelpon.Text
        If TxtLoginAdmin.value = True Then
            rs(2) = "admin"
        Else
            rs(2) = "user"
        End If
        rs.Update
    End If
    cn.Close
    Bersih
    CmdSimpan.Enabled = False
    CmdUbah.Enabled = False
    CmdHapus.Enabled = False
    TxtUser.Text = ""
    TxtUser.SetFocus
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
    TxtPassword.SetFocus
End Sub

Private Sub CmdUbah_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdUbah, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub Form_Load()
    FrmUser.Caption = LangManajemenUser
    Label2.Caption = LangNamaUser
    Label3.Caption = LangPassword
    Label4.Caption = LangNamaAsli
    Label5.Caption = LangAlamat
    Label6.Caption = LangNoTelp
    Label7.Caption = LangJenisUser
    TxtLoginUser.Caption = LangOperator
    TxtLoginAdmin.Caption = LangAdministrator
    CmdSimpan.Caption = LangSimpan
    CmdUbah.Caption = LangUbah
    CmdHapus.Caption = LangHapus
    CmdTutup.Caption = LangTutup
    Call Tampil
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

Private Sub TxtAlamat_GotFocus()
    TxtAlamat.Backcolor = &HFFC0C0
End Sub

Private Sub TxtAlamat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtTelpon.SetFocus
    End If
End Sub

Private Sub TxtAlamat_LostFocus()
    TxtAlamat.Backcolor = &H80000005
End Sub

Private Sub TxtLoginAdmin_Click()
    If CmdSimpan.Enabled = True Then
        CmdSimpan.SetFocus
    End If
End Sub

Private Sub TxtLoginUser_Click()
    If CmdSimpan.Enabled = True Then
        CmdSimpan.SetFocus
    End If
End Sub

Private Sub TxtNama_GotFocus()
    TxtNama.Backcolor = &HFFC0C0
End Sub

Private Sub TxtNama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtAlamat.SetFocus
    End If
End Sub

Private Sub TxtNama_LostFocus()
    TxtNama.Backcolor = &H80000005
End Sub

Private Sub TxtPassword_GotFocus()
    TxtPassword.Backcolor = &HFFC0C0
End Sub

Private Sub TxtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtNama.SetFocus
    End If
End Sub

Private Sub TxtPassword_LostFocus()
    TxtPassword.Backcolor = &H80000005
End Sub

Private Sub TxtTelpon_GotFocus()
    TxtTelpon.Backcolor = &HFFC0C0
End Sub

Private Sub TxtTelpon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtLoginUser.SetFocus
    End If
End Sub

Private Sub TxtTelpon_LostFocus()
    TxtTelpon.Backcolor = &H80000005
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
    rs.Open "SELECT account.* FROM account WHERE account.user = '" & TxtUser.Text & "'", cn, adOpenStatic, adLockOptimistic
    If rs.RecordCount = 0 Then
        Call Baru
    Else
        TxtPassword.Text = DecryptText(rs(5), "password")
        TxtNama.Text = rs(1)
        TxtAlamat.Text = rs(3)
        TxtTelpon.Text = rs(4)
        If rs(2) = "sysadmin" Then
            TxtLoginAdmin.value = True
            TxtLoginUser.value = False
        ElseIf rs(2) = "admin" Then
            TxtLoginAdmin.value = True
            TxtLoginUser.value = False
        Else
            TxtLoginUser.value = True
            TxtLoginAdmin.value = False
        End If
        CmdUbah.Enabled = True
        CmdHapus.Enabled = True
        CmdSimpan.Enabled = False
    End If
    cn.Close
End Sub

Private Sub Bersih()
    TxtPassword.Text = ""
    TxtNama.Text = ""
    TxtAlamat.Text = ""
    TxtTelpon.Text = ""
    TxtLoginUser.value = False
    TxtLoginAdmin.value = False
End Sub

Private Sub Baru()
    Bersih
    TxtPassword.SetFocus
    CmdSimpan.Enabled = True
    CmdUbah.Enabled = False
    CmdHapus.Enabled = False
End Sub

Private Sub Tampil()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim sql As String
    Grid.TextMatrix(0, 0) = LangNamaUser
    Grid.TextMatrix(0, 1) = LangNamaAsli
    Grid.TextMatrix(0, 2) = LangJenisUser
    Grid.ColWidth(0) = 1000
    Grid.ColWidth(1) = 1500
    Grid.ColWidth(2) = 2000
    sql = "SELECT account.* FROM account"
    cn.Open cnstr
    rs.ActiveConnection = cn
    rs.Open sql, cn, adOpenStatic, adLockOptimistic
    If Not rs.RecordCount = 0 Then
        Grid.Rows = rs.RecordCount + 1
        rs.MoveFirst
        For i = 1 To rs.RecordCount
            Grid.TextMatrix(i, 0) = rs(0)
            Grid.TextMatrix(i, 1) = rs(1)
            Grid.TextMatrix(i, 2) = rs(2)
            rs.MoveNext
        Next i
    Else
        Grid.Rows = 1
    End If
    cn.Close
End Sub
