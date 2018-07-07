VERSION 5.00
Begin VB.Form FrmKal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kalkulator"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3285
   Icon            =   "FrmKal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   3285
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdBkspc 
      Caption         =   "Backspace"
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton CmdC 
      Caption         =   "C"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton CmdSamadengan 
      Caption         =   "="
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton CmdTambah 
      Caption         =   "+"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton CmdKoma 
      Caption         =   ","
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton CmdPM 
      Caption         =   "+/-"
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Cmd0 
      Caption         =   "0"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton CmdSeper 
      Caption         =   "1/x"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton CmdKurang 
      Caption         =   "-"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Cmd3 
      Caption         =   "3"
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Cmd2 
      Caption         =   "2"
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "1"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Cmdpersen 
      Caption         =   "%"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton CmdKali 
      Caption         =   "*"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton Cmd6 
      Caption         =   "6"
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton Cmd5 
      Caption         =   "5"
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton Cmd4 
      Caption         =   "4"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton CmdAkar 
      Caption         =   "sqrt"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton CmdBagi 
      Caption         =   "/"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Cmd9 
      Caption         =   "9"
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Cmd8 
      Caption         =   "8"
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Cmd7 
      Caption         =   "7"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label TxtHasil 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "FrmKal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hasil, bilang As Double
Dim bil As Integer
Dim statop As String
Dim operat As String

Private Sub Cmd0_Click()
    bil = 0
    If statop = "input" Then
        TxtHasil.Caption = bilang & bil
        bilang = Val(TxtHasil.Caption)
    ElseIf statop = "input2" Then
        TxtHasil.Caption = bilang & bil
        bilang = Val(TxtHasil.Caption)
    ElseIf statop = "tampil" Then
        statop = "input2"
        TxtHasil.Caption = bilang & bil
        bilang = Val(TxtHasil.Caption)
    End If
End Sub

Private Sub Cmd0_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover Cmd0, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub Cmd1_Click()
    bil = 1
    If statop = "input" Then
        TxtHasil.Caption = bilang & bil
        bilang = Val(TxtHasil.Caption)
    ElseIf statop = "input2" Then
        TxtHasil.Caption = bilang & bil
        bilang = Val(TxtHasil.Caption)
    ElseIf statop = "tampil" Then
        statop = "input2"
        TxtHasil.Caption = bilang & bil
        bilang = Val(TxtHasil.Caption)
    End If
End Sub

Private Sub Cmd1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover Cmd1, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub Cmd2_Click()
    bil = 2
    If statop = "input" Then
        TxtHasil.Caption = bilang & bil
        bilang = Val(TxtHasil.Caption)
    ElseIf statop = "input2" Then
        TxtHasil.Caption = bilang & bil
        bilang = Val(TxtHasil.Caption)
    ElseIf statop = "tampil" Then
        statop = "input2"
        TxtHasil.Caption = bilang & bil
        bilang = Val(TxtHasil.Caption)
    End If
End Sub

Private Sub Cmd2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover Cmd2, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub Cmd3_Click()
    bil = 3
    If statop = "input" Then
        TxtHasil.Caption = bilang & bil
        bilang = Val(TxtHasil.Caption)
    ElseIf statop = "input2" Then
        TxtHasil.Caption = bilang & bil
        bilang = Val(TxtHasil.Caption)
    ElseIf statop = "tampil" Then
        statop = "input2"
        TxtHasil.Caption = bilang & bil
        bilang = Val(TxtHasil.Caption)
    End If
End Sub

Private Sub Cmd3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover Cmd3, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub Cmd4_Click()
    bil = 4
    If statop = "input" Then
        TxtHasil.Caption = bilang & bil
        bilang = Val(TxtHasil.Caption)
    ElseIf statop = "input2" Then
        TxtHasil.Caption = bilang & bil
        bilang = Val(TxtHasil.Caption)
    ElseIf statop = "tampil" Then
        statop = "input2"
        TxtHasil.Caption = bilang & bil
        bilang = Val(TxtHasil.Caption)
    End If
End Sub

Private Sub Cmd4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover Cmd4, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub Cmd5_Click()
    bil = 5
    If statop = "input" Then
        TxtHasil.Caption = bilang & bil
        bilang = Val(TxtHasil.Caption)
    ElseIf statop = "input2" Then
        TxtHasil.Caption = bilang & bil
        bilang = Val(TxtHasil.Caption)
    ElseIf statop = "tampil" Then
        statop = "input2"
        TxtHasil.Caption = bilang & bil
        bilang = Val(TxtHasil.Caption)
    End If
End Sub

Private Sub Cmd5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover Cmd5, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub Cmd6_Click()
    bil = 6
    If statop = "input" Then
        TxtHasil.Caption = bilang & bil
        bilang = Val(TxtHasil.Caption)
    ElseIf statop = "input2" Then
        TxtHasil.Caption = bilang & bil
        bilang = Val(TxtHasil.Caption)
    ElseIf statop = "tampil" Then
        statop = "input2"
        TxtHasil.Caption = bilang & bil
        bilang = Val(TxtHasil.Caption)
    End If
End Sub

Private Sub Cmd6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover Cmd6, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub Cmd7_Click()
    bil = 7
    If statop = "input" Then
        TxtHasil.Caption = bilang & bil
        bilang = Val(TxtHasil.Caption)
    ElseIf statop = "input2" Then
        TxtHasil.Caption = bilang & bil
        bilang = Val(TxtHasil.Caption)
    ElseIf statop = "tampil" Then
        statop = "input2"
        TxtHasil.Caption = bilang & bil
        bilang = Val(TxtHasil.Caption)
    ElseIf statop = "tampil" Then
        statop = "input2"
        TxtHasil.Caption = bilang & bil
        bilang = Val(TxtHasil.Caption)
    End If
End Sub

Private Sub Cmd7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover Cmd7, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub Cmd8_Click()
    bil = 8
    If statop = "input" Then
        TxtHasil.Caption = bilang & bil
        bilang = Val(TxtHasil.Caption)
    ElseIf statop = "input2" Then
        TxtHasil.Caption = bilang & bil
        bilang = Val(TxtHasil.Caption)
    ElseIf statop = "tampil" Then
        statop = "input2"
        TxtHasil.Caption = bilang & bil
        bilang = Val(TxtHasil.Caption)
    End If
End Sub

Private Sub Cmd8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover Cmd8, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub Cmd9_Click()
    bil = 9
    If statop = "input" Then
        TxtHasil.Caption = bilang & bil
        bilang = Val(TxtHasil.Caption)
    ElseIf statop = "input2" Then
        TxtHasil.Caption = bilang & bil
        bilang = Val(TxtHasil.Caption)
    ElseIf statop = "tampil" Then
        statop = "input2"
        TxtHasil.Caption = bilang & bil
        bilang = Val(TxtHasil.Caption)
    End If
End Sub

Private Sub Cmd9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover Cmd9, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdAkar_Click()
    operat = "akar"
    If statop = "input" Then
        statop = "input2"
        hasil = bilang
        bilang = 0
        TxtHasil.Caption = bilang
    ElseIf statop = "input2" Then
        statop = "tampil"
        hasil = Sqr(hasil)
        TxtHasil.Caption = hasil
        bilang = 0
    End If
End Sub

Private Sub CmdAkar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdAkar, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdBagi_Click()
    operat = "bagi"
    If statop = "input" Then
        statop = "input2"
        hasil = bilang
        bilang = 0
        TxtHasil.Caption = bilang
    ElseIf statop = "input2" Then
        statop = "tampil"
        hasil = hasil / bilang
        TxtHasil.Caption = hasil
        bilang = 0
    End If
End Sub

Private Sub CmdBagi_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdBagi, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdBkspc_Click()
    If Len(TxtHasil.Caption) >= 1 Then
        If Len(TxtHasil.Caption) > 1 Then
            TxtHasil.Caption = Mid(TxtHasil.Caption, 1, Len(TxtHasil.Caption) - 1)
        Else
            TxtHasil.Caption = "0"
        End If
    End If
    If statop = "input" Then
        bilang = Val(TxtHasil.Caption)
    End If
End Sub

Private Sub CmdBkspc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdBkspc, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdC_Click()
    Call awal
End Sub

Private Sub CmdC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdC, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdKali_Click()
    operat = "kali"
    If statop = "input" Then
        statop = "input2"
        hasil = bilang
        bilang = 0
        TxtHasil.Caption = bilang
    ElseIf statop = "input2" Then
        statop = "tampil"
        hasil = hasil * bilang
        TxtHasil.Caption = hasil
        bilang = 0
    End If
End Sub

Private Sub CmdKali_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdKali, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdKoma_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdKoma, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdKurang_Click()
    operat = "kurang"
    If statop = "input" Then
        statop = "input2"
        hasil = bilang
        bilang = 0
        TxtHasil.Caption = bilang
    ElseIf statop = "input2" Then
        statop = "tampil"
        hasil = hasil - bilang
        TxtHasil.Caption = hasil
        bilang = 0
    End If
End Sub

Private Sub CmdKurang_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdKurang, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub Cmdpersen_Click()
    operat = "persen"
    If statop = "input" Then
        statop = "input2"
        hasil = bilang
        bilang = 0
        TxtHasil.Caption = bilang
    ElseIf statop = "input2" Then
        statop = "tampil"
        hasil = hasil / bilang * 100
        TxtHasil.Caption = hasil
        bilang = 0
    End If
End Sub

Private Sub Cmdpersen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover Cmdpersen, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdPM_Click()
    operat = "pm"
    If statop = "input" Then
        bilang = Val(TxtHasil.Caption) * -1
        TxtHasil.Caption = bilang
    ElseIf statop = "input2" Then
        bilang = Val(TxtHasil.Caption) * -1
        TxtHasil.Caption = bilang
    ElseIf statop = "tampil" Then
        statop = "input2"
        bilang = Val(TxtHasil.Caption) * -1
        TxtHasil.Caption = bilang
    ElseIf statop = "tampil" Then
        statop = "input2"
        bilang = Val(TxtHasil.Caption) * -1
        TxtHasil.Caption = bilang
    End If
End Sub

Private Sub CmdPM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdPM, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdSamadengan_Click()
    If operat = "kali" Then
        If statop = "input" Then
            statop = "input2"
            hasil = bilang
            bilang = 0
            TxtHasil.Caption = bilang
        ElseIf statop = "input2" Then
            statop = "tampil"
            hasil = hasil * bilang
            TxtHasil.Caption = hasil
            bilang = 0
        End If
    ElseIf operat = "tambah" Then
        If statop = "input" Then
            statop = "input2"
            hasil = bilang
            bilang = 0
            TxtHasil.Caption = bilang
        ElseIf statop = "input2" Then
            statop = "tampil"
            hasil = hasil + bilang
            TxtHasil.Caption = hasil
            bilang = 0
        End If
    ElseIf operat = "bagi" Then
        If statop = "input" Then
            statop = "input2"
            hasil = bilang
            bilang = 0
            TxtHasil.Caption = bilang
        ElseIf statop = "input2" Then
            statop = "tampil"
            hasil = hasil / bilang
            TxtHasil.Caption = hasil
            bilang = 0
        End If
    ElseIf operat = "kurang" Then
        If statop = "input" Then
            statop = "input2"
            hasil = bilang
            bilang = 0
            TxtHasil.Caption = bilang
        ElseIf statop = "input2" Then
            statop = "tampil"
            hasil = hasil - bilang
            TxtHasil.Caption = hasil
            bilang = 0
        End If
    End If
End Sub

Private Sub CmdSamadengan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdSamadengan, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdSeper_Click()
    operat = "seper"
    If statop = "input" Then
        statop = "input2"
        hasil = bilang
        bilang = 0
        TxtHasil.Caption = bilang
    ElseIf statop = "input2" Then
        statop = "tampil"
        hasil = 1 / hasil
        TxtHasil.Caption = hasil
        bilang = 0
    End If
End Sub

Private Sub CmdSeper_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdSeper, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdTambah_Click()
    operat = "tambah"
    If statop = "input" Then
        statop = "input2"
        hasil = bilang
        bilang = 0
        TxtHasil.Caption = bilang
    ElseIf statop = "input2" Then
        statop = "tampil"
        hasil = hasil + bilang
        TxtHasil.Caption = hasil
        bilang = 0
    End If
End Sub

Private Sub CmdTambah_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdTambah, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub Form_Activate()
    CmdC.SetFocus
End Sub

Private Sub Form_Load()
    Call awal
End Sub

Private Sub awal()
    hasil = 0
    bil = 0
    bilang = 0
    statop = "input"
    TxtHasil.Caption = hasil
End Sub
