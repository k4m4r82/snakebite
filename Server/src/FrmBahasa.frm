VERSION 5.00
Begin VB.Form FrmBahasa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ganti Bahasa"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3915
   Icon            =   "FrmBahasa.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   3915
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdOK 
      Appearance      =   0  'Flat
      Caption         =   "&OK"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton CmdBatal 
      Appearance      =   0  'Flat
      Caption         =   "&Batal"
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.ComboBox ComboBahasa 
      Height          =   315
      ItemData        =   "FrmBahasa.frx":030A
      Left            =   1560
      List            =   "FrmBahasa.frx":0314
      TabIndex        =   0
      Text            =   "Indonesia"
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Bahasa"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "FrmBahasa"
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

Private Sub CmdOK_Click()
    If ComboBahasa.Text = "Indonesia" Then
        Call RegSaveString(HKEY_LOCAL_MACHINE, "Software\snakebite\server\" & VersiAplikasi & "\", "lang", "id")
        lang = "id"
        Call Bahasa
    ElseIf ComboBahasa.Text = "English" Then
        Call RegSaveString(HKEY_LOCAL_MACHINE, "Software\snakebite\server\" & VersiAplikasi & "\", "lang", "en")
        lang = "en"
        Call Bahasa
    End If
    Call LangMenu
    Call LangToolBar
    Call AturGrid
    MsgBox langupdateukses & " " & LangRestart, vbInformation, LangInformasi
    End
End Sub

Private Sub CmdOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdOK, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub Form_Load()
    Label1.Caption = LangBahasa
    Me.Caption = LangmnuGantiBahasa
    CmdOK.Caption = LangUpdate
    CmdBatal.Caption = LangBatal
    If lang = "id" Then
        ComboBahasa.Text = "Indonesia"
    ElseIf lang = "en" Then
        ComboBahasa.Text = "English"
    End If
End Sub
