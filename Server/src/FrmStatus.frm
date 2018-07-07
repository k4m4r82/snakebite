VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form FrmStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Status Koneksi Client"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6540
   Icon            =   "FrmStatus.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   6540
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdTutup 
      Caption         =   "Tutup"
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   975
   End
   Begin ComctlLib.ListView ListStatus 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   7011
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "No."
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "No. Client"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Socket"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "FrmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lststatus As ListItem

Private Sub InitStatus()
    For i = 1 To ClientMax
        Set lststatus = ListStatus.ListItems.Add(, , i)
        If CekClientKonek(i) = LangAktif Then
            lststatus.SubItems(1) = LangClient & " " & i
            lststatus.SubItems(2) = CekClientKonek(i)
            lststatus.SubItems(3) = CekClientSocket(i)
        Else
            lststatus.SubItems(1) = LangClient & " " & i
            lststatus.SubItems(2) = CekClientKonek(i)
            lststatus.SubItems(3) = ""
        End If
    Next i
End Sub

Private Sub CmdTutup_Click()
    Unload Me
End Sub

Private Sub CmdTutup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdTutup, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub Form_Load()
    FrmStatus.Caption = LangStatusKoneksi
    ListStatus.ColumnHeaders(1).Text = LangNo
    ListStatus.ColumnHeaders(2).Text = LangNo & " " & LangClient
    ListStatus.ColumnHeaders(3).Text = LangStatus
    ListStatus.ColumnHeaders(4).Text = LangSocket
    CmdTutup.Caption = LangTutup
    InitStatus
End Sub
