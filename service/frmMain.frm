VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close"
      Height          =   375
      Left            =   3300
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2730
      Width           =   1275
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Stop ISS Services"
      Height          =   375
      Left            =   1230
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1650
      Width           =   2025
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Start IIS Services"
      Height          =   375
      Left            =   1230
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   180
      Picture         =   "frmMain.frx":0000
      Top             =   630
      Width           =   720
   End
   Begin VB.Label Label4 
      BackColor       =   &H00EDA84D&
      Caption         =   "  ISS Services"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   240
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6285
   End
   Begin VB.Image Image2 
      Height          =   60
      Left            =   0
      Picture         =   "frmMain.frx":08F0
      Top             =   230
      Width           =   9750
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    ServiceCommand = 0
    frmService.Show 1

    If ServiceCommand = 99 Then End

End Sub

Private Sub Command2_Click()

    ServiceCommand = 1
    frmService.Show 1

    If ServiceCommand = 99 Then End

End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()

    Center Me

End Sub

