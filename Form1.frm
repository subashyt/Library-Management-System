VERSION 5.00
Begin VB.Form frmstart 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080C0FF&
   BorderStyle     =   0  'None
   Caption         =   "Welcome to Library Management System"
   ClientHeight    =   2715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6840
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   2775
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   2715
      ScaleWidth      =   6795
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.PictureBox grnpic 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2880
         Picture         =   "Form1.frx":3CA4E
         ScaleHeight     =   255
         ScaleWidth      =   3135
         TabIndex        =   1
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label lbl_loading 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Loading data files.."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   2
         Top             =   0
         Width           =   2175
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   20
      Left            =   6240
      Top             =   120
   End
End
Attribute VB_Name = "frmstart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmdstart_Click()
 'Unload this screen and display the login window
  Unload Me
  frmlogin.Show
End Sub

Private Sub Form_Load()
grnpic.width = 0
End Sub

Private Sub Timer2_Timer()
Dim width As Integer
grnpic.width = grnpic.width + 20
Select Case grnpic.width
Case Is = 1815
lbl_loading.Caption = "Starting App.."
Case Is = 3135
frmlogin.Show
Timer2.Interval = 0
Timer2.Enabled = False
Unload Me
End Select

End Sub
