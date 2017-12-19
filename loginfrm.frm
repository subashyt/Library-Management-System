VERSION 5.00
Begin VB.Form frmlogin 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   5715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdlogin 
      Caption         =   "Login"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txtlogin 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   600
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label loginlbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter The Security Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "loginfrm.frx":0000
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim secp As login

Private Sub cmdlogin_Click()
 'open the fsecurity text file to read the password stored
 Open App.Path & "\filesecurity.txt" For Random As #1 Len = 20
  Get #1, 1, secp
 Close #1
 If txtlogin.Text = secp.password Then
   Unload Me
   frmmain.Show
  Else
  'creating a message box to inform the user that the password is wrong
  'also to make the user retype the password
  MsgBox "Wrong Password!", vbInformation + vbOKOnly, "Try Again"
  txtlogin.Text = ""
  txtlogin.SetFocus
 End If
End Sub

Private Sub Form_Load()
 'to set the new password if there is not any
 If Dir(App.Path & "\filesecurity.txt") = "" Then
  frmchangepass.Show 1
 End If
End Sub

