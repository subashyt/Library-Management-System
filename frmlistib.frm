VERSION 5.00
Begin VB.Form frmlistib 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List of Issued Books"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdib 
      Caption         =   "Book Information"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   4200
      Width           =   2295
   End
   Begin VB.TextBox txtlib 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1080
      TabIndex        =   1
      Text            =   "Enter book id"
      Top             =   3600
      Width           =   2775
   End
   Begin VB.ListBox listib 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   600
      Picture         =   "frmlistib.frx":0000
      Top             =   3600
      Width           =   405
   End
   Begin VB.Label lblqqqq 
      BackStyle       =   0  'Transparent
      Caption         =   "The list of Issued Books(Only their IDs)is below .You can know about the book details by searching Book Information."
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "frmlistib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim recordlen As Integer
Dim fn As Integer
Dim issuedt As issue
Dim recnum As Integer

Private Sub cmdib_Click()
  'loading the form search book
   frmsearchbook.Show
End Sub

Private Sub Form_Load()
  'adding the firstname of the records in the list
   fn = FreeFile
   recordlen = Len(issuedt)
   recnum = 1
    Open App.Path & "\issuerecord.txt" For Random As #fn Len = recordlen
      Do While Not EOF(fn)
       Get #fn, recnum, issuedt
       recnum = recnum + 1
       listib.AddItem issuedt.ibookid
      Loop
     Close #fn
End Sub

Private Sub listib_Click()
  'loading the information
   recordlen = Len(issuedt)
   fn = FreeFile
   Open App.Path & "\issuerecord.txt" For Random As #fn Len = recordlen
   Get #fn, listib.ListIndex + 1, issuedt
   Close #fn
End Sub

Private Sub txtlib_Change()
  'searching the required information using list
  For c = 0 To listib.ListCount - 1
   If LCase(txtlib.Text) = LCase(Left(listib.List(c), Len(txtlib.Text))) Then
    listib.Selected(c) = True
    Exit For
   End If
  Next c
End Sub

Private Sub txtlib_Click()
  'to make the text box empty when clicked
   txtlib.Text = ""
End Sub
