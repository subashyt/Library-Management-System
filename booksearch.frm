VERSION 5.00
Begin VB.Form frmsearchbook 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Book"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   6615
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         Height          =   4215
         Left            =   120
         Top             =   240
         Width           =   6375
      End
      Begin VB.Label lblsbprice 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   13
         Top             =   3720
         Width           =   4215
      End
      Begin VB.Label lblsbcategory 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   12
         Top             =   2880
         Width           =   4215
      End
      Begin VB.Label lblsbauthor 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   11
         Top             =   2040
         Width           =   4215
      End
      Begin VB.Label lblsbname 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   10
         Top             =   1200
         Width           =   4215
      End
      Begin VB.Label lblsbid 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   9
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Price Rs"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   360
         TabIndex        =   6
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Author"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Book Name"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Book ID"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter the Book ID/Book Name"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6375
      Begin VB.CommandButton cmdsearchbook 
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         Picture         =   "booksearch.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtsbid 
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
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.Shape Shape2 
      Height          =   1335
      Left            =   120
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "frmsearchbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bookdt As book
Dim recnum As Integer
Dim fn As Integer
Dim nrec As Integer
Dim recordlen As Integer
Dim findrecord As String

Private Sub cmdsearchbook_Click()
  'finding the book information from the records text file
  findrecord = Trim(LCase(txtsbid.Text))
  fn = FreeFile
  Open App.Path & "\bookrecord.txt" For Random As fn Len = Len(bookdt)
   recnum = 0
   Do While Not EOF(fn)
    recnum = recnum + 1
    Get #fn, recnum, bookdt
    lblsbid.Caption = bookdt.bookid
    lblsbname.Caption = bookdt.bname
    lblsbauthor.Caption = bookdt.author
    lblsbcategory.Caption = bookdt.categ
    lblsbprice.Caption = bookdt.price
    If findrecord = Trim(LCase(bookdt.bookid)) Or findrecord = Trim(LCase(bookdt.bname)) Then
      Close #fn
      Exit Sub
    Else
     If EOF(fn) Then
      lblsbprice.Caption = ""
      'informing user when book details is not found
      MsgBox "Search not found !", vbExclamation + vbOKOnly, "Search Book"
      txtsbid.SetFocus
      lblsbprice.Caption = ""
    Close #fn
      Exit Sub
     End If
    End If
   Loop
  Close #fn
End Sub

Private Sub txtsbid_Click()
  'clear the text box when clicked
  txtsbid.Text = ""
End Sub
