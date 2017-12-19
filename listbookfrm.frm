VERSION 5.00
Begin VB.Form frmlistbook 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List of all Books"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox listbook 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   2220
      ItemData        =   "listbookfrm.frx":0000
      Left            =   120
      List            =   "listbookfrm.frx":0002
      TabIndex        =   12
      Top             =   3480
      Width           =   3615
   End
   Begin VB.TextBox txtlbname 
      BackColor       =   &H00FFFFC0&
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
      Left            =   3840
      TabIndex        =   6
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         Height          =   3015
         Left            =   120
         Top             =   240
         Width           =   5895
      End
      Begin VB.Label lbllbprice 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Top             =   2760
         Width           =   4095
      End
      Begin VB.Label lbllbcategory 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Top             =   2160
         Width           =   4095
      End
      Begin VB.Label lbllbauthor 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Top             =   1560
         Width           =   4095
      End
      Begin VB.Label lbllbname 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   8
         Top             =   960
         Width           =   4095
      End
      Begin VB.Label lbllbid 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Book ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Book Name"
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
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Author"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Price Rs."
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
         Index           =   2
         Left            =   240
         TabIndex        =   1
         Top             =   2760
         Width           =   1335
      End
   End
   Begin VB.Image Image1 
      Height          =   2025
      Left            =   4080
      Picture         =   "listbookfrm.frx":0004
      Top             =   3360
      Width           =   1755
   End
End
Attribute VB_Name = "frmlistbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bookdt As book
Dim recnum As Integer
Dim fn As Integer
Dim recordlen As Integer

Private Sub Form_Load()
  'listing the book name
  fn = FreeFile
  recordlen = Len(bookdt)
  txtlbname.Text = "Enter a book name"
  recnum = 1
  Open App.Path & "\bookrecord.txt" For Random As #fn Len = recordlen
   Do While Not EOF(fn)
    Get #fn, recnum, bookdt
    recnum = recnum + 1
    listbook.AddItem bookdt.bname
   Loop
 Close #fn
End Sub

Private Sub listbook_Click()
 'loading the information of the selected record
 recordlen = Len(bookdt)
 fn = FreeFile
 Open App.Path & "\bookrecord.txt" For Random As #fn Len = recordlen
  Get #fn, listbook.ListIndex + 1, bookdt
 Close #fn
  lbllbid.Caption = bookdt.bookid
  lbllbname.Caption = bookdt.bname
  lbllbauthor.Caption = bookdt.author
  lbllbcategory.Caption = bookdt.categ
  lbllbprice.Caption = bookdt.price
End Sub

Private Sub txtlbname_Click()
 'to make the text box blank while the user clicks on it
  txtlbname.Text = ""
End Sub

Private Sub txtlbname_Change()
 'searching the required information using list
  For a = 0 To listbook.ListCount - 1
  If LCase(txtlbname.Text) = LCase(Left(listbook.List(a), Len(txtlbname.Text))) Then
   listbook.Selected(a) = True
   Exit For
  End If
 Next a
End Sub
