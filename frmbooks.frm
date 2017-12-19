VERSION 5.00
Begin VB.Form frmbook 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Book"
   ClientHeight    =   3285
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5220
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtbprice 
      BackColor       =   &H00FFFFC0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   2760
      Width           =   3255
   End
   Begin VB.TextBox txtbcategory 
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
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Frame framebook 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Find a book by ID or Name"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   4935
      Begin VB.TextBox txtbfind 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   2655
      End
      Begin VB.CommandButton cmdbfind 
         Caption         =   "Find"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3240
         Picture         =   "frmbooks.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.TextBox txtbauthor 
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
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox txtbname 
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
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox txtbid 
      BackColor       =   &H00FFFFC0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Price Rs."
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Author"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   700
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Book Name"
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
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Book ID"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.Menu mnubfile 
      Caption         =   "&File"
      Begin VB.Menu mnubadd 
         Caption         =   "Add"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnubsave 
         Caption         =   "Save changes"
         Shortcut        =   ^S
      End
      Begin VB.Menu separatorrs 
         Caption         =   "-"
      End
      Begin VB.Menu mnugotomain 
         Caption         =   "Back"
      End
   End
   Begin VB.Menu mnubedit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuabedit 
         Caption         =   "Edit a book"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnubreset 
         Caption         =   "Refresh"
         Shortcut        =   ^R
      End
   End
End
Attribute VB_Name = "frmbook"
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

Private Sub cmdbfind_Click()
'Finding the record details from the record's text file
 findrecord = Trim(LCase(txtbfind.Text))
 fn = FreeFile
 Open App.Path & "\bookrecord.txt" For Random As fn Len = Len(bookdt)
 recnum = 0
 Do While Not EOF(fn)
   recnum = recnum + 1
 Get #fn, recnum, bookdt
 txtbid.Text = bookdt.bookid
 txtbname.Text = bookdt.bname
 txtbauthor.Text = bookdt.author
 txtbcategory.Text = bookdt.categ
 txtbprice.Text = bookdt.price
 If findrecord = Trim(LCase(bookdt.bookid)) Or findrecord = Trim(LCase(bookdt.bname)) Then
   Close #fn
   Exit Sub
   Else
   If EOF(fn) Then
     txtbprice.Text = ""
    'To inform the user that the record could not be found
     MsgBox "Search not found !", vbExclamation + vbOKOnly, "Find"
     txtbfind.SetFocus
     Close #fn
     Exit Sub
   End If
 End If
 Loop
 Close #fn
End Sub

Private Sub Form_Load()
  mnubsave.Enabled = False
  framebook.Enabled = False
End Sub

Private Sub mnuabedit_Click()
'allowing the user to use some menus only in edit mode
 Unload Me
 frmbook.Show
 mnubadd.Enabled = False
 framebook.Enabled = True
 mnubsave.Enabled = True
 frmbook.Height = 5805
 txtbfind.SetFocus
End Sub

Private Sub mnubadd_Click()
 recnum = 1
 fn = FreeFile
 recordlen = Len(bookdt)
 'saving all the book information into a text file
 Open App.Path & "\bookrecord.txt" For Random As #fn Len = recordlen
 nrec = (LOF(fn) / recordlen) + 1
 Do While Not EOF(fn)
  Get fn, recnum, bookdt
  'informing user  if any information already exists
  If Trim(LCase(txtbid.Text)) = Trim(LCase(bookdt.bookid)) Then
            MsgBox "This ID is already used!", vbCritical + vbOKCancel, "Error"
            txtbid.SetFocus
            txtbid.Text = ""
            Exit Sub
          ElseIf Trim(LCase(txtbname.Text)) = Trim(LCase(bookdt.bname)) Then
            MsgBox "Duplicate Book Name!", vbCritical + vbOKCancel, "Error"
            txtbname.SetFocus
            txtbname.Text = ""
            Exit Sub
   End If
   recnum = recnum + 1
  Loop
  'incase of saving blank information
 If txtbid.Text <> "" And txtbname.Text <> "" And txtbauthor.Text <> "" And txtbcategory.Text <> "" And txtbprice.Text <> "" Then
  bookdt.bookid = Trim(txtbid.Text)
  bookdt.bname = Trim(txtbname.Text)
  bookdt.author = Trim(txtbauthor.Text)
  bookdt.categ = Trim(txtbcategory.Text)
  bookdt.price = Trim(txtbprice.Text)
  Put #fn, nrec, bookdt
  Close fn
 'informing user that the record has been saved
  MsgBox "Information Saved!", vbInformation + vbOKOnly, "Book"
  Else
  MsgBox "You have missed something!", vbInformation + vbOKOnly, "Error"
  Exit Sub
 End If
 'reload the form after saving the record
 Unload Me
 frmbook.Show
End Sub

Private Sub mnubreset_Click()
 Unload Me
 frmbook.Show
End Sub

Private Sub mnubsave_Click()
 recordlen = Len(bookdt)
 fn = FreeFile
'save the changes that have made by overwriting the record
 Open App.Path & "\bookrecord.txt" For Random As #fn Len = recordlen
 bookdt.bookid = Trim(txtbid.Text)
 bookdt.bname = Trim(txtbname.Text)
 bookdt.author = Trim(txtbauthor.Text)
 bookdt.categ = Trim(txtbcategory.Text)
 bookdt.price = Trim(txtbprice.Text)
 Put #fn, recnum, bookdt
 MsgBox "Information Changes Saved", vbInformation + vbOKOnly, "Edit a book"
 Close #fn
 Unload Me
 frmbook.Show
End Sub

Private Sub mnugotomain_Click()
 'loading the main form
 Unload Me
 frmmain.Show
End Sub

Private Sub txtbfind_Click()
 txtbfind.Text = ""
End Sub

Private Sub txtbprice_KeyPress(KeyAscii As Integer)
 'informing user to enter the valid information
 If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
  Else
  KeyAscii = 0
  MsgBox "Error in Type", vbInformation + vbOKOnly, "Invalid information"
 End If
End Sub




