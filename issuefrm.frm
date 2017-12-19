VERSION 5.00
Begin VB.Form frmissue 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Issue book"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Please enter the required input"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   240
         Top             =   2160
      End
      Begin VB.CommandButton cmdicancel 
         Caption         =   "Cancel"
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
         Left            =   3960
         TabIndex        =   8
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton cmdissue 
         Caption         =   "Issue"
         Enabled         =   0   'False
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
         Left            =   2280
         TabIndex        =   7
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox txtidate 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
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
         Left            =   2280
         TabIndex        =   6
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox txtimid 
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
         Left            =   2280
         TabIndex        =   5
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox txtibid 
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
         Left            =   2280
         TabIndex        =   4
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Issue Date"
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
         TabIndex        =   3
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Member ID"
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
         TabIndex        =   2
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Book ID"
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
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmissue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim memberdt As member
Dim bookdt As book
Dim issuedt As issue
Dim recnum As Integer
Dim fn As Integer
Dim nrec As Integer
Dim recordlen As Integer
Dim fns As Integer
Dim fnt As Integer
Dim recordlens  As Integer
Dim recordlent As Integer

Private Sub cmdicancel_Click()
  'close this window
  Unload Me
  frmmain.Show
End Sub

Private Sub cmdissue_Click()
  'saving the issue details into a text file
  Dim mrec As Integer
  fn = FreeFile
  fns = fn + 1
  fnt = fns + 1
  recordlen = Len(issuedt)
  recordlens = Len(bookdt)
  recordlent = Len(memberdt)
  Open App.Path & "\issuerecord.txt" For Random As #fn Len = recordlen
  Open App.Path & "\bookrecord.txt" For Random As #fns Len = recordlens
   nrec = (LOF(fn) / recordlen) + 1
   recnum = 1
  Do While Not EOF(fns)
   Get fns, recnum, bookdt
   mrec = 1
   Open App.Path & "\memberrecord.txt" For Random As #fnt Len = recordlent
   Do While Not EOF(fnt)
    Get fnt, mrec, memberdt
  'checking whether the book or member is registered or not
    If Trim(LCase(txtibid.Text)) = Trim(LCase(bookdt.bookid)) And Trim(LCase(txtimid.Text)) = Trim(LCase(memberdt.memid)) Then
     issuedt.ibookid = Trim(txtibid.Text)
     issuedt.imemid = Trim(txtimid.Text)
     issuedt.idate = Trim(txtidate.Text)
     Put #fn, nrec, issuedt
     'informing the user that one book is issued
     MsgBox "One book issued!", vbInformation + vbOKOnly, "Issue"
     GoTo resets
    Else
     mrec = mrec + 1
    End If
   Loop
   Close fnt
   recnum = recnum + 1
   If EOF(fns) Then
    'informing the user that the book or member is not registered
    MsgBox "Item is not registered!", vbExclamation + vbOKOnly, "Issue"
   End If
  Loop
resets:
      txtibid.Text = ""
      txtimid.Text = ""
      txtidate.Text = ""
      txtibid.SetFocus
 Close fn, fns, fnt
End Sub

Private Sub Timer1_Timer()
  'fixing the date in the text box
   txtidate.Text = Date
End Sub

Private Sub txtibid_Change()
  'make the user to type all the required fields
  If txtibid.Text <> "" And txtimid.Text <> "" Then
   cmdissue.Enabled = True
  Else
   cmdissue.Enabled = False
  End If
End Sub

Private Sub txtibid_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And cmdissue.Enabled = True Then cmdissue_Click
End Sub

Private Sub txtimid_Change()
  Call txtibid_Change
End Sub

Private Sub txtimid_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And cmdissue.Enabled = True Then cmdissue_Click
End Sub
