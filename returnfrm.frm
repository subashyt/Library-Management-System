VERSION 5.00
Begin VB.Form frmreturn 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Return book"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame returnframe 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Please enter the required input"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.CommandButton cmdrefresh 
         Caption         =   "Refresh"
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
         Left            =   1680
         TabIndex        =   13
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Timer Timer 
         Interval        =   1000
         Left            =   960
         Top             =   0
      End
      Begin VB.CommandButton cmdrcancel 
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
         Left            =   3360
         TabIndex        =   12
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton cmdreturn 
         Caption         =   "Return"
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
         Left            =   120
         TabIndex        =   11
         Top             =   4080
         Width           =   1335
      End
      Begin VB.TextBox txtdu 
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
         Left            =   1920
         TabIndex        =   10
         Top             =   3240
         Width           =   2775
      End
      Begin VB.TextBox txtrdate 
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
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1920
         TabIndex        =   9
         Top             =   2640
         Width           =   2775
      End
      Begin VB.TextBox txtridate 
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
         Left            =   1920
         TabIndex        =   8
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txtrmid 
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
         Left            =   1920
         TabIndex        =   7
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txtrbid 
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
         Left            =   1920
         TabIndex        =   6
         Top             =   600
         Width           =   2775
      End
      Begin VB.Shape Shape2 
         Height          =   1335
         Left            =   120
         Top             =   480
         Width           =   4695
      End
      Begin VB.Shape Shape1 
         Height          =   1935
         Left            =   120
         Top             =   1920
         Width           =   4695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Days Used"
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
         Index           =   3
         Left            =   240
         TabIndex        =   5
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Return Date"
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
         Left            =   240
         TabIndex        =   4
         Top             =   2640
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
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   2055
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
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   2175
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
         Left            =   240
         TabIndex        =   1
         Top             =   2040
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmreturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim issuedt As issue
Dim recnum As Long
Dim nrec As Integer
Dim recordlen As Integer

Private Sub cmdrcancel_Click()
  'close this window
  Unload Me
End Sub

Private Sub cmdrefresh_Click()
  'reload the form
  Call rrefresh
End Sub

Private Sub cmdreturn_Click()
  Dim ans As String
  recnum = 1
  recordlen = Len(issuedt)
  ans = MsgBox("This book will be returned!", vbInformation + vbOKCancel, "Return Book")
  If ans = vbOK Then
   'creating a temporary text file
   Open App.Path & "\tempo.txt" For Random As #1 Len = recordlen
   Close #2
    Open App.Path & "\issuerecord.txt" For Random As #2 Len = recordlen
     Do While Not EOF(2)
        Get #2, recnum, issuedt
        If Trim(LCase(txtrbid.Text)) = Trim(LCase(issuedt.ibookid)) Then
             recnum = recnum + 1
          Get #2, recnum, issuedt
        End If
         Put #1, recnum, issuedt
         recnum = recnum + 1
      Loop
   ElseIf ans = vbCancel Then
    GoTo dono
  End If
  Close #1, #2
  'deleting a particular record
  Kill App.Path & "\issuerecord.txt"
  'copying the details to the samr record text file
  FileCopy App.Path & "\tempo.txt", App.Path & "\issuerecord.txt"
  'deleting the temporary file
  Kill App.Path & "\tempo.txt"
  Call rrefresh
dono:
End Sub

Private Sub Timer_Timer()
  'fixing the date in the text box
  txtrdate.Text = Date
End Sub

Private Sub txtdu_Click()
  Call txtrbid_Change
  'calculating the date difference in days
  If IsDate(txtridate.Text) Then
   txtdu.Text = CDate(txtrdate.Text) - CDate(txtridate.Text)
  Else
   MsgBox "You must enter a proper date!", vbCritical, "Data error"
   txtridate.SetFocus
  End If
End Sub

Private Sub txtrbid_Change()
  'enable the return command button only when the required information is provided
  If txtrbid.Text <> "" And txtrmid.Text <> "" And txtridate.Text <> "" And txtrdate.Text <> "" And txtdu.Text <> "" Then
    cmdreturn.Enabled = True
  Else
   cmdreturn.Enabled = False
  End If
End Sub

Private Sub txtrdate_Change()
  'calling the same function of txtrbid for this text box also
  Call txtrbid_Change
End Sub

Private Sub txtridate_Click()
  Call txtrbid_Change
  'when the required information is given it automatically find the issued date from the issuerecord text file
  Dim ffindrecord As String
  Dim findrecord As String
  Dim issuedt As issue
  Dim recordlen As Integer
  Dim fn As Integer
  fn = FreeFile
  recordlen = Len(issuedt)
  ffindrecord = Trim(LCase(txtrmid.Text))
  findrecord = Trim(LCase(txtrbid.Text))
  Open App.Path & "\issuerecord.txt" For Random As fn Len = recordlen
   recnum = 0
   Do While Not EOF(fn)
   recnum = recnum + 1
   Get #fn, recnum, issuedt
   txtridate.Text = issuedt.idate
   If ffindrecord = Trim(LCase(issuedt.imemid)) And findrecord = Trim(LCase(issuedt.ibookid)) Then
    Close #fn
    Exit Sub
   Else
    If EOF(fn) Then
     txtridate.Text = ""
     MsgBox "Record not found", vbExclamation + vbOKOnly, "Issue Date"
     Call rrefresh
     txtrbid.SetFocus
    Close #fn
      Exit Sub
     End If
    End If
   Loop
 Close #fn
End Sub

Private Sub txtrmid_Change()
  Call txtrbid_Change
End Sub

Private Function rrefresh()
  'reload the form
  Unload Me
  frmreturn.Show
End Function
