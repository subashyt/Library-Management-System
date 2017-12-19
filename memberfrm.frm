VERSION 5.00
Begin VB.Form frmmember 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Member"
   ClientHeight    =   4665
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   5400
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameopt 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   2040
      TabIndex        =   16
      Top             =   3840
      Width           =   3255
      Begin VB.OptionButton optmale 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Male"
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
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optfemale 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Female"
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
         Left            =   1680
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Timer timer 
      Interval        =   1000
      Left            =   480
      Top             =   4080
   End
   Begin VB.Frame framemember 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Find a member by ID or Name"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   13
      Top             =   4800
      Width           =   5175
      Begin VB.TextBox txtmfind 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   2655
      End
      Begin VB.CommandButton cmdmfind 
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
         Left            =   3480
         Picture         =   "memberfrm.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.TextBox txtdatev 
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
      Left            =   2040
      TabIndex        =   12
      Top             =   3360
      Width           =   3255
   End
   Begin VB.TextBox txtdatej 
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   2760
      Width           =   3255
   End
   Begin VB.TextBox txtmaddress 
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
      Left            =   2040
      TabIndex        =   10
      Top             =   2160
      Width           =   3255
   End
   Begin VB.TextBox txtmphone 
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
      Left            =   2040
      TabIndex        =   9
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox txtmname 
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
      Left            =   2040
      TabIndex        =   8
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox txtmid 
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
      Left            =   2040
      TabIndex        =   7
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Valid Till"
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
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label sexlbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
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
      Index           =   5
      Left            =   1320
      TabIndex        =   5
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Joining"
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
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No."
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
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Member Name"
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
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Member ID"
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
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Menu mnumemberfile 
      Caption         =   "&File"
      Begin VB.Menu mnumemberadd 
         Caption         =   "Add"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnumembersave 
         Caption         =   "Save changes"
         Shortcut        =   ^S
      End
      Begin VB.Menu sssppppppppp 
         Caption         =   "-"
      End
      Begin VB.Menu mnugtmf 
         Caption         =   "Back"
      End
   End
   Begin VB.Menu mnumemberedit 
      Caption         =   "&Edit"
      Begin VB.Menu mnumemberedit1 
         Caption         =   "Edit a member"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnumemberreset 
         Caption         =   "Refresh"
         Shortcut        =   ^R
      End
   End
End
Attribute VB_Name = "frmmember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim memberdt As member
Dim recnum As Integer
Dim fn As Integer
Dim nrec As Integer
Dim recordlen As Integer
Dim findrecord As String

Private Sub cmdmfind_Click()
 'searching the perticular record
 findrecord = Trim(LCase(txtmfind.Text))
 fn = FreeFile
 Open App.Path & "\memberrecord.txt" For Random As fn Len = Len(memberdt)
 recnum = 0
 Do While Not EOF(fn)
 recnum = recnum + 1
 Get #fn, recnum, memberdt
 txtmid.Text = memberdt.memid
 txtmname.Text = memberdt.mname
 txtmphone.Text = memberdt.phone
 txtmaddress.Text = memberdt.addre
 txtdatej.Text = memberdt.doj
 txtdatev.Text = memberdt.valid
 If memberdt.sex = True Then
  optmale.Value = True
 Else
  optfemale.Value = True
 End If
 If findrecord = Trim(LCase(memberdt.memid)) Or findrecord = Trim(LCase(memberdt.mname)) Then
   Close #fn
   Exit Sub
  Else
  If EOF(fn) Then
    txtdatev.Text = ""
    txtdatej.Text = ""
    MsgBox "Search not found !", vbExclamation + vbOKOnly, "Find"
    txtmfind.SetFocus
    Close #fn
    Exit Sub
   End If
 End If
 Loop
 Close #fn
End Sub

Private Sub Form_Load()
 mnumembersave.Enabled = False
 framemember.Enabled = False
End Sub

Private Sub mnugtmf_Click()
 'close this window
 Unload Me
 frmmain.Show
End Sub

Private Sub mnumemberadd_Click()
 recnum = 1
 fn = FreeFile
 recordlen = Len(memberdt)
'saving the information in the text file
 Open App.Path & "\memberrecord.txt" For Random As fn Len = recordlen
  nrec = (LOF(fn) / recordlen) + 1
  Do While Not EOF(fn)
   Get fn, recnum, memberdt
   If Trim(LCase(txtmid.Text)) = Trim(LCase(memberdt.memid)) Then
      MsgBox "This Id is already used!", vbCritical + vbOKOnly, "Error"
      txtmid.Text = ""
      txtmid.SetFocus
      Exit Sub
   Else
   End If
   recnum = recnum + 1
  Loop
  'only when all the text box are filled
  If txtmid.Text <> "" And txtmname.Text <> "" And txtmphone.Text <> "" And txtmaddress.Text <> "" And txtdatej.Text <> "" And txtdatev.Text <> "" Then
   memberdt.memid = Trim(txtmid.Text)
   memberdt.mname = Trim(txtmname.Text)
   memberdt.phone = Trim(txtmphone.Text)
   memberdt.addre = Trim(txtmaddress.Text)
   memberdt.doj = Trim(txtdatej.Text)
   memberdt.valid = Trim(txtdatev.Text)
   If optmale.Value = True Then
    memberdt.sex = True
   ElseIf optfemale.Value = True Then
    memberdt.sex = False
   End If
  Put #fn, nrec, memberdt
  Close fn
  MsgBox "Information Saved!", vbInformation + vbOKOnly, "Member"
  Else
  MsgBox "You have missed something!", vbInformation + vbOKOnly, "Error"
  Exit Sub
  End If
  Unload Me
  frmmember.Show
End Sub

Private Sub mnumemberedit1_Click()
 'enable some objects in the edit mode or disable some of them
 Unload Me
 frmmember.Show
 mnumembersave.Enabled = True
 framemember.Enabled = True
 mnumemberadd.Enabled = False
 txtmfind.SetFocus
 timer.Enabled = False
 txtdatej.Text = ""
 txtdatev.Text = ""
 frmmember.Height = 7260
End Sub

Private Sub mnumemberreset_Click()
 'reloading the form
 Unload Me
 frmmember.Show
End Sub

Private Sub mnumembersave_Click()
 recordlen = Len(memberdt)
 fn = FreeFile
 'saving the changes made by overwriting the record
 Open App.Path & "\memberrecord.txt" For Random As fn Len = recordlen
  If optmale.Value = True Then
   memberdt.sex = True
  ElseIf optfemale.Value = True Then
   memberdt.sex = False
  End If
  memberdt.memid = Trim(txtmid.Text)
  memberdt.mname = Trim(txtmname.Text)
  memberdt.phone = Trim(txtmphone.Text)
  memberdt.addre = Trim(txtmaddress.Text)
  memberdt.doj = Trim(txtdatej.Text)
  memberdt.valid = Trim(txtdatev.Text)
  Put #fn, recnum, memberdt
  MsgBox "Information Changes Saved", vbInformation + vbOKOnly, "Edit a member"
  Close #fn
  Unload Me
  frmmember.Show
End Sub

Private Sub Timer_Timer()
 'fixing the date in the text box
 txtdatej.Text = Date
End Sub

Private Sub txtdatev_Click()
 'clear the text box when clicked
 txtdatev.Text = ""
End Sub

Private Sub txtmfind_Click()
 'clear the text box when clicked
 txtmfind.Text = ""
End Sub

Private Sub txtmphone_KeyPress(KeyAscii As Integer)
 'make the user type numbers only
 If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
 Else
  KeyAscii = 0
  'say that the input data was invalid
  MsgBox "Error in Type", vbInformation + vbOKOnly, "Invalid information"
 End If
End Sub
