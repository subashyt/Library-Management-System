VERSION 5.00
Begin VB.Form frmsearchmember 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Member"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   4815
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   6735
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         Height          =   4455
         Left            =   120
         Top             =   240
         Width           =   6495
      End
      Begin VB.Label lblsmmvalid 
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
         Height          =   495
         Left            =   2400
         TabIndex        =   15
         Top             =   4080
         Width           =   4095
      End
      Begin VB.Label lblsmmsex 
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
         Height          =   495
         Left            =   2400
         TabIndex        =   14
         Top             =   3360
         Width           =   4095
      End
      Begin VB.Label lblsmmaddress 
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
         Height          =   495
         Left            =   2400
         TabIndex        =   13
         Top             =   2640
         Width           =   4095
      End
      Begin VB.Label lblsmmphone 
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
         Height          =   495
         Left            =   2400
         TabIndex        =   12
         Top             =   1920
         Width           =   4095
      End
      Begin VB.Label lblsmmname 
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
         Height          =   495
         Left            =   2400
         TabIndex        =   11
         Top             =   1200
         Width           =   4095
      End
      Begin VB.Label lblsmmid 
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
         Height          =   495
         Left            =   2400
         TabIndex        =   10
         Top             =   480
         Width           =   4095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Valid Till"
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
         Index           =   1
         Left            =   360
         TabIndex        =   7
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Member ID"
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
         Left            =   360
         TabIndex        =   6
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Member Name"
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
         Left            =   360
         TabIndex        =   5
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No."
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
         Left            =   360
         TabIndex        =   4
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Index           =   2
         Left            =   360
         TabIndex        =   3
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
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
         Index           =   3
         Left            =   360
         TabIndex        =   2
         Top             =   3480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter the Member ID/Member Name"
      BeginProperty Font 
         Name            =   "Arial"
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
      Width           =   6495
      Begin VB.CommandButton cmdsearchmember 
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         Picture         =   "membersearchfrm.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox txtsmmid 
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
         TabIndex        =   8
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404040&
      Height          =   1335
      Left            =   120
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "frmsearchmember"
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

Private Function reset()
 'to reset all text boxes
 lblsmmid.Caption = ""
 lblsmmname.Caption = ""
 lblsmmphone.Caption = ""
 lblsmmaddress.Caption = ""
 lblsmmvalid.Caption = ""
 lblsmmsex.Caption = ""
End Function

Private Sub cmdsearchmember_Click()
 'finding the member information from the records text file
  findrecord = Trim(LCase(txtsmmid.Text))
  fn = FreeFile
  Open App.Path & "\memberrecord.txt" For Random As #fn Len = Len(memberdt)
   recnum = 0
   Do While Not EOF(fn)
    recnum = recnum + 1
    Get #fn, recnum, memberdt
    If memberdt.sex = True Then
      lblsmmsex.Caption = "Male"
    Else
      lblsmmsex.Caption = "Female"
    End If
    lblsmmid.Caption = memberdt.memid
    lblsmmname.Caption = memberdt.mname
    lblsmmphone.Caption = memberdt.phone
    lblsmmaddress.Caption = memberdt.addre
    lblsmmvalid.Caption = memberdt.valid
    If findrecord = Trim(LCase(memberdt.memid)) Or findrecord = Trim(LCase(memberdt.mname)) Then
     Close #fn
     Exit Sub
    Else
     If EOF(fn) Then
       lblsmmsex.Caption = ""
       lblsmmvalid.Caption = ""
      'informing the user when the record is not found
      MsgBox "Search not found !", vbExclamation + vbOKOnly, "Search Member"
      'calling the function reset to reset all text boxes
      Call reset
      txtsmmid.SetFocus
      Close #fn
      Exit Sub
     End If
    End If
   Loop
  Close #fn
End Sub

Private Sub txtsmmid_Click()
  'clears the text box when clicked
  txtsmmid.Text = ""
End Sub

