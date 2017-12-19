VERSION 5.00
Begin VB.Form frmlistmember 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List of all Members"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox listmember 
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
      Height          =   1950
      Left            =   120
      TabIndex        =   12
      Top             =   3600
      Width           =   3615
   End
   Begin VB.TextBox txtlistmember 
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
      Top             =   5160
      Width           =   2775
   End
   Begin VB.Frame mf 
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   6735
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         Height          =   3255
         Left            =   120
         Top             =   240
         Width           =   6495
      End
      Begin VB.Label lbllmsex 
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
         Left            =   2160
         TabIndex        =   11
         Top             =   2880
         Width           =   4335
      End
      Begin VB.Label lbllmaddress 
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
         Left            =   2160
         TabIndex        =   10
         Top             =   2280
         Width           =   4335
      End
      Begin VB.Label lbllmphone 
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
         Left            =   2160
         TabIndex        =   9
         Top             =   1680
         Width           =   4335
      End
      Begin VB.Label lbllmname 
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
         Left            =   2160
         TabIndex        =   8
         Top             =   1080
         Width           =   4335
      End
      Begin VB.Label lbllmid 
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
         Left            =   2160
         TabIndex        =   7
         Top             =   480
         Width           =   4335
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
         Left            =   240
         TabIndex        =   5
         Top             =   2880
         Width           =   1095
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
         Left            =   240
         TabIndex        =   4
         Top             =   2280
         Width           =   1095
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
         Left            =   240
         TabIndex        =   3
         Top             =   1680
         Width           =   1695
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
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   1935
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
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Image Image1 
      Height          =   2025
      Left            =   4320
      Picture         =   "listmemfrm.frx":0000
      Top             =   3360
      Width           =   1755
   End
End
Attribute VB_Name = "frmlistmember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim memberdt As member
Dim recnum As Integer
Dim fn As Integer
Dim recordlen As Integer

Private Sub Form_Load()
  'listing the member name
  fn = FreeFile
  recordlen = Len(memberdt)
  txtlistmember.Text = "Enter a member name"
  recnum = 1
  Open App.Path & "\memberrecord.txt" For Random As #fn Len = recordlen
   Do While Not EOF(fn)
    Get #fn, recnum, memberdt
    recnum = recnum + 1
    listmember.AddItem memberdt.mname
   Loop
  Close #fn
End Sub

Private Sub listmember_Click()
  'loading the information of the selected record
  recordlen = Len(memberdt)
  fn = FreeFile
  Open App.Path & "\memberrecord.txt" For Random As #fn Len = recordlen
   Get #fn, listmember.ListIndex + 1, memberdt
  Close #fn
  If memberdt.sex = True Then
   lbllmsex.Caption = "Male"
  Else
   lbllmsex.Caption = "Female"
  End If
  lbllmid.Caption = memberdt.memid
  lbllmname.Caption = memberdt.mname
  lbllmphone.Caption = memberdt.phone
  lbllmaddress.Caption = memberdt.addre
End Sub

Private Sub txtlistmember_Click()
  'to make the box empty when clicked
  txtlistmember.Text = ""
End Sub

Private Sub txtlistmember_Change()
  'searching the required information using list
  For m = 0 To listmember.ListCount - 1
   If LCase(txtlistmember.Text) = LCase(Left(listmember.List(m), Len(txtlistmember.Text))) Then
    listmember.Selected(m) = True
    Exit For
   End If
  Next m
End Sub
