VERSION 5.00
Begin VB.Form frmfine 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fine Information"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdfback 
      Caption         =   "Back"
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
      Left            =   4080
      TabIndex        =   9
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
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
      Left            =   2040
      TabIndex        =   8
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmdcheck 
      Caption         =   "Check"
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
      TabIndex        =   7
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox txtfbid 
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
      Left            =   2880
      TabIndex        =   6
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox txtfmid 
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
      Left            =   2880
      TabIndex        =   5
      Top             =   360
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Fine:"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   5895
      Begin VB.Label lbldays 
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
         Left            =   2760
         TabIndex        =   11
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label lblcharge 
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
         Left            =   2760
         TabIndex        =   10
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Charged Days"
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
         Index           =   3
         Left            =   480
         TabIndex        =   4
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Charge (Rs.)"
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
         Index           =   2
         Left            =   480
         TabIndex        =   3
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Book ID(Returned)"
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
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Member ID"
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
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BorderColor     =   &H000000FF&
      Height          =   1815
      Left            =   120
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "frmfine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bookdt As book
Dim memberdt As member
Dim forfinedt As forfine
Dim recnum As Integer
Dim nrec As Integer
Dim recordlen As Integer
Dim findrecord As String
Dim ffindrecord As String


Private Sub cmdcheck_Click()
On Error GoTo again
ffindrecord = Trim(LCase(txtfmid.Text))
findrecord = Trim(LCase(txtfbid.Text))
Open App.Path & "\finerecord.txt" For Random As #1 Len = Len(forfinedt)
recnum = 0
Do While Not EOF(1)
recnum = recnum + 1
Get #1, recnum, forfinedt
lblcharge.Caption = forfinedt.fnamount
lbldays.Caption = forfinedt.dayu

If ffindrecord = Trim(LCase(forfinedt.mrid)) And findrecord = Trim(LCase(forfinedt.bkid)) Then
Close #1
Exit Sub
Else
If EOF(1) Then
MsgBox "No fine charged for this record.", vbExclamation + vbOKOnly, "Fine"
Call freset
Close #1
Exit Sub
End If
End If
Loop
Close #1
Unload Me
frmfine.Show
again:
End Sub

Private Sub cmdclear_Click()
Dim ans As String

recnum = 1
recordlen = Len(forfinedt)

ans = MsgBox("Clear this Finerecord!", vbInformation + vbOKCancel, "Clear")
If ans = vbOK Then
Open App.Path & "\ftempo.txt" For Random As #1 Len = recordlen
Close #2
    Open App.Path & "\finerecord.txt" For Random As #2 Len = recordlen
     Do While Not EOF(2)

        Get #2, recnum, forfinedt
       ''''''''''
        If Trim(LCase(txtfmid.Text)) = Trim(LCase(forfinedt.mrid)) And Trim(LCase(txtfbid.Text)) = Trim(LCase(forfinedt.bkid)) Then
             recnum = recnum + 1
          Get #2, recnum, forfinedt
        End If
         
        Put #1, recnum, forfinedt
      
             recnum = recnum + 1
     Loop
     
ElseIf ans = vbCancel Then
GoTo donothing
End If
Close #1, #2
Kill App.Path & "\finerecord.txt"
FileCopy App.Path & "\ftempo.txt", App.Path & "\finerecord.txt"
Kill App.Path & "\ftempo.txt"
donothing:
Call freset
End Sub

Private Sub cmdfback_Click()
Unload Me
frmmain.Show
End Sub

Private Function freset()
txtfmid.Text = ""
txtfbid.Text = ""
lblcharge.Caption = ""
lbldays.Caption = ""
txtfmid.SetFocus
End Function
