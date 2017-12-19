VERSION 5.00
Begin VB.Form frmcalculator 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caclulator   "
   ClientHeight    =   3090
   ClientLeft      =   150
   ClientTop       =   135
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNumber 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   3600
      TabIndex        =   18
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton cmdNumber 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   1920
      TabIndex        =   17
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "¬"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   16
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton cmddivide 
      Caption         =   "÷"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   15
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdmultiply 
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   14
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdplusminus 
      Caption         =   "±"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   13
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmddot 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   12
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdNumber 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   2760
      TabIndex        =   11
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton cmdNumber 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   1080
      TabIndex        =   10
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton cmdNumber 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1080
      TabIndex        =   9
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cmdNumber 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   3600
      TabIndex        =   8
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cmdminus 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton cmdplus 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cmdNumber 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton cmdNumber 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   2760
      TabIndex        =   4
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cmdNumber 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   1920
      TabIndex        =   3
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cmdNumber 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cmdequalto 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Enter"
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lbldisplay 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "frmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private operand1 As Double, operand2 As Double
Private operator As String
Private clearDisplay As Boolean, appVer As String

Private Sub cmdback_Click()
  If Len(lbldisplay.Caption) > 0 Then lbldisplay.Caption = Mid$(lbldisplay.Caption, 1, Len(lbldisplay.Caption) - 1)
End Sub

Private Sub cmddivide_Click()
    'performs division
    operand1 = Val(lbldisplay.Caption)
    operator = "/"
    lbldisplay.Caption = vbNullString
End Sub

Private Sub cmddot_Click()
   If InStr(lbldisplay.Caption, ".") Then
        Exit Sub
   Else
        lbldisplay.Caption = lbldisplay.Caption & "."
   End If
End Sub

Private Sub cmdequalto_Click()
  'to get the results of calculation in display
  Dim result As Double
    On Error GoTo there:
    operand2 = Val(lbldisplay.Caption)
    If operator = "+" Then result = operand1 + operand2
    If operator = "-" Then result = operand1 - operand2
    If operator = "*" Then result = operand1 * operand2
    If operator = "/" And operand2 <> 0 Then result = operand1 / operand2
    
    lbldisplay.Caption = result
    clearDisplay = True
  Exit Sub
there:
    Call MsgBox(Err.Description, , Err.Source)
End Sub

Private Sub cmdminus_Click()
  'performs subtracion
  operand1 = Val(lbldisplay.Caption)
    operator = "-"
    lbldisplay.Caption = vbNullString
End Sub

Private Sub cmdmultiply_Click()
  'performs multiplication
  operand1 = Val(lbldisplay.Caption)
    operator = "*"
    lbldisplay.Caption = vbNullString
End Sub

Private Sub cmdNumber_Click(Index As Integer)
    If clearDisplay Then
        lbldisplay.Caption = vbNullString
        clearDisplay = False
    End If
    lbldisplay.Caption = lbldisplay.Caption + cmdNumber(Index).Caption
End Sub

Private Sub cmdplus_Click()
  'performs addition
  operand1 = Val(lbldisplay.Caption)
    operator = "+"
    lbldisplay.Caption = vbNullString
End Sub
Private Sub cmdplusminus_Click()
  lbldisplay.Caption = -Val(lbldisplay.Caption)
End Sub
Private Sub lbldisplay_Change()
    If Len(lbldisplay.Caption) > 23 Then
        Call MsgBox("Only upto 20 digits", vbInformation + vbOKOnly, "Calculator")
        Call cmdback_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'user can perform calculations using numpad
    Select Case Chr$(KeyAscii)
    Case Is = "0":      cmdNumber_Click (0)
    Case Is = "1":      cmdNumber_Click (1)
    Case Is = "2":      cmdNumber_Click (2)
    Case Is = "3":      cmdNumber_Click (3)
    Case Is = "4":      cmdNumber_Click (4)
    Case Is = "5":      cmdNumber_Click (5)
    Case Is = "6":      cmdNumber_Click (6)
    Case Is = "7":      cmdNumber_Click (7)
    Case Is = "8":      cmdNumber_Click (8)
    Case Is = "9":      cmdNumber_Click (9)
    Case Is = "+":      cmdplus_Click
    Case Is = "-":      cmdminus_Click
    Case Is = "*":      cmdmultiply_Click
    Case Is = "/":      cmddivide_Click
    Case Is = ".":      cmddot_Click
    Case Else
        If KeyAscii = vbKeyReturn Then
                        cmdequalto_Click
       
        ElseIf KeyAscii = vbKeyBack Then
                        cmdback_Click
        End If
    End Select
End Sub
Private Sub Form_Load()
    appVer = App.Major & "." & App.Minor
    frmCalculator.KeyPreview = True
End Sub
