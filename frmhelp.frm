VERSION 5.00
Begin VB.Form frmhelp 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   4500
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   4440
   StartUpPosition =   3  'Windows Default
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   375
      Left            =   840
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label7 
      Caption         =   "The software also allows to view the list of all records. Such as list all books, list all members, list all issued books."
      Height          =   855
      Left            =   480
      TabIndex        =   6
      Top             =   3840
      Width           =   3615
   End
   Begin VB.Label Label6 
      Caption         =   "Issue and return the books in Issue form and Return form."
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   3240
      Width           =   3615
   End
   Begin VB.Label Label5 
      Caption         =   $"frmhelp.frx":0000
      Height          =   855
      Left            =   480
      TabIndex        =   4
      Top             =   2280
      Width           =   3735
   End
   Begin VB.Label Label4 
      Caption         =   $"frmhelp.frx":008A
      Height          =   855
      Left            =   480
      TabIndex        =   3
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "Copyright 2008 Subash Marasini"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "This application is build by Subash Marasini."
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Library Management System"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmhelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuback1_Click()
Unload Me
frmmain.Show
End Sub
