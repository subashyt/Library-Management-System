VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Library Management System"
   ClientHeight    =   10740
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   ScaleHeight     =   10740
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1200
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   52
      ImageHeight     =   49
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":10D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":16AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":412D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":6BAF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1125
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   1984
      ButtonWidth     =   1561
      ButtonHeight    =   1826
      Appearance      =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Issue"
            Description     =   "Issue Books"
            Object.ToolTipText     =   "Issue Books"
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Return"
            Description     =   "Return Books"
            Object.ToolTipText     =   "Return Books"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Book"
            Description     =   "Book Form"
            Object.ToolTipText     =   "Book Form"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Member"
            Description     =   "Member Form"
            Object.ToolTipText     =   "Member Form"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Calculator"
            Description     =   "Need a Calculator"
            Object.ToolTipText     =   "Need a Calculator"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   10200
      Width           =   20280
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   375
         Left            =   12360
         TabIndex        =   1
         Top             =   120
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   2
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   5
               TextSave        =   "10:33 PM"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   6
               TextSave        =   "9/20/2008"
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Sylfaen"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Image Image3 
      Height          =   750
      Index           =   3
      Left            =   0
      Picture         =   "mainfrm.frx":9631
      Top             =   0
      Width           =   750
   End
   Begin VB.Image Image3 
      Height          =   750
      Index           =   2
      Left            =   0
      Picture         =   "mainfrm.frx":9D2D
      Top             =   0
      Width           =   750
   End
   Begin VB.Image Image3 
      Height          =   750
      Index           =   1
      Left            =   0
      Picture         =   "mainfrm.frx":A429
      Top             =   0
      Width           =   750
   End
   Begin VB.Image Image4 
      Height          =   21600
      Left            =   -4080
      Picture         =   "mainfrm.frx":AB25
      Top             =   960
      Width           =   28800
   End
   Begin VB.Image Image3 
      Height          =   750
      Index           =   0
      Left            =   5520
      Picture         =   "mainfrm.frx":9807B
      Top             =   1920
      Width           =   750
   End
   Begin VB.Image Image2 
      Height          =   2250
      Index           =   0
      Left            =   5880
      Picture         =   "mainfrm.frx":98777
      Top             =   2160
      Width           =   2250
   End
   Begin VB.Image Image1 
      Height          =   3960
      Index           =   0
      Left            =   1440
      Picture         =   "mainfrm.frx":999A6
      Top             =   1320
      Width           =   3765
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnubook 
         Caption         =   "Book"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnumember 
         Caption         =   "Member"
         Shortcut        =   ^M
      End
      Begin VB.Menu separator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
      Begin VB.Menu mnuallbooks 
         Caption         =   "All Books"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuallmembers 
         Caption         =   "All Members"
         Shortcut        =   ^Z
      End
      Begin VB.Menu sept 
         Caption         =   "-"
      End
      Begin VB.Menu mnuib 
         Caption         =   "Issued Books"
      End
   End
   Begin VB.Menu mnutransaction 
      Caption         =   "&Transaction"
      Begin VB.Menu mnuissue 
         Caption         =   "Issue"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnureturn 
         Caption         =   "Return"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnusearch 
      Caption         =   "&Search"
      Begin VB.Menu mnubooksinfo 
         Caption         =   "Books info"
         Shortcut        =   ^J
      End
      Begin VB.Menu mnumembersinfo 
         Caption         =   "Members info"
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu mnusecurity 
      Caption         =   "&Security"
      Begin VB.Menu mnuchangepass 
         Caption         =   "Change Password"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
 'set time and date in status bar
 Dim pnl As Panel
 Set pnl = StatusBar1.Panels.Add(1, , , sbrTime)
 Set pnl = StatusBar1.Panels.Add(2, , , sbrDate)
End Sub

Private Sub mnuallbooks_Click()
 'load list of all books form
 frmlistbook.Show
End Sub

Private Sub mnuallmembers_Click()
 'load list of all members form
 frmlistmember.Show
End Sub

Private Sub mnubook_Click()
 'load book form
 frmbook.Show
End Sub

Private Sub mnubooksinfo_Click()
 'load search book form
 frmsearchbook.Show
End Sub

Private Sub mnuchangepass_Click()
 'load password setting form
 frmchangepass.Show
End Sub

Private Sub mnuexit_Click()
 'terminates the application
 End
End Sub

Private Sub mnuhelp_Click()
'this displays the form help
frmhelp.Show
End Sub

Private Sub mnuib_Click()
 'load issued books list form
 frmlistib.Show
End Sub

Private Sub mnuissue_Click()
 'load issue form
 frmissue.Show
End Sub

Private Sub mnumember_Click()
 'load member form
 frmmember.Show
End Sub

Private Sub mnumembersinfo_Click()
 'load search member form
 frmsearchmember.Show
End Sub

Private Sub mnureturn_Click()
 'load book return form
 frmreturn.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
     
     Select Case Button.Index
        'Declaring Cases for the button clicks in the toolbar
        Case 1
            'load book issue form
            frmissue.Show
        Case 2
            'load book return form
            frmreturn.Show
        Case 3
            'load book form
             frmbook.Show
        Case 4
            'load member form
            frmmember.Show
        Case 5
            'load calculator
             frmCalculator.Show
    End Select

End Sub

