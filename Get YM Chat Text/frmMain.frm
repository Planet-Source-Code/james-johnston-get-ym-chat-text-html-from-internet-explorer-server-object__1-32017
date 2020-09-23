VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMain 
   Caption         =   "Get YM! Chat Text and HTML Demo"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7920
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser wbHTML 
      Height          =   2625
      Left            =   90
      TabIndex        =   5
      Top             =   3375
      Width           =   7710
      ExtentX         =   13600
      ExtentY         =   4630
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get YM! Chat Text && HTML"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   6120
      Width           =   2355
   End
   Begin VB.TextBox txtText 
      Height          =   2625
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   315
      Width           =   7710
   End
   Begin VB.Label lblStatus 
      Height          =   285
      Left            =   90
      TabIndex        =   4
      Top             =   6165
      Width           =   5190
   End
   Begin VB.Label Label2 
      Caption         =   "YM Chat HTML"
      Height          =   195
      Left            =   135
      TabIndex        =   2
      Top             =   3150
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "YM Chat Text"
      Height          =   195
      Left            =   135
      TabIndex        =   1
      Top             =   90
      Width           =   1590
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call GetYMChatText
    txtText.Text = YMText
    'Initialize Web Browser Control
    If YMHTML <> "" Then
        wbHTML.navigate "about:blank"
        While wbHTML.document.body.innerText <> ""
            DoEvents
        Wend
        wbHTML.document.write Replace(YMHTML, "onscroll=$HandleScroll()", "")
        lblStatus.Caption = ""
    Else
        lblStatus.Caption = "Couldn't get Text.  Is the YM! Chat window open?"
    End If
End Sub

Private Sub Form_Load()
    wbHTML.navigate "about:blank"
    While wbHTML.document Is Nothing
        DoEvents
    Wend
    wbHTML.document.write "."
End Sub
