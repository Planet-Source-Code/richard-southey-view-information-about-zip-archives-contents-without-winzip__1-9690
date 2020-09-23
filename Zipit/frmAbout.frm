VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Richsoft Zipit"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4170
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblWebsite 
      BackStyle       =   0  'Transparent
      Caption         =   "www.geocities.com/richardsouthey"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label lblZipIt 
      BackStyle       =   0  'Transparent
      Caption         =   "ZipIt 0.9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblRichsoft 
      BackStyle       =   0  'Transparent
      Caption         =   "Richsoft"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Richsoft Computing © 2000"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Image imgPic 
      Height          =   720
      Left            =   240
      Picture         =   "frmAbout.frx":030A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
'Richsoft Computing 2000
'Richard Southey
'This code is e-mailware, if you use it please e-mail me and tell me about
'your program.
'Please visit my website at www.geocities.com/richardsouthey.
'If you would like to make any comments/suggestions then please e-mail them to
'richardsouthey@hotmail.com.
'==============================================================================

Private Sub cmdOK_Click()
    'Close the about box
    Unload Me
End Sub





