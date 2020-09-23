VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{E8889BD8-5764-11D4-BDD7-E09052C10310}#1.0#0"; "ZIPITCONTROL.OCX"
Begin VB.Form frmZipViewer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zip Viewer"
   ClientHeight    =   3225
   ClientLeft      =   150
   ClientTop       =   765
   ClientWidth     =   7845
   Icon            =   "frmZipViewer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   7845
   StartUpPosition =   3  'Windows Default
   Begin ZipitControl.Zipit Zipit1 
      Left            =   6360
      Top             =   2400
      _ExtentX        =   1746
      _ExtentY        =   1746
   End
   Begin MSComDlg.CommonDialog cdlZip 
      Left            =   5760
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   ".zip"
      DialogTitle     =   "Open Zip Archive"
      Filter          =   "Zip Files|*.zip"
      MaxFileSize     =   256
   End
   Begin ComctlLib.ListView lvwZip 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   4048
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Filename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Date/Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Size"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Packed"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Ratio"
         Object.Width           =   882
      EndProperty
   End
   Begin VB.Label lblWeb 
      Caption         =   "www.geocities.com/richardsouthey"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label lblRichsoft 
      BackStyle       =   0  'Transparent
      Caption         =   "Richsoft Computing 2000"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label lblFiles 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About Zipit Control"
      End
   End
End
Attribute VB_Name = "frmZipViewer"
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

Private Sub Form_Unload(Cancel As Integer)
    'Show the about form
    Zipit1.About
End Sub


Private Sub mnuAbout_Click()
    'Show the control's help
    Zipit1.About
End Sub


Private Sub mnuExit_Click()
    'Exit the program
    Unload Me
End Sub

Private Sub mnuOpen_Click()
    'Open an archive
    
    On Error Resume Next
    cdlZip.ShowOpen
    'Check if cancel was pressed
    If Err = cdlCancel Then Exit Sub
    
    Zipit1.FileName = cdlZip.FileName
End Sub


Private Sub Zipit1_OnArchiveUpdate()
    'The archive has been updated so refresh the list
    Dim itmX As ListItem
    Dim r As Long
    Dim i As Long
    Dim Files As New ZipFileEntry
    
    'Get the number of files in the archive
    r = Zipit1.vFiles.Count
    
    'Show the amount of files in the archive
    lblFiles.Caption = Format(r) & " file(s) in archive"
    
    'Clear the list
    lvwZip.ListItems.Clear
    
    'Loop through each file in the archive
    For i = 1 To r
        'Store file info in a variable for ease of use
        'because the intellisense will give help
        Set Files = Zipit1.vFiles.Item(i)
        With Files
            'Add a item to the list
            Set itmX = lvwZip.ListItems.Add(, , .FileName)
            'Add the info
            itmX.Tag = i
            itmX.SubItems(1) = .FileDateTime
            itmX.SubItems(2) = .CompressedSize
            itmX.SubItems(3) = .UncompressedSize
            'Trap div by zero
            If .UncompressedSize <> 0 Then
                itmX.SubItems(4) = Format(CInt((1 - (.CompressedSize / .UncompressedSize)) * 100)) & "%"
            Else
                itmX.SubItems(4) = "0%"
            End If
        End With
    Next i
End Sub


