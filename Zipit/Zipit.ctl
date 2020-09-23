VERSION 5.00
Begin VB.UserControl Zipit 
   BackStyle       =   0  'Transparent
   ClientHeight    =   975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   960
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   975
   ScaleWidth      =   960
   ToolboxBitmap   =   "Zipit.ctx":0000
   Begin VB.Frame fra3D 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   735
      Begin VB.Image imgPic 
         Height          =   480
         Left            =   120
         Picture         =   "Zipit.ctx":0312
         Top             =   240
         Width           =   480
      End
   End
End
Attribute VB_Name = "Zipit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==============================================================================
'Richsoft Computing 2000
'Richard Southey
'This code is e-mailware, if you use it please e-mail me and tell me about
'your program.
'Please visit my website at www.geocities.com/richardsouthey.
'If you would like to make any comments/suggestions then please e-mail them to
'richardsouthey@hotmail.com.
'==============================================================================

'Zip archive collection
Public vFiles As New Collection
'Archive Filename
Private ZipFilename As String

Event OnArchiveUpdate()

Public Sub About()
    'Show the about box
    frmAbout.Show 1
End Sub

Private Sub AddEntry(zFile As ZipFile)
    Dim xFile As New ZipFileEntry
    'Adds a file from the archive into the collection
    
    xFile.Version = zFile.Version
    xFile.Flag = zFile.Flag
    xFile.CompressionMethod = zFile.CompressionMethod
    xFile.CRC32 = zFile.CRC32
    xFile.FileDateTime = GetDateTime(zFile.Date, zFile.Time)
    xFile.CompressedSize = zFile.CompressedSize
    xFile.UncompressedSize = zFile.UncompressedSize
    xFile.FileNameLength = zFile.FileNameLength
    xFile.Filename = zFile.Filename
    xFile.ExtraFieldLength = zFile.ExtraFieldLength
    
    vFiles.Add xFile
End Sub

Public Property Let Filename(New_Filename As String)
    Dim r As Long
    'Called when the filename is updated
    ZipFilename = New_Filename
    'Read in the contents of the file
    r = Read
    'Raise the update event
    RaiseEvent OnArchiveUpdate
End Property

Public Property Get Filename() As String
    'Called when the filename is read
    Filename = ZipFilename
End Property

Private Function GetDateTime(ZipDate As Integer, ZipTime As Integer) As Date
    'Converts the file date/time dos stamp from the archive
    'in to a normal date/time string
    
    Dim r As Long
    Dim FTime As FileTime
    Dim Sys As SYSTEMTIME
    Dim ZipDateStr As String
    Dim ZipTimeStr As String
    
    'Convert the dos stamp into a file time
    r = DosDateTimeToFileTime(CLng(ZipDate), CLng(ZipTime), FTime)
    'Convert the file time into a standard time
    r = FileTimeToSystemTime(FTime, Sys)

    ZipDateStr = Sys.wDay & "/" & Sys.wMonth & "/" & Sys.wYear
    ZipTimeStr = Sys.wHour & ":" & Sys.wMinute & ":" & Sys.wSecond

    GetDateTime = ZipDateStr & " " & ZipTimeStr
End Function
Public Function Read() As Long
    'Reads the archive and places each file into a collection
    Dim Sig As Long
    Dim ZipStream As Integer
    Dim Res As Long
    Dim zFile As ZipFile
    Dim Name As String
    Dim i As Integer
    
    'Check there is something to do
    If ZipFilename = "" Then
        Read = 0
        Exit Function
    End If
    
    'Clears the collection
    'begin
    'vFiles.Clear;
    For i = vFiles.Count To 1 Step -1
        vFiles.Remove i
    Next i
    
    'Opens the archive for binary access
    ZipStream = FreeFile
    Open ZipFilename For Binary As ZipStream
    'Loop through archive
    Do While True
        Get ZipStream, , Sig
        'See if the file header has been found
              If Sig = LocalFileHeaderSig Then
                    'Read each part of the file header
                    Get ZipStream, , zFile.Version
                    Get ZipStream, , zFile.Flag
                    Get ZipStream, , zFile.CompressionMethod
                    Get ZipStream, , zFile.Time
                    Get ZipStream, , zFile.Date
                    Get ZipStream, , zFile.CRC32
                    Get ZipStream, , zFile.CompressedSize
                    Get ZipStream, , zFile.UncompressedSize
                    Get ZipStream, , zFile.FileNameLength
                    Get ZipStream, , zFile.ExtraFieldLength
                    'Get the filename
                    'Set up a empty string so the right number of
                    'bytes is read
                    Name = String$(zFile.FileNameLength, " ")
                    Get ZipStream, , Name
                    zFile.Filename = Mid$(Name, 1, zFile.FileNameLength)
                    'Move on through the archive
                    'Skipping extra space, and compressed data
                    Seek ZipStream, (Seek(ZipStream) + zFile.ExtraFieldLength)
                    Seek ZipStream, (Seek(ZipStream) + zFile.CompressedSize)
                    'Add the fileinfo to the collection
                    AddEntry zFile
              Else
              Debug.Print Sig
                If Sig = CentralFileHeaderSig Or Sig = 0 Then
                    'All the filenames have been found so
                    'exit the loop
                    Exit Do
                'End
                Else
                If Sig = EndCentralDirSig Then
                    'Exit the loop
                    Exit Do
                End If
                End If
            End If
        Loop
        'Close the archive
        Close ZipStream
        'Return the number of files in the archive
        Read = vFiles.Count

    'Fire the update event
    RaiseEvent OnArchiveUpdate
End Function




Private Sub UserControl_Resize()
    'Fix the control's size
    UserControl.Size 990, 990
End Sub

