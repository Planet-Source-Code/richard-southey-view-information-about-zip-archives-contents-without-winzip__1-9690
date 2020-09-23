Attribute VB_Name = "ZipUtils"
'==============================================================================
'Richsoft Computing 2000
'Richard Southey
'This code is e-mailware, if you use it please e-mail me and tell me about
'your program.
'Please visit my website at www.geocities.com/richardsouthey.
'If you would like to make any comments/suggestions then please e-mail them to
'richardsouthey@hotmail.com.
'==============================================================================

'Zip file format type
Type ZipFile
  Version As Integer                    ': WORD;
  Flag As Integer                       ': WORD;
  CompressionMethod As Integer          ': WORD;
  Time As Integer                       ': WORD;
  Date As Integer                       ': WORD;
  CRC32 As Long                      ': Longint;
  CompressedSize As Long             ': Longint;
  UncompressedSize As Long           ': Longint;
  FileNameLength As Integer             ': WORD;
  ExtraFieldLength As Integer           ': WORD;
  Filename As String                 ': String;
End Type

'Zip file constants
Public Const LocalFileHeaderSig = &H4034B50
Public Const CentralFileHeaderSig = &H2014B50
Public Const EndCentralDirSig = &H6054B50

'File dates/times functions and types
Public Declare Function DosDateTimeToFileTime Lib "kernel32" (ByVal wFatDate As Long, ByVal wFatTime As Long, lpFileTime As FileTime) As Long
Public Type FileTime
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FileTime, lpSystemTime As SYSTEMTIME) As Long
Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type




