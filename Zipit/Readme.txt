Richsoft Zipit Control Version 0.9
-------- ----- ------- ------- ---

This ActiveX control will read a Zip file and will return its contents.
It also returns information about the files within the archive.

The control has the following Properties:

* Filename			Holds the archive's filename
* vFiles			A collection containing the files within the archive

The control has the following Methods:

* About				Shows an About box containing information about the control
* Read				Forces the control to re-read the archive

It has the following events

* OnRefresh			Triggered when the archive contains changes
				i.e. when the filename is changed


The vFiles property contains the file info as a collection of ZipDirEntry classes.
The class is exposed by the control, and contains the following variables about the files:

* Version As Integer
* Flag As Integer
* CompressionMethod As Integer
* FileDateTime As String
* CRC32 As Long
* CompressedSize As Long
* UncompressedSize As Long
* FileNameLength As Integer
* ExtraFieldLength As Integer
* Filename As String

This property can be read directly i.e Filename = Zipit1.vFiles(1).Filename, but intellisense will not give any help.  Intead, if a new class is created and set 
to the vFiles property

i.e.		Dim Files As New ZipDirEntry
		Set Files = Zipit1.vFiles(1)
		Filename = Files.Filename

intellisense will give help.


* The Filename is returned with the path in UNIX style, i.e. uses forward slashes
* Directories are also returned if path info was stored within the archive


Limitations
-----------

* This version does not provide any compression/decompression.
* Self extracting zip files and not supported yet.


Future Plans
------ -----

* Add support for self extracting archives.
* Add archive functions such as adding, extracting and deleting files from the archive.

Contacting Me
---------- --

Website: www.geocities.com/richardsouthey
E-mail:  richardsouthey@hotmail.com