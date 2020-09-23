VERSION 5.00
Begin VB.UserControl zLibControl 
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1755
   ScaleHeight     =   360
   ScaleWidth      =   1755
End
Attribute VB_Name = "zLibControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************************'
'*                        zLibControl by Jonne Lehtinen                              *'
'*                                                                                   *'
'*  What does it do?                                                                 *'
'*     - Compress multiple files.                                                    *'
'*     - Decompress.                                                                 *'
'*     - Calculate CRC32 of a file.                                                  *'
'*                                                                                   *'
'*  What does it need to run?                                                        *'
'*     - zlib.dll in windows\system(32) or the app.path directory.                   *'
'*     - Windows 9x/ME/2k/XP.                                                        *'
'*     - Visual Basic (only VB6 tested by the author).                               *'
'*                                                                                   *'
'*  Is it free?                                                                      *'
'*     Yes, it's complitely free to use in any kind of an application.               *'
'*     You may also modify the code to your needs, but                               *'
'*     it's not allowed to claim that you wrote the code complitely.                 *'
'*                                                                                   *'
'*  Copyright,                                                                       *'
'*     - zlib: Copyright (C) 1995-2002 Jean-loup Gailly and Mark Adler               *'
'*             http://www.gzip.org/zlib/zlib_license.html                            *'
'*     - this code: Copyright (C) 2003 Jonne Lehtinen                                *'
'*                                                                                   *'
'*  Notes:                                                                           *'
'*     - Progressbar code originally by David Crowell (www.davidcrowell.com),        *'
'*       Modified to the needs of this project by the author.                        *'
'*     - This code is provided 'as-is', without any express or implied               *'
'*       warranty.  In no event will the author be held liable for any damages       *'
'*       arising from the use of this code.                                          *'
'*     - If you close your application with End Statement, the files added won't be  *'
'*       removed (aka. memory leak). So use Unload Form statement instead.           *'
'*       UserControl_Terminate will be called after the unload event of the form it  *'
'*       lies in.                                                                    *'
'*                                                                                   *'
'*  Known Limitations:                                                               *'
'*     - Does not support files over 2GB.                                            *'
'*************************************************************************************'

Option Explicit

Public Event Click()

' DLL Function Call Declarations
' Read zlib.h (available from www.gzip.org/zlib) for more info about the zlib.dll functions
Private Declare Function ZLibCompress Lib "zlib.dll" Alias "compress2" ( _
     dest As Byte, _
     destLen As Long, _
     source As Byte, _
     ByVal sourceLen As Long, _
     ByVal level As Integer) As Long

Private Declare Function ZlibUncompress Lib "zlib.dll" Alias "uncompress" ( _
     dest As Byte, _
     destLen As Long, _
     source As Byte, _
     ByVal sourceLen As Long) As Long

Private Declare Function ZlibCrc32 Lib "zlib.dll" Alias "crc32" ( _
     ByVal crc As Long, buf As Byte, _
     ByVal buf_len As Long) As Long

' W32API Function Call Declarations
Private Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" ( _
     ByVal nBufferLength As Long, _
     ByVal lpBuffer As String) As Long
     
Private Declare Function GetTempFileName Lib "kernel32.dll" Alias "GetTempFileNameA" ( _
     ByVal lpszPath As String, _
     ByVal lpPrefixString As String, _
     ByVal wUnique As Long, _
     ByVal lpTempFileName As String) As Long

' Constants
Private Const MAX_PATH As Integer = 260
Private Const Header As String = "TaSSuSoft zLibControl"
Private Const TempFilePrefix As String * 3 = "zLi"

' Filedata type, used to save filedata (*duh* ;p)
Private Type FileData
    Path As String
    filename As String
    Original_Crc32 As Long
    CompressedFile_crc32 As Long
    raw_data As Long
    comp_data As Long
    TempName As String
End Type

' just some variables
Private mTempPath As String
Private Files() As FileData
Private mFiles As Long
Private mReadBuf As Long

' progressbar variables
Private mBackColor As Long
Private mBarColor As Long
Private mVertical As Boolean
Private mMin As Long
Private mMax As Long
Private mValue As Long

'*************************************************************************************'
'*   Function Calculate_Crc32                                                        *'
'*     Calculates the CRC32 of a file                                                *'
'*                                                                                   *'
'*   Arguments:                                                                      *'
'*     filename - full path and name of the file from which to calculate the crc32.  *'
'*                                                                                   *'
'*   Return Values:                                                                  *'
'*     CRC32 - calculated CRC32                                                      *'
'*                                                                                   *'
'*   Note: There's no actual error return value, but most likely there was an error  *'
'*         if the function returns 0.                                                *'
'*************************************************************************************'

Public Function Calculate_Crc32(ByVal filename As String) As Long
    ' Check that filename exists and it's not a directory
    If Dir(filename) <> "" And (GetAttr(filename) And vbDirectory) <> vbDirectory Then
        Dim Buffer() As Byte
        Dim BufferLen As Long
        Dim crc As Long
        Dim blocks As Long
        Dim FileNum As Integer
        
        ' set up variables, initialize buffer
        FileNum = FreeFile
        BufferLen = mReadBuf
        ReDim Buffer(mReadBuf - 1)
        ' Initial CRC32 value returned by ZlibCRC32.
        ' This is the same as: crc = ZlibCRC32(0&, cbyte(0), 0&)
        crc = 0
        
        ' Open file for read-only access to prevent accidental writing into the file.
        Open filename For Binary Access Read As #FileNum
        
            ' How many times to loop to read the whole file
            blocks = ((LOF(FileNum) + mReadBuf - 1) \ mReadBuf)
            
            Do While blocks > 0
                
                If blocks = 1 Then
                    ' ok, so we're almost at the end of the file here,
                    ' so what we need to do is to read the rest of the file.
                    ' Determine amount remaining bytes.
                    BufferLen = LOF(FileNum) - Loc(FileNum)
                    If BufferLen <> 0 Then
                        ' Resize buffer to hold the remaining bytes.
                        ReDim Buffer(BufferLen - 1)
                    Else
                        ' whoops, we're already at the end of file.
                        Exit Do
                    End If
                End If
                
                ' Read bytes into the buffer array
                Get #FileNum, , Buffer()
                ' calculate crc32
                crc = ZlibCrc32(crc, Buffer(0), BufferLen)
                
                ' Progressbar control
                value = value + BufferLen

                blocks = blocks - 1
            Loop
            
        Close #FileNum
        
        Erase Buffer    ' free buffer from memory
        
        ' return crc32
        Calculate_Crc32 = crc
    End If
End Function

'*************************************************************************************'
'*   Function Compress()                                                             *'
'*     Compresses files                                                              *'
'*                                                                                   *'
'*   Arguments:                                                                      *'
'*     filename - path and filename of the file to be created.                       *'
'*     OverWrite - Overwrite filename if it already exists.                          *'
'*     level - Level of compression (1-9), might not have any effect.                *'
'*                                                                                   *'
'*   Return values:                                                                  *'
'*      0, Everything OK.                                                            *'
'*     -2, Invalid level argument,                                                   *'
'*           there's an error handler for this, so it _shouldn't_ occur.             *'
'*     -4, Not enough memory, this terminates the process immediately.               *'
'*     -5, Not large enough buffer, handled by the function,                         *'
'*           but might still be returned.                                            *'
'*     -6, Invalid zlib.dll version.                                                 *'
'*     -7, No files to compress.                                                     *'
'*     -8, filename already existed and overwrite argument was false.                *'
'*     -9, filename already existed as a directory on the system.                    *'
'*                                                                                   *'
'*   Note: Error numbers -1 and -3 shouldn't occur, but if they do,                  *'
'*         there's something terribly wrong.                                         *'
'*************************************************************************************'

Public Function Compress(ByVal filename As String, Optional ByVal Overwrite As Boolean = True, Optional ByVal level As Integer = 6) As Integer
    ' check that the user really has added something to compress.
    If mFiles > -1 Then
        
        ' check the existance of the filename and remove if the user allows removal.
        If Dir(filename) <> "" Then
            If (GetAttr(filename) And vbDirectory) <> vbDirectory Then
                If Overwrite Then
                    'removal allowed
                    Kill filename
                Else
                    'removal not allowed, return error code and exit function.
                    Compress = -8
                    Exit Function
                End If
            Else
                ' Eh! The filename exists but it's a directory!
                ' We don't want to handle directory deletions in this function
                ' since it might be unsafe
                Compress = -9
                Exit Function
            End If
        End If
        
        ' Variables to hold the file number
        Dim FileNum1 As Integer, FileNum2 As Integer, filenum3 As Integer
        
        ' Variables used in the compression algorithm.
        ' The arrays are bufferes which are passed to the ZlibCompress function.
        ' comp_len's the size of comp_bytes().
        ' BytesLeft is the size of raw_bytes() (gah, bad name :o).
        ' ZlibReturn is used to hold the return value of the ZlibCompress function
        ' and to handle errors if one occurs.
        Dim raw_bytes() As Byte, comp_bytes() As Byte
        Dim comp_len As Long, BytesLeft As Long, ZlibReturn As Integer
        
        ' Just some variables
        Dim blocks As Long
        Dim crc As Long
        Dim loopcount As Byte
        Dim i As Long

        ' open filename
        filenum3 = FreeFile
        Open filename For Binary Access Write As #filenum3
        
            ' Write archive main header
            ' Lenght of the headertext
            Put #filenum3, , Len(Header)
            ' Headertext
            Put #filenum3, , Header
            ' number of compressed files in archive
            Put #filenum3, , mFiles
            
            ' Loop through each file and compress
            For i = 0 To mFiles
                
                If Dir(Files(i).Path) <> "" And Files(i).Path <> "" Then
                If (GetAttr(Files(i).Path) & vbDirectory) <> vbDirectory Then
                    ' file exists and isn't a directory
                    
                    ' Initialize variables
                    crc = 0
                    value = 0
                    Max = Files(i).raw_data
                    FileNum1 = FreeFile
                    BytesLeft = mReadBuf
                    
                    ' Open temporary file
                    Open Files(i).TempName For Binary As #FileNum1
                    
                        FileNum2 = FreeFile
                        ' Open the actually file to compress (Read-only)
                        Open Files(i).Path For Binary Access Read As #FileNum2
                        
                            ' Count the amount of loops to do to read the whole file
                            blocks = ((LOF(FileNum2) + mReadBuf - 1) \ mReadBuf)
                            
                            Do While blocks > 0
                            
                                ' if the current loop's supposed to be the last one, then
                                ' calculate the remaining amount of bytes to read
                                If blocks = 1 Then BytesLeft = LOF(FileNum2) - Loc(FileNum2)
                                If BytesLeft = 0 Then Exit Do
                                
                                ' zLib.h comments says that the output buffer should be at least 0.1% larger than
                                ' the input buffer, so let it be so
                                comp_len = BytesLeft + BytesLeft / 1000 + 12
                                ' resize buffers
                                ReDim comp_bytes(comp_len - 1)
                                ReDim raw_bytes(BytesLeft - 1)
                                
                                ' read bytes from file to buffer
                                Get #FileNum2, , raw_bytes()
                                
                                ' compress
                                ZlibReturn = ZLibCompress(comp_bytes(0), comp_len, raw_bytes(0), BytesLeft, level)
                                
                                ' compression error handler
                                loopcount = 0
                                Do While ZlibReturn <> 0
                                
                                    Select Case ZlibReturn
                                        Case -2:    ' Invalid level parameter
                                            level = 6
                                            ZlibReturn = ZLibCompress(comp_bytes(0), comp_len, raw_bytes(0), BytesLeft, level)
                                        Case -4:    ' Not enough memory
                                            loopcount = 11      ' there's nothing we could do :/
                                        Case -5:    ' Too small output buffer
                                            ' let's double the buffer
                                            comp_len = 2 * comp_len
                                            ReDim comp_bytes(comp_len)
                                            ' and try again
                                            ZlibReturn = ZLibCompress(comp_bytes(0), comp_len, raw_bytes(0), BytesLeft, level)
                                        Case Else:  ' Unidentified error
                                            loopcount = 11
                                    End Select
                                    
                                    ' tries to solve the problem a few times, then return an error code if not successful
                                    If loopcount >= 10 And ZlibReturn <> 0 Then
                                        
                                        ' fatal error was encountered and could not be corrected
                                        ' close filehandles
                                        Close #FileNum2
                                        Close #FileNum1
                                        Close #filenum3
                                        
                                        ' Remove buffers from memory
                                        Erase raw_bytes
                                        Erase comp_bytes
                                        
                                        ' Return errorcode
                                        Compress = ZlibReturn
                                        Exit Function
                                        
                                    Else
                                        loopcount = loopcount + 1
                                    End If
                                    
                                Loop
                                
                                ' remove the garbage data from the output buffer
                                ReDim Preserve comp_bytes(comp_len - 1)
                                
                                ' count the crc for it
                                crc = ZlibCrc32(crc, comp_bytes(0), comp_len)
                                
                                ' write buffer sizes and the output buffer
                                Put #FileNum1, , BytesLeft
                                Put #FileNum1, , comp_len
                                Put #FileNum1, , comp_bytes()
                                
                                ' Progressbar control
                                value = value + mReadBuf
                                
                                blocks = blocks - 1
                                
                            Loop
                        Close #FileNum2
                        
                        ' Return to the beginning of the output file
                        Seek #FileNum1, 1
                        
                        ' Save some data of the output file
                        Files(i).CompressedFile_crc32 = crc
                        Files(i).comp_data = LOF(FileNum1)
                        
                        ' Reset progressbar
                        value = 0
                        Max = Files(i).comp_data + 6
                        
                        ' Build File header
                        Put #filenum3, , Len(Files(i).filename)
                        Put #filenum3, , Files(i).filename
                        Put #filenum3, , Files(i).raw_data
                        Put #filenum3, , Files(i).Original_Crc32
                        Put #filenum3, , Files(i).CompressedFile_crc32
                        Put #filenum3, , Files(i).comp_data
                        
                        value = value + 6
                        
                        ' Count blocks
                        blocks = ((LOF(FileNum1) + mReadBuf - 1) \ mReadBuf)
                        
                        ' resize buffer
                        BytesLeft = mReadBuf
                        ReDim raw_bytes(BytesLeft)
                        
                        Do While blocks > 0
                            
                            If blocks = 1 Then
                                ' last block
                                BytesLeft = LOF(FileNum1) - Loc(FileNum1)
                                If BytesLeft > 0 Then
                                    ReDim raw_bytes(BytesLeft - 1)
                                Else
                                    Exit Do
                                End If
                            End If
                            
                            ' copy bytes
                            Get #FileNum1, , raw_bytes()
                            Put #filenum3, , raw_bytes()
                            
                            ' progressbar control
                            value = value + BytesLeft
                            
                            blocks = blocks - 1
                        
                        Loop
                        
                    Close #FileNum1
                    
                    ' remove temporary file
                    Kill Files(i).TempName
                End If
                End If  ' file exists and isn't a directory - end
            Next i
        Close #filenum3
        
        ' return 0 to indicate a success
        Compress = 0
        
        ' Remove buffers from memory
        Erase raw_bytes
        Erase comp_bytes
        
    Else
        ' Error, no files to compress
        Compress = -7
    End If
End Function

' The following code needs to be rewritten

Private Function CheckDir(ByVal directory As String, ByVal outputpath As String, ByVal UsePaths As Boolean, ByVal Overwrite As Boolean) As String
    Dim temp As String, i As Long
    If InStr(directory, ":") = 2 Then directory = Right(directory, Len(directory) - 2) ' if, for some odd reason, the Driveletter is still there, remove it
    Do While Mid(outputpath, Len(outputpath)) = "\"
        outputpath = Left(outputpath, Len(outputpath) - 1)
    Loop ' the outputpath shouldn't have "\" as the last character, remove it if there is one
    If InStr(outputpath, ":") <> 2 Then         ' The outputpath should also have drive
        If InStr(outputpath, "\") <> 1 Then     ' if it isn't, assume "C:"
            outputpath = "C:\" & outputpath     ' and add it there
        Else
            outputpath = "C:" & outputpath
        End If
    End If
    If Not UsePaths Then
        i = Len(directory)
        Do While Mid(directory, i, 1) <> "\"
            i = i - 1
        Loop
        directory = Mid(directory, i)
    End If
    If Dir(outputpath & directory) = "" Then        ' if the file doesn't exist
        i = 4
        Do                                      ' make all directories which are missing but required
            DoEvents
            temp = Mid(outputpath & directory, 1, i)
            i = i + 1
            If Mid(temp, Len(temp)) = "\" And Dir(temp, vbDirectory) = "" Then If Dir(temp) = "" Then MkDir temp
            If temp = outputpath & directory Then Exit Do
        Loop
    Else    ' if the file exists
        If Not Overwrite Then ' prevent overwrite if the users doesn't want to overwrite the file
            outputpath = ""
            directory = ""
        Else ' destroy the file to overwrite, the client's always right
            Kill outputpath & directory
        End If
    End If
    CheckDir = outputpath & directory
End Function

'*************************************************************************************'
'*   Function Decompress()                                                           *'
'*     Decompresses files.                                                           *'
'*                                                                                   *'
'*     Arguments:                                                                    *'
'*       filename - path and filename of the file to decompress.                     *'
'*       outputpath - Optional, directory where to extract the files.                *'
'*       GetFileInfoOnly - Optional, tells the function to only get the file names,  *'
'*                         sizes, crc32, etc. Does not decompress.                   *'
'*       CheckCRC32 - Optional, Function checks archive for corruptions.             *'
'*       DoubleCRC32 - Optional, Checks both, the compressed data crc32 and          *'
'*                     decompressed file crc32.                                      *'
'*                     If false, only original data crc32 checked.                   *'
'*       SaveCorruptedFiles - Optional, saves files even if crc check doesn't match. *'
'*       UsePaths - Optional, Use relative paths of the files in the archive.        *'
'*       OverWrite - Optional, Overwrite by default.                                 *'
'*                   Otherwise existing files are skipped                            *'
'*                                                                                   *'
'*     Return Values:                                                                *'
'*        0, Everything OK                                                           *'
'*       -3, Compressed data was corrupted and there was nothing could be done.      *'
'*       -4, Not enough memory, process terminated.                                  *'
'*       -5, Too small buffer, process tried to make it larger but it was a failure. *'
'*       -7, Invalid archive header.                                                 *'
'*       -8, outputpath already exists but it's not a directory.                     *'
'*       -9, filename was not found or it is a directory.                            *'
'*                                                                                   *'
'*     Note: Other return values shouldn't occur, but if they do,                    *'
'*           there's something terribly wrong.                                       *'
'*************************************************************************************'

Public Function Decompress( _
    ByVal filename As String, _
    Optional ByVal outputpath As String = "C:\", _
    Optional ByVal GetFileInfoOnly As Boolean = False, _
    Optional ByVal CheckCRC32 As Boolean = True, _
    Optional ByVal DoubleCRC32 As Boolean = False, _
    Optional ByVal SaveCorruptedFiles As Boolean = True, _
    Optional ByVal UsePaths As Boolean = True, _
    Optional ByVal Overwrite As Boolean = True) As Integer

    ' if file exists and is not a directory, then
    If Dir(filename) <> "" And (GetAttr(filename) And vbDirectory) <> vbDirectory Then
        
        Dim FileNum1 As Integer
        FileNum1 = FreeFile
        
        ' Open the archive
        Open filename For Binary Access Read As #FileNum1
            
            Dim HeaderLen As Long, FileHeader As String
            Max = LOF(FileNum1)
            
            ' Read the header
            Get #FileNum1, , HeaderLen
            FileHeader = Space(HeaderLen)
            Get #FileNum1, , FileHeader
            
            If FileHeader <> Header Then
            
                ' Invalid header, return error
                Decompress = -7
                Exit Function
                
            End If
            
            ' Read how many files compressed in this file
            Get #FileNum1, , mFiles
            ReDim Files(mFiles)
            
            ' Update progressbar
            value = Loc(FileNum1)
            
            Dim i As Long
            Dim FilenameLen As Long
            ' loop through all files
            For i = 0 To mFiles
            
                With Files(i)
                    
                    ' Read fileheader
                    Get #FileNum1, , FilenameLen
                    .filename = Input(FilenameLen, #FileNum1)
                    Get #FileNum1, , .raw_data
                    Get #FileNum1, , .Original_Crc32
                    Get #FileNum1, , .CompressedFile_crc32
                    Get #FileNum1, , .comp_data
                    
                    ' update progressbar
                    value = Loc(FileNum1)
                    
                    If .comp_data > 0 Then
                        If GetFileInfoOnly Then
                            ' the user wanted only information on the files
                            ' (size, crc32, filename, etc), so let's skip the
                            ' actual file data
                            Seek #FileNum1, Loc(FileNum1) + .comp_data + 1
                            
                            ' update progressbar
                            value = Loc(FileNum1)
                        Else
                            ' Decompress code
                            ' check and create any missing directories, etc
                            ' returns "" if overwrite = false and the file already exists
                            Dim ChkDir As String
                            ChkDir = CheckDir(.filename, outputpath, UsePaths, Overwrite)
                            
                            If ChkDir <> "" Then
                                
                                Dim crc As Long, raw_bytes() As Byte, decomp_bytes() As Byte
                                Dim ZlibReturn As Long, BytesLeft As Long
                                Dim FileNum2 As Integer, loopcount As Byte, crc2 As Long
                                Dim Decomp_Len As Long, BytesToRead As Long
                                
                                ' get new TempFileName
                                .TempName = NewTempFileName
                                
                                ' initialize variables
                                BytesLeft = .comp_data
                                BytesToRead = -8
                                crc = 0
                                crc2 = 0
                                
                                ' Open the tempfile
                                FileNum2 = FreeFile
                                Open .TempName For Binary Access Write As #FileNum2
                                    Do
                                    
                                        ' bytes left to read
                                        BytesLeft = BytesLeft - BytesToRead - 8
                                        If BytesLeft <= 0 Then Exit Do
                                        
                                        ' get decompressed buffer lenght and the coresponding compressed data lenght
                                        Get #FileNum1, , Decomp_Len
                                        Get #FileNum1, , BytesToRead
                                        
                                        ' resize buffers
                                        ReDim raw_bytes(BytesToRead - 1)
                                        ReDim decomp_bytes(Decomp_Len - 1)
                                        
                                        ' read
                                        Get #FileNum1, , raw_bytes()
                                        
                                        ' decompress
                                        ZlibReturn = ZlibUncompress(decomp_bytes(0), Decomp_Len, raw_bytes(0), BytesToRead)
                                        
                                        ' error handler
                                        loopcount = 0
                                        Do While ZlibReturn <> 0
    
                                            Select Case ZlibReturn
    
                                                Case -3, -4:    ' compressed data is corrupted or out of memory
                                                    loopcount = 11
                                                Case -5:        ' too small output buffer, make it larger and try again
                                                    Decomp_Len = 2 * (loopcount + 1) * BytesLeft - 1
                                                    ReDim decomp_bytes(Decomp_Len)
                                                    ZlibReturn = ZlibUncompress(decomp_bytes(0), Decomp_Len, raw_bytes(0), BytesLeft)
                                                Case Else:
                                                    loopcount = 11
                                                    
                                            End Select
    
                                            ' let the error handler loop some time before returning error code
                                            If loopcount >= 10 And ZlibReturn <> 0 Then
                                                ' no success in error handling, function failed, return error code
                                                ' close files
                                                Close #FileNum2
                                                Close #FileNum1
                                                
                                                ' remove buffers from memory
                                                Erase raw_bytes
                                                Erase decomp_bytes
                                                
                                                ' return error code
                                                Decompress = ZlibReturn
                                                Exit Function
                                            Else
                                                loopcount = loopcount + 1
                                            End If
    
                                        Loop
    
                                        ' resize the decompression buffer, if needed
                                        ReDim Preserve decomp_bytes(Decomp_Len - 1)
    
                                        ' if the user wants to check for crc32, then let's do so
                                        If CheckCRC32 Then
                                            crc = ZlibCrc32(crc, decomp_bytes(0), Decomp_Len)
                                            ' DoubleCRC32 might be counted if the user wants to do so
                                            ' The second crc32 is counted from the compressed data
                                            ' and the first from the decompressed data
                                            If DoubleCRC32 Then crc2 = ZlibCrc32(crc2, raw_bytes(0), BytesLeft)
                                        End If
    
                                        ' write decompressed data buffer
                                        Put #FileNum2, , decomp_bytes()
    
                                        ' update progressbar
                                        value = Loc(FileNum1)
                                    
                                    Loop
                                Close #FileNum2
                                
                                ' detect corruption
                                Dim NotCorrupted As Boolean
                                NotCorrupted = True
                                
                                ' SaveCorruptedFiles negates the effect of Checking CRC32
                                If CheckCRC32 And Not SaveCorruptedFiles Then
                                    If crc <> .Original_Crc32 Then NotCorrupted = False
                                    If DoubleCRC32 Then NotCorrupted = NotCorrupted & (crc2 = .CompressedFile_crc32)
                                End If
                                
                                If NotCorrupted Then
                                    ' file was not corrupted or the user wanted to save corrupted data
                                    ' copy the tempfile to its assigned directory
                                    FileCopy .TempName, ChkDir
                                End If
                                
                                ' destroy tempfile
                                Kill .TempName
                            Else
                                ' skip over the filedata, file already existed and user didn't allow overwriting
                                Seek #FileNum1, Loc(FileNum1) + .comp_data + 1
                            End If  ' chkdir <> ""
                            
                        End If  ' GetFileInfoOnly
                    End If
                    
                End With
            Next i
            
        Close #FileNum1
        ' return success code
        Decompress = 0
    Else
        ' file was not found or it was a directory
        Decompress = -9
    End If

End Function

'*************************************************************************************'
'*  Function CheckBoundaries                                                         *'
'*    Checks that the index of a file does not go out of range.                      *'
'*************************************************************************************'
Private Function CheckBoundaries(ByVal index As Long) As Boolean
    CheckBoundaries = (mFiles > -1 And index > -1 And index <= mFiles)
End Function

'*************************************************************************************'
'*  Property Original_Crc32                                                          *'
'*    Returns specified file's CRC32 as String or if the file was not found then     *'
'*    it returns ""                                                                  *'
'*************************************************************************************'

Public Property Get Original_Crc32(ByVal index As Long) As String
    If CheckBoundaries(index) Then
        Dim temp As String
        temp = Hex(Files(index).Original_Crc32)
        Original_Crc32 = String(8 - Len(temp), "0") & temp
    Else
        Original_Crc32 = ""
    End If
End Property

'*************************************************************************************'
'*  Property filename                                                                *'
'*    Get/Change filename                                                            *'
'*************************************************************************************'

Public Property Let filename(ByVal index As Long, ByVal value As String)
    If CheckBoundaries(index) Then
        If InStr(value, ":") = 2 Then value = Right(value, Len(value) - 2)
        If InStr(value, "\") <> 1 Then value = "\" & value
        Files(index).filename = value
    End If
End Property

Public Property Get filename(ByVal index As Long) As String
    If CheckBoundaries(index) Then
        filename = Files(index).filename
    Else
        filename = ""
    End If
End Property

'*************************************************************************************'
'*  Property Path                                                                    *'
'*    Get Path of the file to be compressed.                                         *'
'*************************************************************************************'

Public Property Get Path(ByVal index As Long) As String
    If CheckBoundaries(index) Then
        Path = Files(index).Path
    Else
        Path = ""
    End If
End Property

'*************************************************************************************'
'*  Property Compressed_CRC32                                                        *'
'*    Get compressed file CRC32, most likely returns "00000000" if the file hasn't   *'
'*    yet been compressed.                                                           *'
'*************************************************************************************'

Public Property Get Compressed_CRC32(ByVal index As Long) As String
    If CheckBoundaries(index) Then
        Dim temp As String
        temp = Hex(Files(index).CompressedFile_crc32)
        Compressed_CRC32 = String(8 - Len(temp), "0") & temp
    Else
        Compressed_CRC32 = ""
    End If
End Property

'*************************************************************************************'
'*  Property Compressed_Size                                                         *'
'*    Get compressed filesize, returns 0 if the file hasn't been compressed and -1   *'
'*    in case of an out of boundaries -error.                                        *'
'*************************************************************************************'

Public Property Get Compressed_Size(ByVal index As Long) As Long
    If CheckBoundaries(index) Then
        Compressed_Size = Files(index).comp_data
    Else
        Compressed_Size = -1
    End If
End Property

'*************************************************************************************'
'*  Property FileSize                                                                *'
'*    Get filesize, returns -1 in case of an out of boundaries -error.               *'
'*************************************************************************************'

Public Property Get FileSize(ByVal index As Long) As Long
    If CheckBoundaries(index) Then
        FileSize = Files(index).raw_data
    Else
        FileSize = -1
    End If
End Property

'*************************************************************************************'
'*  Property ReadBuf                                                                 *'
'*    Get/Change ReadBuf                                                             *'
'*************************************************************************************'

Public Property Get ReadBuf() As Long
    ReadBuf = mReadBuf
End Property

Public Property Let ReadBuf(ByVal value As Long)
    If value > 0 Then mReadBuf = value Else mReadBuf = 64 * 1024&
End Property

'*************************************************************************************'
'*  Property ReadBuf                                                                 *'
'*    Get the number of files to be compressed.                                      *'
'*************************************************************************************'

Public Property Get FileCount() As Long
    FileCount = mFiles + 1
End Property

'*************************************************************************************'
'*  Function NewTempFileName                                                         *'
'*    Determine a unique temporary filename using W32API.                            *'
'*************************************************************************************'

Private Function NewTempFileName() As String
    Dim temp As String * MAX_PATH
    If GetTempFileName(mTempPath, TempFilePrefix, 0, temp) = 0 Then Err.Raise Err.LastDllError
    NewTempFileName = Left(temp, InStr(temp, Chr(0)) - 1)
End Function

'*************************************************************************************'
'*  Sub FileAdd                                                                      *'
'*    Increases filecount and saves some data on the new file.                       *'
'*************************************************************************************'

Private Sub FileAdd(ByVal filename As String)
    ' increase filecount
    mFiles = mFiles + 1
    ReDim Preserve Files(mFiles)
    
    ' save some filespecific data
    Files(mFiles).Path = filename
    Files(mFiles).filename = Right(filename, Len(filename) - 2)
    Files(mFiles).raw_data = FileLen(filename)
    Files(mFiles).TempName = NewTempFileName
    
    ' reset progressbar
    Max = Files(mFiles).raw_data
    value = 0
    
    ' calculate file's crc32
    Files(mFiles).Original_Crc32 = Calculate_Crc32(filename)
    
    ' create a temp file to prevent multiple files with same tempfile
    Dim FileNum As Integer
    FileNum = FreeFile
    Open Files(mFiles).TempName For Binary As #FileNum
    Close #FileNum
End Sub

'*************************************************************************************'
'*  Function RemoveFile                                                              *'
'*    Removes a file from the buffer.                                                *'
'*                                                                                   *'
'*    Arguments:                                                                     *'
'*      index - index of the file to be removed.                                     *'
'*                                                                                   *'
'*    Return Values:                                                                 *'
'*      boolean - false if the index was out of boundaries, else true.               *'
'*************************************************************************************'

Public Function RemoveFile(ByVal index As Long) As Boolean
    If CheckBoundaries(index) Then
        If mFiles = 0 Then
            ' if only one file exists, then remove the array from memory
            Erase Files
            mFiles = -1
        Else
            Dim i As Long
            ' decrease filecount
            mFiles = mFiles - 1
            ' fix up "holes" in the array (remove the wanted item)
            For i = index To mFiles
                Files(i) = Files(i + 1)
            Next i
            ' resize array
            ReDim Preserve Files(mFiles)
        End If
        RemoveFile = True
    Else
        ' out of boundaries
        RemoveFile = False
    End If
End Function

'*************************************************************************************'
'*  Function RemoveFile                                                              *'
'*    Removes a file from the buffer.                                                *'
'*                                                                                   *'
'*    Arguments:                                                                     *'
'*      filename - full path and filename to be added. (wildcards (*, #, ?) allowed.)*'
'*      attrib - File attribute. If filename includes wildcards, then only this kind *'
'*               of files will be added.                                             *'
'*                                                                                   *'
'*    Return Values:                                                                 *'
'*      boolean - false if the file didn't exists was out of boundaries, else true.  *'
'*************************************************************************************'

Public Function AddFile(ByVal filename As String, Optional ByVal attrib As VbFileAttribute = vbNormal) As Boolean
    If Dir(filename, attrib) <> "" Then
        If InStr(filename, "*") Or InStr(filename, "?") Or InStr(filename, "#") Then
            ' wildcards specified
            Dim temp As String, Path As String, i As Long, file() As String, f1 As Long
            ' determine path
            Path = filename
            i = Len(Path)
            Do While InStr(Mid(Path, i), "\") <> 1
                i = i - 1
            Loop
            Path = Mid(Path, 1, i)
            temp = Dir(filename, attrib)
            ' find files that match the wildcards and attribs
            f1 = -1
            Do While temp <> ""
                If (GetAttr(Path & temp) And vbDirectory) <> vbDirectory Then
                    f1 = f1 + 1
                    ReDim Preserve file(f1)
                    file(f1) = Path & temp
                End If
                temp = Dir
            Loop
            For i = 0 To f1
                FileAdd file(i)
            Next i
        ElseIf attrib <> vbDirectory And (GetAttr(filename) And vbDirectory) <> vbDirectory Then
            ' add file
            FileAdd filename
        End If
        AddFile = True
    Else
        ' file did not exists
        AddFile = False
    End If
End Function

'*************************************************************************************'
'*  Sub Initialize                                                                   *'
'*    Control initialization.                                                        *'
'*************************************************************************************'

Private Sub UserControl_Initialize()
    Dim i As Long
    mTempPath = Space(MAX_PATH)             ' Search for Windows Temporary files path
    i = GetTempPath(MAX_PATH, mTempPath)    '   |
    If i > MAX_PATH Then                    '   |
        mTempPath = Space(i)                '   |
        i = GetTempPath(i, mTempPath)       '   |
    End If                                  '   |
    mTempPath = Left(mTempPath, i)          '   V  Done
    mFiles = -1
End Sub

'*************************************************************************************'
'*  Property BackColor                                                               *'
'*    Get/Change BackColor                                                           *'
'*************************************************************************************'

Public Property Let BackColor(ByVal NewColor As OLE_COLOR)
    mBackColor = NewColor
    UserControl.BackColor = NewColor
    UserControl_Paint
    PropertyChanged "BackColor"
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackColor
End Property

'*************************************************************************************'
'*  Property BarColor                                                                *'
'*    Get/Change BarColor                                                            *'
'*************************************************************************************'

Public Property Let BarColor(ByVal NewColor As OLE_COLOR)
    mBarColor = NewColor
    UserControl_Paint
    PropertyChanged "BarColor"
End Property
Public Property Get BarColor() As OLE_COLOR
    BarColor = mBarColor
End Property

'*************************************************************************************'
'*  Property Vertical                                                                *'
'*    Get/Change Vertical                                                            *'
'*************************************************************************************'

Public Property Let Vertical(ByVal val As Boolean)
    mVertical = val
    UserControl_Resize
    PropertyChanged "Vertical"
End Property
Public Property Get Vertical() As Boolean
    Vertical = mVertical
End Property

'*************************************************************************************'
'*  Property Max                                                                     *'
'*    Get/Change Max value of the progressbar.                                       *'
'*************************************************************************************'

Private Property Let Max(ByVal val As Long)
    ' don't let it fall out of limits, and keep other variables within new limits
    If val < 1 Then val = 1
    If val <= mMin Then val = mMin + 1
    mMax = val
    If value > mMax Then value = mMax
    UserControl_Resize
    PropertyChanged "Max"
End Property
Private Property Get Max() As Long
    Max = mMax
End Property

'*************************************************************************************'
'*  Property Max                                                                     *'
'*    Get/Change Max value of the progressbar.                                       *'
'*************************************************************************************'

Private Property Let value(ByVal val As Long)
    ' keep the value within the limits
    If val > mMax Then val = Max
    If val < mMin Then val = mMin
    mValue = val
    ' update progressbar
    UserControl_Paint
    PropertyChanged "Value"
End Property
Private Property Get value() As Long
    value = mValue
End Property

'*************************************************************************************'
'*  Property Flat                                                                    *'
'*    Get/Change Flatness value of the progressbar.                                  *'
'*************************************************************************************'

Public Property Let Flat(val As Boolean)
    If val = True Then
        UserControl.BorderStyle = vbBSNone
    Else
        UserControl.BorderStyle = vbFixedSingle
    End If
    UserControl_Resize
    PropertyChanged "Flat"
End Property
Public Property Get Flat() As Boolean
    Flat = (UserControl.BorderStyle = vbBSNone)
End Property

'*************************************************************************************'
'*  Sub InitProperties                                                               *'
'*    Initialize needed values.                                                      *'
'*************************************************************************************'

Private Sub UserControl_InitProperties()
    BackColor = vbButtonFace
    BarColor = vbHighlight
    Vertical = False
    Max = 100
    mMin = 0
    value = 0
    Flat = False
    mReadBuf = 65536
End Sub

'*************************************************************************************'
'*  Sub ReadProperties                                                               *'
'*    Read needed values.                                                            *'
'*************************************************************************************'

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
    BarColor = PropBag.ReadProperty("BarColor", vbHighlight)
    Vertical = PropBag.ReadProperty("Vertical", False)
    Max = PropBag.ReadProperty("Max", 100)
    value = PropBag.ReadProperty("Value", 0)
    Flat = PropBag.ReadProperty("Flat", False)
    ReadBuf = PropBag.ReadProperty("ReadBuf", 65536)
End Sub

'*************************************************************************************'
'*  Sub Terminate                                                                    *'
'*    Removes files array from memory.                                               *'
'*************************************************************************************'

Private Sub UserControl_Terminate()
    Erase Files
End Sub

'*************************************************************************************'
'*  Sub WriteProperties                                                              *'
'*    Write needed values.                                                           *'
'*************************************************************************************'

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "ReadBuf", mReadBuf, 65536
    PropBag.WriteProperty "BackColor", BackColor, vbButtonFace
    PropBag.WriteProperty "BarColor", BarColor, vbHighlight
    PropBag.WriteProperty "Vertical", Vertical, False
    PropBag.WriteProperty "Max", Max, 100
    PropBag.WriteProperty "Value", value, 0
    PropBag.WriteProperty "Flat", Flat, False
End Sub

'*************************************************************************************'
'*  Sub Resize                                                                       *'
'*    Handles resizing of the control                                                *'
'*************************************************************************************'

Private Sub UserControl_Resize()
    On Error Resume Next        ' just in case
    UserControl.ScaleWidth = mMax - mMin
    UserControl.ScaleHeight = mMax - mMin
    UserControl_Paint               ' repaint the control
End Sub

'*************************************************************************************'
'*  Sub Click                                                                        *'
'*    To give the control of the Click event to the coder.                           *'
'*************************************************************************************'

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

'*************************************************************************************'
'*  Sub Reset                                                                        *'
'*    Reset values, remove files.                                                    *'
'*************************************************************************************'

Public Sub reset()
    Erase Files
    mFiles = -1
    value = 0
End Sub

'*************************************************************************************'
'*  Sub Paint                                                                        *'
'*    Repaints the progressbar.                                                      *'
'*************************************************************************************'

Private Sub UserControl_Paint()
    Dim w As Long           ' I'm storing some properties
    Dim h As Long           ' in variables to improve performance
    Dim v As Long
    v = mValue - mMin
    w = UserControl.ScaleWidth
    h = UserControl.ScaleHeight
    If mVertical Then                                                     ' is this a vertical control?
        UserControl.Line (0, 0)-(w, h - v), mBackColor, BF  ' draw the background color
        If v > 0 Then   ' only draw the bar if there is one to draw
            UserControl.Line (0, h)-(w, h - v), mBarColor, BF   ' draw the bar
        End If
    Else
        UserControl.Line (v, 0)-(w, h), mBackColor, BF          ' this is the same code as above
        If v > 0 Then
            UserControl.Line (0, 0)-(v, h), mBarColor, BF         ' but for horizontal controls
        End If
    End If
End Sub
