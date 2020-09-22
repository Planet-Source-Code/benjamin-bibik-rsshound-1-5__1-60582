Attribute VB_Name = "modFileSys"
Option Explicit

' Important - any project that makes use of this module must
'             also include the Common.bas module.

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' The Archive bit is set to True whenever the File Save dialog saves
' a file. It is convention to do this if you use alternative methods
' to modify files.

' A backup procedure could reset the Archive bits for all files it
' backs up, and so the next backup need only update those files with
' their Archive bit set to True.

Private Declare Sub SHAddToRecentDocsA Lib "Shell32" Alias "SHAddToRecentDocs" (ByVal wFlags As Long, ByVal sFileSpec As String)
Private Declare Function SHFileOperationA Lib "Shell32" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function SHGetSpecialFolderPathA Lib "Shell32" (ByVal hWndOwner As Long, ByVal lpszPath As String, ByVal lFolder As SpecialFolders, ByVal fCreate As Long) As Long

' For many system folders, a qualified path can be obtained by calling
' SHGetSpecialFolderLocation or SHGetSpecialFolderPath with an appropriate
' CSIDL constant specifying the folder for which to retrieve the location.
' This parameter can be one of the following values:
Public Enum SpecialFolders
    'sfDESKTOP = &H0&                  ' Windows desktop — virtual folder at the root of the name space.
    sfPROGRAMS = &H2&                 ' Start Menu\Programs. File system directory that contains the user's program groups which are also file system directories.
    'sfCONTROLS = &H3&                 ' Control Panel — virtual folder containing icons for the control panel applications.
    'sfPRINTERS = &H4&                 ' Printers folder — virtual folder containing installed printers.
    sfPERSONAL = &H5&                 ' My Documents. File system directory that serves as a common repository for documents.
    sfFAVORITES = &H6&                ' Favorites. File system directory that serves as a common repository for favourite web documents.
    sfSTARTUP = &H7&                  ' Start Menu\Programs\Startup. File system directory that corresponds to the user's Startup program group.
    sfRECENT = &H8&                   ' Recent documents. File system directory containing the user's most recently used documents.
    sfSENDTO = &H9&                   ' SendTo folder. File system directory containing Send To menu items.
    sfBITBUCKET = &HA                 ' Recycle bin — file system directory containing file objects in the user's recycle bin. The location of this directory is not in the registry; it is marked with the hidden and system attributes to prevent the user from moving or deleting it.
    sfSTARTMENU = &HB&                ' Start Menu. File system directory containing Start menu items.
    sfDESKTOPDIR = &H10&              ' Windows\Desktop. File system directory used to physically store file objects on the desktop (not to be confused with the desktop virtual folder itself).
    'sfDRIVES = &H11                   ' My Computer — virtual folder containing everything on the local computer: storage devices, printers, and Control Panel. The folder can also contain mapped network drives.
    'sfNETWORK = &H12                  ' Network Neighborhood — virtual folder representing the top level of the network hierarchy.
    sfNETHOOD = &H13&                 ' NetHood. File system directory containing objects that appear in the network neighborhood.
    'sfFONTS = &H14&                   ' Fonts - virtual folder containing fonts.
    sfTEMPLATES = &H15&               ' ShellNew templates folder. File system directory that serves as a common repository for document templates.
    sfCOMMON_STARTMENU = &H16&        ' Start Menu (All Users). File system directory that contains the programs and folders that appear on the Start menu for all users.
    sfCOMMON_PROGRAMS = &H17&         ' All Users\Start Menu\Programs. File system directory that contains the directories for the common program groups that appear on the Start menu for all users.
    sfCOMMON_STARTUP = &H18&          ' All Users\Start Menu\Programs\Startup. File system directory that contains the programs that appear in the Startup folder for all users. The system starts these programs whenever any user logs on to a Windows desktop platform.
    sfCOMMON_DESKTOPDIR = &H19&       ' All Users\Desktop. File system directory that contains files and folders that appear on the desktop for all users.
    sfCOMMON_FAVORITES = &H1F&        ' All Users\Favorites
    sfINTERNETCACHE = &H20&           ' Internet Cache folder
    sfCOOKIES = &H21&                 ' Cookies folder
    sfHISTORY = &H22&                 ' History folder
    sfAPPDATA = &H1A&                 ' Application Data
    sfPRINTHOOD = &H1B&               ' PrintHood
End Enum

'+++++++++ Integers! +++++++++
Private Const FO_MOVE = &H1&    ' Moves the files specified in pFrom to the location
                                ' specified in pTo.
Private Const FO_COPY = &H2&    ' Copies the files specified in the pFrom member
                                ' to the location specified in the pTo member.
Private Const FO_DELETE = &H3&  ' Deletes the files specified in pFrom (pTo is ignored).
Private Const FO_RENAME = &H4&  ' Renames the files specified in pFrom.
'+++++++++++++++++++++++++++++

'++++++++++++++++++++++++++++++++++++++++++
Private Const FOF_ALLOWUNDO = &H40         ' Preserve Undo information, if possible. Put
                                           ' deleted files (except those from floppy
                                           ' disks) in Recycle Bin. If pFrom does not
                                           ' contain fully qualified path and filenames,
                                           ' this flag is ignored.
Private Const FOF_FILESONLY = &H80         ' Perform the operation on files only if a
                                           ' wildcard file name (*.*) is specified.
Private Const FOF_MULTIDESTFILES = &H1     ' The pTo member specifies multiple dest'n
                                           ' files (one for each source file) rather
                                           ' than one directory where all source
                                           ' files are to be deposited.
Private Const FOF_NOCONFIRMATION = &H10    ' Respond with 'Yes to All' for any dialog
                                           ' box that is displayed.
Private Const FOF_NOCONFIRMMKDIR = &H200   ' Does not confirm the creation of a new
                                           ' directory if the operation requires one
                                           ' to be created.
Private Const FOF_RENAMEONCOLLISION = &H8  ' Give the file being operated on a new
                                           ' name in a move, copy, or rename op'n if a
                                           ' file with the target name already exists.
Private Const FOF_SILENT = &H4             ' Does not display a progress dialog box.
Private Const FOF_SIMPLEPROGRESS = &H100   ' Displays a progress dialog box but does
                                           ' not show the file names.
'++++++++++++++++++++++++++++++++++++++++++

'+++++++++++++++++++++++
Public Enum FileOpFlags
    AutoMkDir = &H200
    FilesOnly = &H80
    NoDlgBox = &H4
    RenIfExist = &H8
    YesToAll = &H10
End Enum
'+++++++++++++++++++++++

'+++++++++++++++++++++++++++++++++++++
Private Type SHFILEOPSTRUCT
    hWnd As Long                      ' Window owner of any dialogs
    wFunc As Long                     ' Copy, move, rename, or delete code
    pFrom As String                   ' Source file
    pTo As String                     ' Destination file or directory
    fFlags As Integer                 ' Options to control the operations
    fAnyOperationsAborted As Boolean  ' Indicates partial failure
    hNameMappings As Long             ' Array indicating each success
    lpszProgressTitle As String       ' Title for progress dialog, only
End Type                              ' used if FOF_SIMPLEPROGRESS
'+++++++++++++++++++++++++++++++++++++

' The LongPathName function converts the specified path to its long form.
' If no long path is found, this function returns the specified name.
Private Declare Function LongPathName Lib "kernel32" Alias _
    "GetLongPathNameA" (ByVal lpszShortPath As String, _
    ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long

' The ShortPathName function converts the specified path to its short form.
Private Declare Function ShortPathName Lib "kernel32" Alias _
    "GetShortPathNameA" (ByVal lpszLongPath As String, _
    ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As Currency, lpLastAccessTime As Currency, lpLastWriteTime As Currency) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As Currency, lpLastAccessTime As Currency, lpLastWriteTime As Currency) As Long

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_READ = &H80000000
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const OPEN_EXISTING = &H3
Private Const OF_READ = &H0

Private Declare Function GetFileInformationByHandle Lib "kernel32" _
    (ByVal hFile As Long, lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Long

Public Type BY_HANDLE_FILE_INFORMATION
    dwFileAttributes As Long
    ftCreationTime As Currency
    ftLastAccessTime As Currency
    ftLastWriteTime As Currency
    dwVolumeSerialNumber As Long
    nFileSizeHigh As Long
    nFileSizeLow As Long
    nNumberOfLinks As Long
    nFileIndexHigh As Long
    nFileIndexLow As Long
End Type

Private Const OFS_MAXPATHNAME = 128
Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Public Declare Function GetDriveType32 Lib "kernel32" Alias _
    "GetDriveTypeA" (ByVal sDrive As String) As Long

Public Declare Function GetDrivesString Lib "kernel32" Alias _
    "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, _
     ByVal lpBuffer As String) As Long

' GetDriveType32 return values
Public Const DRIVE_TYPE_UNKNOWN = 0   ' Unknown media type
Public Const DRIVE_ROOT_NOT_EXIST = 1 ' No such root directory exists
Public Const DRIVE_REMOVABLE = 2      ' Drive can be removed
Public Const DRIVE_FIXED = 3          ' Drive cannot be removed
Public Const DRIVE_REMOTE = 4         ' Network disk drive
Public Const DRIVE_CDROM = 5          ' CD-ROM disk drive
Public Const DRIVE_RAMDISK = 6        ' RAM disk drive

' GetDriveType32 return error value
Public Const DRIVE_INVALID = DRIVE_TYPE_UNKNOWN Or DRIVE_ROOT_NOT_EXIST

Public Enum eDriveTypes
    edtRemovables = 2       ' Drive can be removed
    edtHardDrives = 3       ' Drive cannot be removed
    edtRemoteDrives = 4     ' Network disk drive
    edtCDROMDrives = 5      ' CD-ROM disk drive
    edtRamDrives = 6        ' RAM disk drive
    edtAllDrives = 7        ' All Drive Types
End Enum

Public Declare Function GetAttributes Lib "kernel32" Alias "GetFileAttributesA" _
    (ByVal lpSpec As String) As Long

Public Declare Function SetAttributes Lib "kernel32" Alias "SetFileAttributesA" _
    (ByVal lpSpec As String, ByVal dwAttributes As Long) As Long

Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" _
             (ByVal lpszPath As String, ByVal lpPrefixString As String, _
              ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" _
             (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Declare Function GetWinDir Lib "kernel32" Alias "GetWindowsDirectoryA" _
             (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Function GetWinSysDir Lib "kernel32" Alias "GetSystemDirectoryA" _
             (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Function SearchNullPath Lib "kernel32" Alias "SearchPathA" _
    (ByVal nNullPath As Long, ByVal lpFileName As String, ByVal nNullExt As Long, _
     ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long

' The SearchTreeForFile function is used to search a directory tree for a specified file.
Private Declare Function SearchTreeForFile Lib "ImageHlp" (ByVal lpRootPath As String, ByVal lpFileName As String, ByVal lpBuffer As String) As Long

' The GetPrivateProfileString function is not case-sensitive.
Private Declare Function GetINIString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpSectName As String, ByVal lpKeyName As String, ByVal lpDefault As String, _
     ByVal lpRetString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function WriteINIString Lib "kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal lpSectName As String, ByVal lpKeyName As String, ByVal lpString As String, _
     ByVal lpFileName As String) As Long


Public Enum eFilePart
    efpFullFileSpec
    efpFullPath
    efpFileNameExt
    efpFileName
    efpFileExt
End Enum

Private Const INVALID_ARG As Long = FIVE
Private Const FILE_NOT_FOUND As Long = 53
Private Const PATH_NOT_FOUND As Long = 76
Private Const EMPTY_STR As String = "Empty string passed."

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Private Declare Function GetDiskFreeSpaceExA Lib "kernel32" _
        (ByVal lpDrive As String, _
         lpFreeBytesAvailableToCallerDividedBy10000 As Currency, _
         lpTotalNumberOfBytesDividedBy10000 As Currency, _
         lpTotalNumberOfFreeBytesDividedBy10000 As Currency) As Long
' Note - the values obtained by GetDiskFreeSpaceExA are ULARGE_INTEGER.
' Therefore, they are declared as Currency, and so need to be multiplied
' by 10000 to get the correct number of Bytes (moves the decimal place).
' Use the GetDiskFreeSpace32 wrapper function below.
  Private Const ERROR_CALL_NOT_IMPLEMENTED = 120&
' If GetDiskFreeSpaceExA fails with the ERROR_CALL_NOT_IMPLEMENTED code,
' use the GetDiskFreeSpace function instead (pre-Windows 95b/OSR2).
' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' If this function succeeds, the return value is nonzero.
Public Function GetDiskFreeSpace32(sDrive As String, FreeBytesAvailable As Currency, _
        TotalNumberOfBytes As Currency, TotalNumberOfFreeBytes As Currency) As Long

    Dim cAvailable As Currency, cTotal As Currency, cTotalFree As Currency
    Dim rc As Long, sDrv As String

    sDrv = IIf(Len(sDrive) > THREE, Left$(sDrive, THREE), sDrive)
    rc = GetDiskFreeSpaceExA(sDrv, cAvailable, cTotal, cTotalFree)

    If (rc <> ERROR_CALL_NOT_IMPLEMENTED) Then
        FreeBytesAvailable = cAvailable * 10000
        TotalNumberOfBytes = cTotal * 10000
        TotalNumberOfFreeBytes = cTotalFree * 10000
        GetDiskFreeSpace32 = rc
    End If
End Function


'-----------------------------------------------------------
' Creates the specified directory sPath.
' Returns: 2 if created, 1 if existed, 0 if error.
'-----------------------------------------------------------
Public Function MakePath(sPath As String) As Long
    If LenB(sPath) = ZERO Then Err.Raise INVALID_ARG, , EMPTY_STR
    Dim sDir As String, sTemp As String, Idx As Integer

    On Error GoTo FailedMakePath
    If (DirExists(sPath)) Then
        MakePath = ONE
        Exit Function
    End If
    sDir = sPath

    ' Add trailing backslash if missing
    If Right$(sDir, ONE) <> DIR_SEP Then sDir = sDir & DIR_SEP

    ' Set Idx to the first backslash
    Idx = InStr(ONE, sDir, DIR_SEP)

    Do ' Loop and make each subdir of the path separately
        Idx = InStr(Idx + ONE, sDir, DIR_SEP)
        If (Idx) Then
            sTemp = Left$(sDir, Idx - ONE)
            ' Determine if this directory already exists
            If (Not DirExists(sTemp)) Then
                ' We must create this directory
                MkDir sTemp
                MakePath = TWO
            End If
        End If
    Loop Until Idx = ZERO
FailedMakePath:
End Function

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++-©Rd-+
' SubstFile gets the creation time and attributes from the old
' file and assigns these to the new file. It then DELETES the
' old file, and renames the new file to old file name.
' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Sub SubstFile(sOldFile As String, sNewFile As String)
    If (LenB(sOldFile) = ZERO) Or (LenB(sNewFile) = ZERO) Then Err.Raise INVALID_ARG, , EMPTY_STR
    If Not (FileExists(sOldFile) And FileExists(sNewFile)) Then Err.Raise FILE_NOT_FOUND

    ' Get file time and attributes from old file
    SetCreateTime sNewFile, GetCreateTime(sOldFile)
    SetAttributes sNewFile, GetAttributes(sOldFile)

    SetAttributes sOldFile, vbNormal
    Kill sOldFile
    Name sNewFile As sOldFile
End Sub

Public Function GetCreateTime(sFileSpec As String) As Currency
    If LenB(sFileSpec) = ZERO Then Err.Raise INVALID_ARG, , EMPTY_STR
    If (Not FileExists(sFileSpec)) Then Err.Raise FILE_NOT_FOUND

    On Error GoTo HandleIt
    ' Gets the creation time for the specified file
    Dim junk1 As Currency, junk2 As Currency
    Dim hFile As Long, dtCreationTime As Currency

    hFile = CreateFile(sFileSpec, GENERIC_READ, ZERO, ZERO, _
                OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, ZERO)

    GetFileTime hFile, dtCreationTime, junk1, junk2
    CloseHandle hFile
    GetCreateTime = dtCreationTime
HandleIt:
End Function

Public Sub SetCreateTime(sFileSpec As String, ByVal dtFileCreated As Currency)
    If LenB(sFileSpec) = ZERO Then Err.Raise INVALID_ARG, , EMPTY_STR
    If (Not FileExists(sFileSpec)) Then Err.Raise FILE_NOT_FOUND

    On Error GoTo HandleIt
    ' Updates the date/time for the specified file
    Dim hFile As Long, Junk As Currency
    Dim dtAccTime As Currency, dtModTime As Currency

    hFile = CreateFile(sFileSpec, GENERIC_WRITE Or GENERIC_READ, ZERO, _
                     ZERO, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, ZERO)

    GetFileTime hFile, Junk, dtAccTime, dtModTime
    SetFileTime hFile, dtFileCreated, dtAccTime, dtModTime
    DoEvents
    CloseHandle hFile
HandleIt:
End Sub

Public Function SearchForFile(sFileName As String, Optional sRootPath As String, _
              Optional ByVal FilePart As eFilePart = efpFullFileSpec) As String
    If LenB(sFileName) = ZERO Then Err.Raise INVALID_ARG, , EMPTY_STR
    Dim lLen As Long, sFileSpec As String
    sFileSpec = String$(MAX_PATH, vbNullChar)
    If LenB(sRootPath) <> ZERO Then
        lLen = SearchTreeForFile(sRootPath, sFileName, sFileSpec)
        ' sRootPath specifies the root path to be searched for the file.
    Else
        lLen = SearchNullPath(ZERO&, sFileName, ZERO&, MAX_PATH, sFileSpec, ZERO&)
        ' Search for sFileName in the following dirs in this sequence:
        '  1. The directory from which the application loaded.
        '  2. The current directory.
        '  3. The Windows System directory (Win95), or the 32-bit Windows
        '     System32 directory (WinNT). Use the GetSystemDirectory function
        '     to get the path of this directory, or modFileSys.WinSysDir.
        '  4. The 16-bit Windows System directory (WinNT).
        '  5. The Windows directory. Use the GetWindowsDirectory function to
        '     get the path of this directory, or modFileSys.WinDir.
        '  6. The directories that are listed in the PATH environment variable.
    End If
    If (lLen) Then
        sFileSpec = TrimZ(sFileSpec)
        If (FilePart = efpFullFileSpec) Then
            SearchForFile = sFileSpec
        Else
            SearchForFile = GetFilePart(sFileSpec, FilePart)
        End If
    End If
End Function

Public Function GetFilePart(sFileSpec As String, ByVal FilePart As eFilePart) As String
    If LenB(sFileSpec) = ZERO Then Err.Raise INVALID_ARG, , EMPTY_STR
    Dim lBackSlash As Long, lPeriod As Long
    lBackSlash = InStrR(sFileSpec, DIR_SEP)
    lPeriod = InStrR(sFileSpec, ".")
    ' Extract the path, file, and/or extension
    Select Case FilePart
        Case efpFullFileSpec: GetFilePart = sFileSpec
        Case efpFullPath:     GetFilePart = Left$(sFileSpec, lBackSlash - ONE)
        Case efpFileNameExt:  GetFilePart = Mid$(sFileSpec, lBackSlash + ONE)
        Case efpFileName:     GetFilePart = Mid$(sFileSpec, lBackSlash + ONE, lPeriod - ONE - lBackSlash)
        Case efpFileExt:      If (lPeriod) Then GetFilePart = Mid$(sFileSpec, lPeriod + ONE)
    End Select
End Function

'---GetTextFileByLine------------------------------------------
' This function opens the specified file for input and assigns
' the first line in the file to sNextLine, returning True.
' On subsequent calls (omitting sInitFileSpec) it assigns
' the next line in the file to sNextLine, returning True.
' It returns False when there are no more lines in the file,
' and sNextLine is returned unaffected.

' *All opened files must be closed by calling this function*
' until the end of the file, or else a file can be closed by:
' An open file can be prematurely closed by specifying a single
' backslash (\) for sInitFileSpec, or simply by opening another
' file by assigning another valid file spec to sInitFileSpec.

' You may wish to only open the file on the first init call,
' by setting the third argument to True. In this case sNextLine
' is returned unaffected, but the function still returns True
' to indicate success. You can then retrieve the first line in
' in the file in the next call (while omitting sInitFileSpec).

' Note: an opened file can be read from but not written to.
'--------------------------------------------------------------
Public Function GetTextFileByLine(ByRef sNextLine As String, Optional sInitFileSpec As String, Optional fFirstLineNextCallAfterInit As Boolean) As Boolean
    Static iFile As Integer
    If LenB(sInitFileSpec) <> ZERO Then
        If (sInitFileSpec = DIR_SEP) Then
            If (iFile <> ZERO) Then
                Close #iFile
                iFile = ZERO
            End If
            Exit Function
        End If
        If (Not FileExists(sInitFileSpec)) Then Err.Raise FILE_NOT_FOUND
    End If

    ' Handle errors if they occur
    On Error GoTo GetFileError

    If LenB(sInitFileSpec) <> ZERO Then
        If (iFile <> ZERO) Then Close #iFile
        iFile = FreeFile()
        ' Let others read but not write
        Open sInitFileSpec For Input Access Read Lock Write As #iFile
        If fFirstLineNextCallAfterInit Then
            GetTextFileByLine = True
            Exit Function
        End If
    End If

    If (iFile = ZERO) Then Exit Function

    ' Get the file in line by line
    If (Not EOF(iFile)) Then
        Line Input #iFile, sNextLine

        ' Return this line in the file
        GetTextFileByLine = True
        Exit Function
    End If

GetFileError:
    If (iFile <> ZERO) Then
        Close #iFile
        iFile = ZERO
    End If

End Function

Public Function GetBinaryFile(sFileSpec As String, aFile() As Byte) As Boolean
    ' Returns True on success, or False otherwise
    If LenB(sFileSpec) = ZERO Then Err.Raise INVALID_ARG, , EMPTY_STR
    If (Not FileExists(sFileSpec)) Then Err.Raise FILE_NOT_FOUND

    ' Handle errors if they occur
    On Error GoTo GetFileError

    ' Change the mouse pointer to an hourglass
    HourGlass True

    Dim iFile As Integer
    iFile = FreeFile

    ' Let others read but not write
    Open sFileSpec For Binary Access Read Lock Write As #iFile
    ReDim aFile(ZERO To LOF(iFile) - ONE)
    Get #iFile, ONE, aFile()
    
    GetBinaryFile = True
GetFileError:
    Close #iFile

    ' Reset to the previous mouse pointer
    HourGlass False
End Function

Public Function SaveBinaryFile(sFileSpec As String, aFile() As Byte) As Long
    If LenB(sFileSpec) = ZERO Then Err.Raise INVALID_ARG, , EMPTY_STR

    ' Handle errors if they occur
    On Error GoTo SaveFileError

    ' Change the mouse pointer to an hourglass
    HourGlass True

    Dim iFile As Integer
    iFile = FreeFile

    Open sFileSpec For Binary Access Write Lock Write As #iFile
    Put #iFile, ONE, aFile()
SaveFileError:
    Close #iFile

    ' Reset to the previous mouse pointer
    HourGlass False

    SaveBinaryFile = Err
End Function

Public Function GetTextFile(sFileSpec As String) As String
    If LenB(sFileSpec) = ZERO Then Err.Raise INVALID_ARG, , EMPTY_STR
    If (Not FileExists(sFileSpec)) Then Err.Raise FILE_NOT_FOUND

    ' Handle errors if they occur
    On Error GoTo GetFileError

    ' Change the mouse pointer to an hourglass
    HourGlass True

    Dim iFile As Integer
    iFile = FreeFile

    ' Open in binary mode, let others read but not write
    Open sFileSpec For Binary Access Read Lock Write As #iFile
    ' Allocate the length first
    GetTextFile = Space$(LOF(iFile))
    ' Get the file in one chunk
    Get #iFile, , GetTextFile
GetFileError:
    Close #iFile ' Close the file

    ' Reset to the previous mouse pointer
    HourGlass False

End Function

Public Function SaveTextFile(sFileSpec As String, sText As String) As Long
    If LenB(sFileSpec) = ZERO Then Err.Raise INVALID_ARG, , EMPTY_STR

    ' Handle errors if they occur
    On Error GoTo SaveFileError

    ' Change the mouse pointer to an hourglass
    HourGlass True

    Dim iFile As Integer
    iFile = FreeFile

    Open sFileSpec For Output Access Write Lock Write As #iFile
    Print #iFile, sText;
SaveFileError:
    Close #iFile

    ' Reset to the previous mouse pointer
    HourGlass False

    SaveTextFile = Err
End Function

Public Function WriteAccess(sDirName As String) As Boolean
    If LenB(sDirName) = ZERO Then Err.Raise INVALID_ARG, , EMPTY_STR
    If (Not DirExists(sDirName)) Then Err.Raise PATH_NOT_FOUND
    Dim FileNum As Integer, sDir As String

    On Error Resume Next
    ' Add trailing backslash if missing
    sDir = AddBackslash(sDirName)

    FileNum = FreeFile
    Open sDir & "IsTemp.tmp" For Output As FileNum

    WriteAccess = (Err = ZERO)

    Close FileNum

    If WriteAccess Then Kill sDir & "IsTemp.tmp"
End Function

'----------------------------------------------------------------
' Adds a document to the shell's list of recently used documents
'----------------------------------------------------------------
Public Sub AddToRecentDocs(sFileSpec As String)
    If LenB(sFileSpec) = ZERO Then Err.Raise INVALID_ARG, , EMPTY_STR
    If (Not FileExists(sFileSpec)) Then Err.Raise FILE_NOT_FOUND
    SHAddToRecentDocsA TWO&, sFileSpec
End Sub

'------------------------------------------------------------
' Clears all documents from the recently used documents list
'------------------------------------------------------------
Public Sub ClearRecentDocs()
    SHAddToRecentDocsA ZERO&, vbNullString
End Sub

'-----------------------------------------------------------
' Retrieves the path of a special folder
'-----------------------------------------------------------
Public Function GetSpecialFolder(ByVal id As SpecialFolders, Optional ByVal Me_hWnd As Long, Optional ByVal fCreate As Boolean) As String
    Dim sBuffer As String
    sBuffer = String$(MAX_PATH, vbNullChar)

    ' SHGetSpecialFolderPathA returns zero on success(one?),
    ' or an OLE-defined error result otherwise
    If (SHGetSpecialFolderPathA(Me_hWnd, sBuffer, id, fCreate) <> ZERO) Then
        GetSpecialFolder = TrimZ(sBuffer)
    End If
End Function

Public Function GetFileInformation(sFileSpec As String, FileInfo As BY_HANDLE_FILE_INFORMATION) As Boolean
    If LenB(sFileSpec) = ZERO Then Err.Raise INVALID_ARG, , EMPTY_STR
    If (Not FileExists(sFileSpec)) Then Err.Raise FILE_NOT_FOUND

    On Error GoTo HandleIt
    ' Get the file information for the specified file
    Dim rc As Long
    Dim hFile As Long
    Dim lpReOpenBuff As OFSTRUCT

    hFile = OpenFile(sFileSpec, lpReOpenBuff, OF_READ)
    GetFileInformation = GetFileInformationByHandle(hFile, FileInfo)
    rc = CloseHandle(hFile)
HandleIt:
End Function

Public Function KillFolder(sSpec As String) As Boolean
    If LenB(sSpec) = ZERO Then Err.Raise INVALID_ARG, , EMPTY_STR
    If (Not DirExists(sSpec)) Then Err.Raise PATH_NOT_FOUND

    On Error GoTo HandleIt
    Dim sDir As String, sFile As String

    ' Add trailing backslash if missing
    sDir = AddBackslash(sSpec)

    sFile = Dir(sDir & "*.*")
    Do While LenB(sFile) <> ZERO
        SetAttributes sDir & sFile, vbNormal
        Kill sDir & sFile
        sFile = Dir
    Loop
    RmDir RemoveBackslash(sDir)
HandleIt:
    If (Not DirExists(sDir)) Then KillFolder = True
End Function

'-------------------------------------------------------------
' This function converts the specified path to its long form.
' If no long path is found, returns the specified short name.
'-------------------------------------------------------------
Public Function GetLongPathName(sShortPath As String) As String
    If LenB(sShortPath) = ZERO Then Err.Raise INVALID_ARG, , EMPTY_STR

    ' Default to the short name
    GetLongPathName = sShortPath

    On Error GoTo GetFailed
    Dim sPath As String
    Dim lResult As Long

    sPath = String$(MAX_PATH, vbNullChar)
    lResult = LongPathName(sShortPath, sPath, MAX_PATH)
    If (lResult) Then GetLongPathName = TrimZ(sPath)
GetFailed:
End Function

'-------------------------------------------------------------
' This function converts the specified path to its short form.
' If no short path is found, returns the specified long name.
'-------------------------------------------------------------
Public Function GetShortPathName(sLongPath As String) As String
    If LenB(sLongPath) = ZERO Then Err.Raise INVALID_ARG, , EMPTY_STR

    ' Default to the long name
    GetShortPathName = sLongPath

    On Error GoTo GetFailed
    Dim sPath As String
    Dim lResult As Long

    sPath = String$(MAX_PATH, vbNullChar)
    lResult = ShortPathName(sLongPath, sPath, MAX_PATH)
    If (lResult) Then GetShortPathName = TrimZ(sPath)
GetFailed:
End Function

Public Function CreateTempFile(Optional sPrefix As String = "tmp", _
                               Optional sPath As String, _
                               Optional lHex As Long) As String
    Dim lResult As Long, sName As String, sTemp As String
    sName = String$(MAX_PATH, vbNullChar)
    sTemp = String$(MAX_PATH, vbNullChar)
    If LenB(sPath) <> ZERO Then
        'If Not DirExists(sPath) Then Err.Raise 76 ' Path not found
        If Not DirExists(sPath) Then
            If Not CBool(MakePath(sPath)) Then sPath = TempDir
        End If
    Else
        sPath = TempDir
    End If
    lResult = GetTempFileName(sPath, sPrefix, lHex, sName)
    If (lResult) Then CreateTempFile = TrimZ(sName)
End Function

Public Function TempDir(Optional EndBackslash As Boolean = True) As String
    Dim sTemp As String, lResult As Long
    sTemp = String$(MAX_PATH, vbNullChar)

    lResult = GetTempPath(MAX_PATH, sTemp)
    If EndBackslash Then
        TempDir = AddBackslash(TrimZ(sTemp))
    Else
        TempDir = RemoveBackslash(TrimZ(sTemp))
    End If
End Function

Public Function WinDir(Optional EndBackslash As Boolean = True) As String
    Dim sTemp As String, lResult As Long
    sTemp = String$(MAX_PATH, vbNullChar)

    lResult = GetWinDir(sTemp, MAX_PATH)
    If EndBackslash Then
        WinDir = AddBackslash(TrimZ(sTemp))
    Else
        WinDir = RemoveBackslash(TrimZ(sTemp))
    End If
End Function

Public Function WinSysDir(Optional EndBackslash As Boolean = True) As String
    Dim sTemp As String, lResult As Long
    sTemp = String$(MAX_PATH, vbNullChar)

    lResult = GetWinSysDir(sTemp, MAX_PATH)
    If EndBackslash Then
        WinSysDir = AddBackslash(TrimZ(sTemp))
    Else
        WinSysDir = RemoveBackslash(TrimZ(sTemp))
    End If
End Function

'fSuccess = SetINIKey("Vbaddin.ini", "Add-Ins32", "MyAddin.Connect", 3)
Public Function SetINIKey(sFile As String, sSection As String, _
                          sKey As String, sValue As String) As Boolean
    SetINIKey = WriteINIString(sSection, sKey, sValue, sFile)
End Function

'sValue = GetINIKey("Vbaddin.ini", "Add-INS32", "MyAddIn.CoNNect")
Public Function GetINIKey(sFile As String, sSection As String, _
                          sKey As String) As String
    Dim lLen As Long, sVal As String, sDefault As String
    Const sDEF As String = "~!~!~"
    sVal = String$(MAX_PATH, vbNullChar)
    sDefault = sDEF
    lLen = GetINIString(sSection, sKey, sDefault, sVal, MAX_PATH, sFile)
    sVal = Left$(sVal, lLen) ' Rd :)
    If (sVal <> sDEF) Then GetINIKey = sVal
End Function

'-----------------------------------------------------------
' This function returns a one based array of drive letters
' assigned to a dynamic string array passed ByRef.
'-----------------------------------------------------------
Public Function GetDrives(DrivesArray() As String, ByVal DriveType As eDriveTypes) As Boolean
    On Error GoTo HandleIt
    Dim aDrives() As String, sDrives As String, sDrive As String
    Dim lType As Long, lStrLen As Long, Idx As Long, lStart As Long
    sDrives = Space$(150)
    ' Get the logial drives on the system
    lStrLen = GetDrivesString(150, sDrives)
    If (lStrLen) Then
        ReDim aDrives(ONE To (lStrLen \ FOUR)) ' Allow for all
        For lStart = ONE To lStrLen Step FOUR
            sDrive = UCase(Mid(sDrives, lStart, THREE))
            lType = GetDriveType32(sDrive)
            If (lType <> DRIVE_INVALID) Then
                If (DriveType = edtAllDrives) Then
                    Idx = Idx + ONE
                    aDrives(Idx) = sDrive
                ElseIf (DriveType = lType) Then
                    Idx = Idx + ONE
                    aDrives(Idx) = sDrive
                End If
            End If
        Next
        GetDrives = (Idx)
        If GetDrives Then
            ReDim Preserve aDrives(ONE To Idx)
            DrivesArray = aDrives ' Return the array
        End If
    End If
HandleIt:
End Function

'-----------------------------------------------------------
' If this function succeeds, the return value is zero.
' If it fails, the return value is the Err.Number.
'-----------------------------------------------------------
Public Function FixLineEnds(sFileSpec As String, Optional ProgressBar As Object) As Long
    If LenB(sFileSpec) = ZERO Then Err.Raise INVALID_ARG, , EMPTY_STR
    If (Not FileExists(sFileSpec)) Then Err.Raise FILE_NOT_FOUND

    ' Change the mouse pointer to an hourglass
    HourGlass True

    Dim sByLine As String, sHolder As String, sCrLf As String
    Dim lByteCount As Long, lBytesIn As Long
    Dim iFile As Byte, Idx As Byte
    iFile = FreeFile

    ' Handle errors if they occur
    On Error GoTo HandleIt
    Open sFileSpec For Binary Access Read Lock Write As #iFile

    ' Record the files byte size
    lByteCount = LOF(iFile)

    ' Set the Progress Bar to match the file byte size
    If (Not ProgressBar Is Nothing) Then ProgressBar.max = lByteCount

    ' Get the file in line by line
    Do Until (EOF(iFile))
        Line Input #iFile, sByLine

        For Idx = ONE To Len(sByLine) ' vbLf = Chr$(10)
            If (Mid(sByLine, Idx, ONE) = vbLf) Then
                Mid(sByLine, Idx) = vbCrLf & Mid(sByLine, Idx + ONE)
                Idx = Idx + TWO
            End If
        Next Idx

        sHolder = sHolder & sCrLf & sByLine

        If LenB(sCrLf) = ZERO Then sCrLf = vbCrLf
        lBytesIn = Len(sHolder)

        ' Step the Progress Bar along a little
        If (Not ProgressBar Is Nothing) Then
            ProgressBar.Value = IIf(lBytesIn < lByteCount, lBytesIn, lByteCount)
        End If
    Loop
    Close #iFile

    FixLineEnds = SaveTextFile(sFileSpec, sHolder)

    ' Reset to the previous mouse pointer
    HourGlass False
    Exit Function
HandleIt:
    Close #iFile

    ' Reset to the previous mouse pointer
    HourGlass False

    ' Return the error number (zero on success)
    FixLineEnds = Err
End Function

Public Function ValidateFiles(sFileSpecs As String) As Boolean
    If LenB(sFileSpecs) = ZERO Then Err.Raise INVALID_ARG, , EMPTY_STR
    On Error GoTo ValidateFilesError

    If (InStr(sFileSpecs, vbNullChar) = ZERO) Then
        ' Client passed a single file name
        If (Not FileExists(sFileSpecs)) Then GoTo ValidateFilesError
    Else
        ' Client passed multiple file names, Chr$(0) is the
        ' NUL character used to seperate file names
        Dim iLength As Integer, sSpec As String
        Dim idx1 As Integer, idx2 As Integer
        idx1 = ONE
        Do While idx1 < Len(sFileSpecs)
            idx2 = InStr(idx1, sFileSpecs, vbNullChar)
            If (idx2 = ZERO) Then idx2 = Len(sFileSpecs) + ONE
            iLength = idx2 - idx1
            sSpec = Mid$(sFileSpecs, idx1, iLength)
            If (Not FileExists(sSpec)) Then GoTo ValidateFilesError
            idx1 = idx2 + ONE
        Loop
    End If
    ValidateFiles = True
ValidateFilesError:
End Function

'-------------------------------------------------------------------
' To specify multiple source files they must be seperated by the NUL
' character (Chr$(0) or vbNullChar). Wildcard characters are allowed.
' This function returns True if no errors were encounted.
'-------------------------------------------------------------------
Public Function CopyFiles(sSource As String, sDestination As String, Optional ByVal cfFlags As FileOpFlags) As Boolean
    If (LenB(sSource) = ZERO) Or (LenB(sDestination) = ZERO) Then Err.Raise INVALID_ARG, , EMPTY_STR
    If (Not ValidateFiles(sSource)) Then GoTo CopyFilesError
    Dim fos As SHFILEOPSTRUCT, rc As Long, flags As Integer

    On Error GoTo CopyFilesError
    If (InStr(sDestination, vbNullChar)) Then flags = FOF_MULTIDESTFILES
    flags = flags Or cfFlags
    
    fos.pFrom = sSource & vbNullChar & vbNullChar
    fos.pTo = sDestination & vbNullChar & vbNullChar
    fos.fFlags = flags
    fos.wFunc = FO_COPY
    rc = SHFileOperationA(fos)
    CopyFiles = True

CopyFilesError:
End Function

'-------------------------------------------------------------------
' To specify multiple source files they must be seperated by the NUL
' character (Chr$(0) or vbNullChar). Wildcard characters are allowed.
' This function returns True if no errors were encounted.
'-------------------------------------------------------------------
Public Function MoveFiles(sSource As String, sDestination As String, Optional ByVal cfFlags As FileOpFlags) As Boolean
    If (LenB(sSource) = ZERO) Or (LenB(sDestination) = ZERO) Then Err.Raise INVALID_ARG, , EMPTY_STR
    If (Not ValidateFiles(sSource)) Then GoTo MoveFilesError
    Dim fos As SHFILEOPSTRUCT, rc As Long, flags As Integer

    On Error GoTo MoveFilesError
    If (InStr(sDestination, vbNullChar)) Then flags = FOF_MULTIDESTFILES
    flags = flags Or cfFlags

    fos.pFrom = sSource & vbNullChar & vbNullChar
    fos.pTo = sDestination & vbNullChar & vbNullChar
    fos.fFlags = flags
    fos.wFunc = FO_MOVE
    rc = SHFileOperationA(fos)
    MoveFiles = True

MoveFilesError:
End Function

'-------------------------------------------------------------------
' To specify multiple file names you need to have a magic wand!
' This function returns True if no errors were encounted.
'-------------------------------------------------------------------
Public Function RenameFiles(sOldName As String, sNewName As String, Optional ByVal cfFlags As FileOpFlags = NoDlgBox) As Boolean
    If (LenB(sOldName) = ZERO) Or (LenB(sNewName) = ZERO) Then Err.Raise INVALID_ARG, , EMPTY_STR
    If (Not ValidateFiles(sOldName)) Then GoTo RenameFilesError
    Dim fos As SHFILEOPSTRUCT, rc As Long, flags As Integer

    On Error GoTo RenameFilesError
    If (InStr(sNewName, vbNullChar)) Then flags = FOF_MULTIDESTFILES
    flags = flags Or cfFlags

    fos.pFrom = sOldName & vbNullChar & vbNullChar
    fos.pTo = sNewName & vbNullChar & vbNullChar
    fos.fFlags = flags
    fos.wFunc = FO_RENAME
    rc = SHFileOperationA(fos)
    RenameFiles = True

RenameFilesError:
End Function

'----------------------------------------------------------------------
' To specify multiple files to delete they must be seperated by the NUL
' character (Chr$(0) or vbNullChar). Wildcard characters are allowed.
'----------------------------------------------------------------------
Public Sub DeleteFiles(sSpec As String, Optional ToRecycleBin As Boolean = True, Optional ByVal cfFlags As FileOpFlags = NoDlgBox)
    If LenB(sSpec) = ZERO Then Err.Raise INVALID_ARG, , EMPTY_STR
    Dim fos As SHFILEOPSTRUCT, rc As Long, flags As Integer

    On Error GoTo DeleteFilesError
    If ToRecycleBin Then flags = FOF_ALLOWUNDO
    flags = flags Or cfFlags

    fos.pFrom = sSpec & vbNullChar & vbNullChar
    fos.fFlags = flags
    fos.wFunc = FO_DELETE
    rc = SHFileOperationA(fos)

DeleteFilesError:
End Sub
