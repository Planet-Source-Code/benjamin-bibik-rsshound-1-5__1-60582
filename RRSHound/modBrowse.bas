Attribute VB_Name = "modBrowse"
Option Explicit
Public Const INVALID_HANDLE_VALUE = -1
Public Const MAX_PATH = 260

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Public Declare Function FindFirstFile _
               Lib "kernel32" _
               Alias "FindFirstFileA" (ByVal lpFileName As String, _
                                       lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function FindClose _
               Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function PathFindFileName _
                Lib "shlwapi" _
                Alias "PathFindFileNameA" (ByVal pPath As String) As Long
  
Private Declare Function lstrcpyA _
                Lib "kernel32" (ByVal RetVal As String, _
                                ByVal Ptr As Long) As Long
                        
Private Declare Function lstrlenA _
                Lib "kernel32" (ByVal Ptr As Any) As Long

Private Type BROWSEINFO
   hOwner           As Long
   pidlRoot         As Long
   pszDisplayName   As String
   lpszTitle        As String
   ulFlags          As Long
   lpfn             As Long
   lParam           As Long
   iImage           As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000
'Private Const MAX_PATH = 260

Private Declare Function SHGetPathFromIDList Lib "shell32" _
   Alias "SHGetPathFromIDListA" _
  (ByVal pidl As Long, _
   ByVal pszPath As String) As Long

Private Declare Function SHBrowseForFolder Lib "shell32" _
   Alias "SHBrowseForFolderA" _
  (lpBrowseInfo As BROWSEINFO) As Long

Private Declare Sub CoTaskMemFree Lib "ole32" _
   (ByVal pv As Long)

Public Function BrowseForFolder(ByVal hWnd As Long, szTitle As String) As String
Dim bi As BROWSEINFO
  Dim pidl As Long
  Dim path As String
  Dim pos As Long
    
 'Fill the BROWSEINFO structure with the
 'needed data. To accommodate comments, the
 'With/End With syntax has not been used, though
 'it should be your 'final' version.

   With bi
      
     'hwnd of the window that receives messages
     'from the call. Can be your application
     'or the handle from GetDesktopWindow()
      .hOwner = hWnd

     'pointer to the item identifier list specifying
     'the location of the "root" folder to browse from.
     'If NULL, the desktop folder is used.
     .pidlRoot = 0&

     'message to be displayed in the Browse dialog
     .lpszTitle = szTitle

     'the type of folder to return.
      .ulFlags = BIF_RETURNONLYFSDIRS
   End With
    
  'show the browse for folders dialog
   pidl = SHBrowseForFolder(bi)
 
  'the dialog has closed, so parse & display the
  'user's returned folder selection contained in pidl
   path = Space$(MAX_PATH)
    
   If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
      pos = InStr(path, Chr$(0))
      BrowseForFolder = Left(path, pos - 1)
   End If

   Call CoTaskMemFree(pidl)

End Function
Public Function FileExists(sSource As String) As Boolean

    Dim WFD As WIN32_FIND_DATA
    Dim hFile As Long
   
    hFile = FindFirstFile(sSource, WFD)
    FileExists = hFile <> INVALID_HANDLE_VALUE
   
    Call FindClose(hFile)
   
End Function

Public Function GetFilePart(ByVal sPath As String) As String
  
    GetFilePart = GetStrFromPtrA(PathFindFileName(sPath))
   
End Function

Private Function GetStrFromPtrA(ByVal lpszA As Long) As String

    GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
    Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)
   
End Function

Private Function TrimNull(Item As String)

    Dim pos As Integer
   
    pos = InStr(Item, Chr$(0))
   
    If pos Then
        TrimNull = Left$(Item, pos - 1)
        Else: TrimNull = Item
    End If
   
End Function



