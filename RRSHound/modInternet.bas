Attribute VB_Name = "modInternet"


Option Explicit

Private Declare Function FindExecutable Lib "shell32" _
   Alias "FindExecutableA" _
  (ByVal lpFile As String, _
   ByVal lpDirectory As String, _
   ByVal sResult As String) As Long

Private Declare Function GetTempPath Lib "kernel32" _
   Alias "GetTempPathA" _
  (ByVal nSize As Long, _
   ByVal lpBuffer As String) As Long

Private Const MAX_PATH As Long = 260
Private Const ERROR_FILE_NO_ASSOCIATION As Long = 31
Private Const ERROR_FILE_NOT_FOUND As Long = 2
Private Const ERROR_PATH_NOT_FOUND As Long = 3
Private Const ERROR_FILE_SUCCESS As Long = 32 'my constant
Private Const ERROR_BAD_FORMAT As Long = 11
Private Declare Function InternetGetConnectedState Lib "wininet" _
  (ByRef dwflags As Long, _
   ByVal dwReserved As Long) As Long

'Local system uses a modem to connect to the Internet.
Private Const INTERNET_CONNECTION_MODEM As Long = &H1

'Local system uses a LAN to connect to the Internet.
Private Const INTERNET_CONNECTION_LAN As Long = &H2

'Local system uses a proxy server to connect to the Internet.
Private Const INTERNET_CONNECTION_PROXY As Long = &H4

'No longer used.
Private Const INTERNET_CONNECTION_MODEM_BUSY As Long = &H8

Private Const INTERNET_RAS_INSTALLED As Long = &H10
Private Const INTERNET_CONNECTION_OFFLINE As Long = &H20
Private Const INTERNET_CONNECTION_CONFIGURED As Long = &H40

Private Declare Function URLDownloadToFile Lib "urlmon" _
   Alias "URLDownloadToFileA" _
  (ByVal pCaller As Long, _
   ByVal szURL As String, _
   ByVal szFileName As String, _
   ByVal dwReserved As Long, _
   ByVal lpfnCB As Long) As Long
   
Private Declare Function DeleteUrlCacheEntry Lib "Wininet.dll" _
   Alias "DeleteUrlCacheEntryA" _
  (ByVal lpszUrlName As String) As Long
   
Private Const ERROR_SUCCESS As Long = 0
Private Const BINDF_GETNEWESTVERSION As Long = &H10
Private Const INTERNET_FLAG_RELOAD As Long = &H80000000


Public Function isConnectedtoNet() As Boolean
   
   Dim bConn  As Boolean
   
   If bConn = False Then bConn = IsNetConnectViaLAN()
   If bConn = False Then bConn = IsNetConnectViaModem()
   If bConn = False Then bConn = IsNetConnectViaProxy()
   If bConn = False Then bConn = IsNetConnectOnline()
   If bConn = False Then bConn = IsNetRASInstalled()

    isConnectedtoNet = bConn

End Function

Public Function IsNetConnectViaLAN() As Boolean

   Dim dwflags As Long
   
  'pass an empty variable into which the API will
  'return the flags associated with the connection
   Call InternetGetConnectedState(dwflags, 0&)

  'return True if the flags indicate a LAN connection
   IsNetConnectViaLAN = dwflags And INTERNET_CONNECTION_LAN
     
End Function


Public Function IsNetConnectViaModem() As Boolean

   Dim dwflags As Long
   
  'pass an empty variable into which the API will
  'return the flags associated with the connection
   Call InternetGetConnectedState(dwflags, 0&)

  'return True if the flags indicate a modem connection
   IsNetConnectViaModem = dwflags And INTERNET_CONNECTION_MODEM
     
End Function


Public Function IsNetConnectViaProxy() As Boolean

   Dim dwflags As Long
   
  'pass an empty variable into which the API will
  'return the flags associated with the connection
   Call InternetGetConnectedState(dwflags, 0&)

  'return True if the flags indicate a proxy connection
   IsNetConnectViaProxy = dwflags And INTERNET_CONNECTION_PROXY
     
End Function


Public Function IsNetConnectOnline() As Boolean

  'no flags needed here - the API returns True
  'if there is a connection of any type
   IsNetConnectOnline = InternetGetConnectedState(0&, 0&)
     
End Function


Public Function IsNetRASInstalled() As Boolean

   Dim dwflags As Long
   
  'pass an empty variable into which the API will
  'return the flags associated with the connection
   Call InternetGetConnectedState(dwflags, 0&)

  'return True if the flags include RAS installed
   IsNetRASInstalled = dwflags And INTERNET_RAS_INSTALLED
     
End Function


Public Function GetNetConnectString() As String

   Dim dwflags As Long
   Dim msg As String

  'build a string for display
   If InternetGetConnectedState(dwflags, 0&) Then
     
      If dwflags And INTERNET_CONNECTION_CONFIGURED Then
         msg = msg & "You have a network connection configured." & vbCrLf
      End If

      If dwflags And INTERNET_CONNECTION_LAN Then
         msg = msg & "The local system connects to the Internet via a LAN"
      End If
      
      If dwflags And INTERNET_CONNECTION_PROXY Then
         msg = msg & ", and uses a proxy server. "
      Else
         msg = msg & "."
      End If
      
      If dwflags And INTERNET_CONNECTION_MODEM Then
         msg = msg & "The local system uses a modem to connect to the Internet. "
      End If
      
      If dwflags And INTERNET_CONNECTION_OFFLINE Then
         msg = msg & "The connection is currently offline. "
      End If
      
      If dwflags And INTERNET_CONNECTION_MODEM_BUSY Then
         msg = msg & "The local system's modem is busy with a non-Internet connection. "
      End If
      
      If dwflags And INTERNET_RAS_INSTALLED Then
         msg = msg & "Remote Access Services are installed on this system."
      End If
      
   Else
    
      msg = "Not connected to the internet now."
      
   End If
   
   GetNetConnectString = msg

End Function
Public Function GetBrowserName(dwFlagReturned As Long) As String

   Dim hFile As Long
   Dim sResult As String
   Dim sTempFolder As String
        
  'get the user's temp folder
   sTempFolder = GetTempDir()
   
  'create a dummy html file in the temp dir
   hFile = FreeFile
      Open sTempFolder & "dummy.html" For Output As #hFile
   Close #hFile

  'get the file path & name associated with the file
   sResult = Space$(MAX_PATH)
   dwFlagReturned = FindExecutable("dummy.html", sTempFolder, sResult)
  
  'clean up
   Kill sTempFolder & "dummy.html"
   
  'return result
   GetBrowserName = TrimNull(sResult)
   
End Function


Private Function TrimNull(Item As String)

    Dim pos As Integer
   
    pos = InStr(Item, Chr$(0))
    
    If pos Then
       TrimNull = Left$(Item, pos - 1)
    Else
       TrimNull = Item
    End If
  
End Function


Private Function GetTempDir() As String

    Dim nSize As Long
    Dim tmp As String
    
    tmp = Space$(MAX_PATH)
    nSize = Len(tmp)
    Call GetTempPath(nSize, tmp)
    
    GetTempDir = TrimNull(tmp)
    
End Function

'   If DownloadFile(sSourceUrl, sLocalFile) Then
'
'      hFile = FreeFile
'      Open sLocalFile For Input As #hFile
'         Text1.Text = Input$(LOF(hFile), hFile)
'      Close #hFile
'
'   End If

Public Function DownloadFile(sSourceUrl As String, _
                             sLocalFile As String) As Boolean
  
  'Download the file. BINDF_GETNEWESTVERSION forces
  'the API to download from the specified source.
  'Passing 0& as dwReserved causes the locally-cached
  'copy to be downloaded, if available. If the API
  'returns ERROR_SUCCESS (0), DownloadFile returns True.
    Call DeleteUrlCacheEntry(sSourceUrl)

   DownloadFile = URLDownloadToFile(0&, _
                                    sSourceUrl, _
                                    sLocalFile, _
                                    BINDF_GETNEWESTVERSION, _
                                    0&) = ERROR_SUCCESS
   
End Function
