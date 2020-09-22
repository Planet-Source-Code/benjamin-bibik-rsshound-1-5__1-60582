Attribute VB_Name = "modDroppedFile"
Option Explicit

Public Type POINTAPI
   x As Long
   y As Long
End Type
 

Public Type MSG
   hWnd As Long
   message As Long
   wParam As Long
   lParam As Long
   time As Long
   pt As POINTAPI
End Type


Public Declare Sub DragFinish Lib "Shell32" _
  (ByVal hDrop As Long)

Public Declare Function DragQueryFile Lib "Shell32" _
   Alias "DragQueryFileA" _
  (ByVal hDrop As Long, _
   ByVal UINT As Long, _
   ByVal lpStr As String, _
   ByVal ch As Long) As Long

Public Declare Function PeekMessage Lib "user32" _
   Alias "PeekMessageA" _
  (lpMsg As MSG, _
   ByVal hWnd As Long, _
   ByVal wMsgFilterMin As Long, _
   ByVal wMsgFilterMax As Long, _
   ByVal wRemoveMsg As Long) As Long

Public Const PM_NOREMOVE = &H0
Public Const PM_NOYIELD = &H2
Public Const PM_REMOVE = &H1
Public Const WM_DROPFILES = &H233


Public Sub WatchForFiles()
   
  'This watches for all WM_DROPFILES messages

   Dim FileDropMessage As MSG    'Msg Type
   Dim fileDropped As Boolean    'True if Files where dropped
   Dim hDrop As Long             'Pointer to the dropped file structure
   Dim filename As String * 128  'the dropped filename
   Dim numOfDroppedFiles As Long 'the number of dropped files
   Dim curFile As Long           'the current file number
   
  'loop to keep checking for files
  'NOTE: Do any code you want to execute before this set
   
   Do
      
      'check for Dropped file messages
       fileDropped = PeekMessage(FileDropMessage, 0, _
                     WM_DROPFILES, WM_DROPFILES, PM_REMOVE Or PM_NOYIELD)

       If fileDropped Then
         
         'get the pointer to the dropped file structure
          hDrop = FileDropMessage.wParam
         
         'get the total number of files
          numOfDroppedFiles = DragQueryFile(hDrop, True, filename, 127)

          For curFile = 1 To numOfDroppedFiles
             
             'get the file name
              Call DragQueryFile(hDrop, curFile - 1, filename, 127)
             
             'at this pointer you can do what you want with the filename
             'the filename will be a full qualified path
              'fMainForm.lblNumDropped = LTrim$(Str$(numOfDroppedFiles))
              'fMainForm.List1.AddItem filename
                MsgBox "Trying to add " & filename
          Next curFile
         
         'we are now done with the structure, tell windows to discard it
          DragFinish (hDrop)

      End If
        If fMainForm Is Nothing Then Exit Do
     'be nice
      DoEvents

   Loop

End Sub

