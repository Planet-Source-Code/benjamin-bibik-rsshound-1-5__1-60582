Attribute VB_Name = "modListView"
Option Explicit

Private Const MAX_PATH As Long = 260
Private Const MAXDWORD As Long = &HFFFFFFFF
Private Const SHGFI_DISPLAYNAME As Long = &H200
Private Const SHGFI_EXETYPE As Long = &H2000
Private Const SHGFI_TYPENAME As Long = &H400
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

Private Type POINTAPI
   x As Long
   y As Long
End Type

Private Type LV_FINDINFO
   flags       As Long
   psz         As String
   lParam      As Long
   pt          As POINTAPI
   vkDirection As Long
End Type

Private Type SHFILEINFO
   hIcon          As Long
   iIcon          As Long
   dwAttributes   As Long
   szDisplayName  As String * MAX_PATH
   szTypeName     As String * 80
End Type

Private Type FILETIME
  dwLowDateTime  As Long
  dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
  wYear          As Integer
  wMonth         As Integer
  wDayOfWeek     As Integer
  wDay           As Integer
  wHour          As Integer
  wMinute        As Integer
  wSecond        As Integer
  wMilliseconds  As Integer
End Type

Private Type WIN32_FIND_DATA
  dwFileAttributes  As Long
  ftCreationTime    As FILETIME
  ftLastAccessTime  As FILETIME
  ftLastWriteTime   As FILETIME
  nFileSizeHigh     As Long
  nFileSizeLow      As Long
  dwReserved0       As Long
  dwReserved1       As Long
  cFileName         As String * MAX_PATH
  cAlternate        As String * 14
End Type

   



    
Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long


Public Sub lvAutosizeControl(lv As ListView)

  'Size each column based on the maximum of
  'EITHER the columnheader text width, or,
  'if the items below it are wider, the
  'widest list item in the column
   Dim col2adjust As Long

   For col2adjust = 0 To lv.ColumnHeaders.Count - 1
   
      Call SendMessage(lv.hWnd, _
                       LVM_SETCOLUMNWIDTH, _
                       col2adjust, _
                       ByVal LVSCW_AUTOSIZE_USEHEADER)
   Next
   
   
End Sub


Public Sub lvAutosizeItems(lv As ListView)

  'Size each column based on the width
  'of the widest list item in the column.
  'If the items are shorter than the column
  'header text, the header text is truncated.

  'You may need to lengthen column header
  'captions to see this effect.
   Dim col2adjust As Long

   For col2adjust = 0 To lv.ColumnHeaders.Count - 1
   
      Call SendMessage(lv.hWnd, _
                       LVM_SETCOLUMNWIDTH, _
                       col2adjust, _
                       ByVal LVSCW_AUTOSIZE)
   Next
   
End Sub

Public Sub lvAutosizeMax(lv As ListView)
   
  'Because applying the LVSCW_AUTOSIZE_USEHEADER
  'message to the last column in the control always
  'sets its width to the maximum remaining control
  'space, calling SendMessage passing the last column
  'will cause the listview data to utilize the full
  'control width space. For example, if a four-column
  'listview had a total width of 2000, and the first
  'three columns each had individual widths of 250,
  'calling this will cause the last column to widen
  'to cover the remaining 1250.

  'For this message to (visually) work as expected,
  'all columns should be within the viewing rect of the
  'listview control; if the last column is wider than
  'the control the message works, but the columns
  'remain wider than the control.
   Dim col2adjust As Long
   
   col2adjust = lv.ColumnHeaders.Count - 1
   
   Call SendMessage(lv.hWnd, _
            LVM_SETCOLUMNWIDTH, _
            col2adjust, _
            ByVal LVSCW_AUTOSIZE_USEHEADER)
   
End Sub


