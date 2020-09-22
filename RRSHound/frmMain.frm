VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmMain 
   Caption         =   "RSSHound"
   ClientHeight    =   6945
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   6945
   ScaleWidth      =   11760
   StartUpPosition =   3  'Windows Default
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   855
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   1508
      BandCount       =   2
      _CBWidth        =   11760
      _CBHeight       =   855
      _Version        =   "6.7.8988"
      Child1          =   "tbToolBar"
      MinHeight1      =   390
      Width1          =   6180
      NewRow1         =   0   'False
      Child2          =   "picSearch"
      MinHeight2      =   375
      Width2          =   4995
      NewRow2         =   -1  'True
      Begin VB.PictureBox picSearch 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   165
         ScaleHeight     =   375
         ScaleWidth      =   11505
         TabIndex        =   11
         Top             =   450
         Width           =   11505
         Begin VB.ComboBox cboSearchIn 
            Height          =   315
            ItemData        =   "frmMain.frx":0000
            Left            =   5040
            List            =   "frmMain.frx":000D
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   30
            Width           =   2295
         End
         Begin VB.CommandButton cmdGO 
            Caption         =   "Find Now"
            Height          =   255
            Left            =   7500
            TabIndex        =   14
            Top             =   60
            Width           =   855
         End
         Begin VB.TextBox txtSearch 
            Height          =   285
            Left            =   840
            TabIndex        =   13
            Top             =   30
            Width           =   3195
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Look for"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   30
            TabIndex        =   15
            Top             =   60
            Width           =   660
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Search In"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   4125
            TabIndex        =   12
            Top             =   60
            Width           =   780
         End
      End
      Begin MSComctlLib.Toolbar tbToolBar 
         Height          =   390
         Left            =   165
         TabIndex        =   10
         Top             =   30
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         ImageList       =   "imlToolbarIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "New"
               Object.ToolTipText     =   "New"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Delete"
               Object.ToolTipText     =   "Delete"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Properties"
               Object.ToolTipText     =   "Properties"
               ImageKey        =   "Properties"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "refresh"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "history"
               ImageIndex      =   9
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSComctlLib.ImageList imgTree 
      Left            =   840
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0028
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":013A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":02A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":03B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":04CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":05DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":06EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0800
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0912
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A24
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B36
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvFeeds 
      Height          =   2055
      Left            =   60
      TabIndex        =   8
      Top             =   600
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   3625
      _Version        =   393217
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgTree"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.PictureBox picRight 
      Height          =   6315
      Left            =   2160
      ScaleHeight     =   6255
      ScaleWidth      =   9435
      TabIndex        =   2
      Top             =   660
      Width           =   9495
      Begin VB.PictureBox lvFeeds 
         BackColor       =   &H00FFFFFF&
         Height          =   1695
         Left            =   300
         ScaleHeight     =   1635
         ScaleWidth      =   4515
         TabIndex        =   17
         Top             =   360
         Width           =   4575
         Begin VB.PictureBox picFeedHeader 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   555
            Left            =   0
            ScaleHeight     =   555
            ScaleWidth      =   3915
            TabIndex        =   21
            Top             =   0
            Width           =   3915
            Begin VB.TextBox lblDescription 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   375
               Left            =   1140
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   22
               Top             =   60
               Width           =   1335
            End
            Begin VB.Image imgFeed 
               Height          =   435
               Left            =   0
               Top             =   0
               Width           =   495
            End
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   1275
            LargeChange     =   735
            Left            =   4260
            Max             =   1000
            SmallChange     =   100
            TabIndex        =   20
            Top             =   0
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.PictureBox picFeeds 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1035
            Left            =   0
            ScaleHeight     =   1035
            ScaleWidth      =   3735
            TabIndex        =   18
            Top             =   600
            Width           =   3735
            Begin RSSHound.FeedList Feed 
               Height          =   735
               Index           =   0
               Left            =   0
               TabIndex        =   19
               Top             =   0
               Visible         =   0   'False
               Width           =   4275
               _ExtentX        =   6165
               _ExtentY        =   1296
            End
         End
      End
      Begin VB.PictureBox PicSplitH 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         FillColor       =   &H00808080&
         Height          =   4800
         Left            =   60
         ScaleHeight     =   2090.126
         ScaleMode       =   0  'User
         ScaleWidth      =   780
         TabIndex        =   5
         Top             =   1380
         Visible         =   0   'False
         Width           =   72
      End
      Begin VB.PictureBox picBrowser 
         BorderStyle     =   0  'None
         Height          =   3915
         Left            =   60
         ScaleHeight     =   3915
         ScaleWidth      =   9585
         TabIndex        =   3
         Top             =   2280
         Width           =   9585
         Begin VB.PictureBox wbHeader 
            BackColor       =   &H00FFFFFF&
            Height          =   3855
            Left            =   180
            ScaleHeight     =   3795
            ScaleWidth      =   8715
            TabIndex        =   6
            Top             =   -60
            Width           =   8775
            Begin VB.Image Image1 
               Height          =   1125
               Left            =   960
               Picture         =   "frmMain.frx":0C48
               Top             =   1500
               Width           =   7500
            End
            Begin VB.Label lblLink 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "View Feed"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   225
               Left            =   7740
               TabIndex        =   7
               Top             =   60
               Visible         =   0   'False
               Width           =   870
            End
         End
         Begin SHDocVwCtl.WebBrowser wb 
            Height          =   975
            Left            =   3480
            TabIndex        =   4
            Top             =   1200
            Visible         =   0   'False
            Width           =   3615
            ExtentX         =   6376
            ExtentY         =   1720
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   0
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "http:///"
         End
      End
      Begin VB.Image imgSplitH 
         Height          =   3105
         Left            =   360
         MousePointer    =   7  'Size N S
         Top             =   2580
         Width           =   150
      End
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   3600
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   1
      Top             =   705
      Visible         =   0   'False
      Width           =   72
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   6675
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15081
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "5/18/2005"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "2:01 PM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   7800
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5A00
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5B12
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5C24
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5D36
            Key             =   "View Large Icons"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5E48
            Key             =   "View Small Icons"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5F5A
            Key             =   "View List"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":606C
            Key             =   "View Details"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":617E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":62D8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgSplitter 
      Height          =   2145
      Left            =   1440
      MousePointer    =   9  'Size W E
      Top             =   2940
      Width           =   210
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFilemnuFileNewFeed 
         Caption         =   "New Feed"
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Propert&ies"
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewHistory 
         Caption         =   "History"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "&Search For Help On..."
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
   Begin VB.Menu mnuGroup 
      Caption         =   "GroupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuGroupAdd 
         Caption         =   "Add Group"
      End
      Begin VB.Menu mnuGroupEdit 
         Caption         =   "Edit Group Name"
      End
      Begin VB.Menu mnuGroupDelete 
         Caption         =   "Delete Group"
      End
      Begin VB.Menu mnuGroupSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGroupAddFeed 
         Caption         =   "Add Feed"
      End
      Begin VB.Menu mnuGroupEditFeed 
         Caption         =   "Edit Feed"
      End
      Begin VB.Menu mnuGroupUnsubscribe 
         Caption         =   "Unscubscribe Feed"
      End
   End
   Begin VB.Menu mnuFeeds 
      Caption         =   "Feeds"
      Visible         =   0   'False
      Begin VB.Menu mnuFeedOpenInBrowser 
         Caption         =   "Open in browser"
      End
   End
   Begin VB.Menu mnuHistory 
      Caption         =   "History"
      Begin VB.Menu mnuHistoryDelete 
         Caption         =   "Delete from History"
      End
      Begin VB.Menu mnuHistoryClear 
         Caption         =   "Clear History"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
  
Dim mbMoving As Boolean
Const sglSplitLimit = 500
Dim bloading As Boolean
Dim mOpenURL As String

Private Const CREATE_NEW_CONSOLE As Long = &H10
Private Const NORMAL_PRIORITY_CLASS As Long = &H20
Private Const INFINITE As Long = -1
Private Const STARTF_USESHOWWINDOW As Long = &H1
Private Const SW_SHOWNORMAL As Long = 1

Private Const MAX_PATH As Long = 260
Private Const ERROR_FILE_NO_ASSOCIATION As Long = 31
Private Const ERROR_FILE_NOT_FOUND As Long = 2
Private Const ERROR_FILE_SUCCESS As Long = 32 'my constant

Private Type STARTUPINFO
  cb As Long
  lpReserved As String
  lpDesktop As String
  lpTitle As String
  dwX As Long
  dwY As Long
  dwXSize As Long
  dwYSize As Long
  dwXCountChars As Long
  dwYCountChars As Long
  dwFillAttribute As Long
  dwFlags As Long
  wShowWindow As Integer
  cbReserved2 As Integer
  lpReserved2 As Long
  hStdInput As Long
  hStdOutput As Long
  hStdError As Long
End Type

Private Type PROCESS_INFORMATION
  hProcess As Long
  hThread As Long
  dwProcessId As Long
  dwThreadID As Long
End Type

Private Declare Function CreateProcess Lib "kernel32" _
   Alias "CreateProcessA" _
  (ByVal lpAppName As String, _
   ByVal lpCommandLine As String, _
   ByVal lpProcessAttributes As Long, _
   ByVal lpThreadAttributes As Long, _
   ByVal bInheritHandles As Long, _
   ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, _
   ByVal lpCurrentDirectory As Long, _
   lpStartupInfo As STARTUPINFO, _
   lpProcessInformation As PROCESS_INFORMATION) As Long
     
Private Declare Function CloseHandle Lib "kernel32" _
  (ByVal hObject As Long) As Long

Private Declare Function FindExecutable Lib "Shell32" _
   Alias "FindExecutableA" _
  (ByVal lpFile As String, _
   ByVal lpDirectory As String, _
   ByVal sResult As String) As Long

Private Declare Function GetTempPath Lib "kernel32" _
   Alias "GetTempPathA" _
  (ByVal nSize As Long, _
   ByVal lpBuffer As String) As Long
         
Private bInBrowser As Boolean
Private bViewHistory As Boolean
Private SelectedNode As MSComctlLib.Node

Private Sub cmdGO_Click()
    
    Dim mli As MSComctlLib.ListItem
    Dim rs As New ADODB.Recordset
    Dim oNodelist As IXMLDOMNodeList
    Dim oNode As IXMLDOMElement
    
    'lvFeeds.ListItems.Clear

    If txtSearch.Text = "" Then Exit Sub
    
    If cboSearchIn.ListIndex = 0 Or cboSearchIn.ListIndex = 1 Then
        rs.Open "", cn
        Do Until rs.EOF
        Loop
        rs.Close
    End If
    
    If cboSearchIn.ListIndex = 0 Or cboSearchIn.ListIndex = 2 Then
    End If

End Sub

Private Sub CoolBar1_HeightChanged(ByVal NewHeight As Single)
    Form_Resize
End Sub

Private Sub Feed_ItemClick(Index As Integer, _
                           URL As String)

    Dim rs As New ADODB.Recordset
    Dim feedName As String
    Dim feedid As Long
    Dim GroupID As Long
    Dim iCount As Long
    
    For iCount = 1 To Feed.Count - 1
        Feed(iCount).Selected = False
    Next

    Feed(Index).Selected = True
    
    'wb.Navigate Item.Tag
    
    wb.Visible = True
    wbHeader.Visible = False
    
    Feed(Index).Read = True

    If GetOption("keepinhistory", -1) Then
        
        cn.Execute "DELETE FROM History WHERE DateViewed <=#" & DateAdd("d", -(GetOption("daysinhistory", 20)), Now) & "#"
        
        rs.Open "Select Feedname, Feedid, GroupId from Feeds where Feedid = " & lvFeeds.Tag, cn, adOpenDynamic, adLockOptimistic
    
        If Not (rs.EOF And rs.BOF) Then
            feedName = rs("FeedName")
            feedid = rs("Feedid")
            GroupID = rs("Groupid")
        End If
        
        rs.Close
       
        rs.Open "SELECT [History].HistoryID, [History].FeedID, " & _
                "[History].FeedName, [History].[Extract], [History].[Title], " & _
                "[History].ArticleURL, [History].DateViewed, " & _
                "[History].GroupID " & _
                "From [History] " & _
                "Where [History].ArticleURL = '" & Feed(Index).URL & "'", cn, adOpenDynamic, adLockOptimistic
        
        If rs.EOF And rs.BOF Then
            rs.AddNew
            rs("FeedID") = feedid
            rs("FeedName") = feedName
            rs("Extract") = Left(Feed(Index).Description, 255)
            rs("Title") = Feed(Index).Caption
            rs("ArticleURL") = Feed(Index).URL
            rs("DateViewed") = Date
            rs("GroupID") = GroupID
            rs.Update
        End If
    End If

    rs.Close
        
    If bInBrowser Then
        StartNewBrowser Feed(Index).URL
    Else
        wb.Navigate Feed(Index).URL
    End If
    '    wbHeader_Resize

End Sub

Private Sub Form_Load()
    'Dim mCol As MSComctlLib.ColumnHeader
    
    'For Each mCol In lvFeeds.ColumnHeaders
    '    mCol.Width = GetSetting(App.Title, "Settings", "col" & mCol.Index, 1400)
    'Next
    
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    imgSplitter.Left = GetSetting(App.Title, "Settings", "vsplit", imgSplitter.Left)
    imgSplitH.Top = GetSetting(App.Title, "Settings", "hsplit", imgSplitH.Top)

    bInBrowser = GetOption("openinbrowser", 0)
    checkInBrowser
    cboSearchIn.ListIndex = 0
    
    'wb.Navigate App.path & "\images\index.htm"
    wb.Navigate "about:blank"
    If cn.State = 1 Then cn.Close
    cn.Provider = "Microsoft.Jet.OLEDB.4.0"
    cn.Open App.path & "\rsshound.mdb"
    
    LoadSources
    
    If isConnectedtoNet() = True Then
        sbStatusBar.Panels(1).Text = "Online Mode" 'GetNetConnectString()
    Else
        sbStatusBar.Panels(1).Text = "Off-line Mode" ' GetNetConnectString()
    End If

 

End Sub

Public Sub checkInBrowser()

    If bInBrowser Then
        wb.Visible = False
    Else
        wb.Visible = False
    End If
    
    SizeSubControls imgSplitH.Top
    
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim rss As String
    Dim frm As frmAddContent
    
    rss = Data.Getdata(1)
    
    Set frm = New frmAddContent
    
    Load frm
    
    frm.txtURL.Text = rss
    frm.Getdata
    frm.Show vbModal
    
    If Not frm.Cancel Then
        LoadSources
    End If
    
    Unload frm
    
    Set frm = Nothing
    
    
End Sub

Private Sub Form_Paint()
    'lvFeeds.View = Val(GetSetting(App.Title, "Settings", "ViewMode", "0"))
'    Select Case lvFeeds.View
'        Case lvwIcon
'     '       tbToolBar.Buttons(LISTVIEW_MODE0).Value = tbrPressed
'        Case lvwSmallIcon
'     '       tbToolBar.Buttons(LISTVIEW_MODE1).Value = tbrPressed
'        Case lvwList
'     '       tbToolBar.Buttons(LISTVIEW_MODE2).Value = tbrPressed
'        Case lvwReport
'     '       tbToolBar.Buttons(LISTVIEW_MODE3).Value = tbrPressed
'    End Select
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    Dim mCol As MSComctlLib.ColumnHeader


    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    'SaveSetting App.Title, "Settings", "ViewMode", lvFeeds.View
    SaveSetting App.Title, "Settings", "vsplit", imgSplitter.Left
    SaveSetting App.Title, "Settings", "hsplit", imgSplitH.Top
    
'    For Each mCol In lvFeeds.ColumnHeaders
'        SaveSetting App.Title, "Settings", "col" & mCol.Index, mCol.Width
'    Next
    
    Set fMainForm = Nothing
    
End Sub



Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < 3000 Then Me.Width = 3000
    SizeControls imgSplitter.Left
End Sub


Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub


Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sglPos As Single
    

    If mbMoving Then
        sglPos = x + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub


Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
    
End Sub


Private Sub lvFeeds_DblClick()
    
'    If lvFeeds.SelectedItem Is Nothing Then Exit Sub
'
'    If bInBrowser Then
'        StartNewBrowser (lvFeeds.SelectedItem.Tag)
'    Else
'        wbHeader.Visible = False
'        wb.Visible = True
'        wb.Navigate lvFeeds.SelectedItem.Tag
'        lblLink.Visible = False
'    End If


End Sub

Private Sub lvFeeds_Resize()
    
    On Error Resume Next
    
    Dim oLV As FeedList
    Dim iCount As Long
    Dim cTotHeight As Double
    
    For iCount = 1 To Feed.Count - 1
        cTotHeight = Feed(iCount).Height + cTotHeight
    Next
    picFeedHeader.Move 0, 0, lvFeeds.ScaleWidth
        
    picFeeds.Height = cTotHeight
    VScroll1.Move lvFeeds.ScaleWidth - VScroll1.Width, picFeedHeader.Height, VScroll1.Width, lvFeeds.ScaleHeight - picFeedHeader.Height
    
    If cTotHeight > lvFeeds.ScaleHeight Then
        picFeeds.Width = lvFeeds.ScaleWidth - VScroll1.Width
        VScroll1.Visible = True
    Else
        picFeeds.Width = lvFeeds.ScaleWidth
        VScroll1.Visible = False
    End If
    
    VScroll1.min = picFeedHeader.Height
    VScroll1.max = (lvFeeds.ScaleHeight - picFeedHeader.Height) - picFeeds.Height
    
End Sub

Private Sub mnuGroupAdd_Click()

    Dim rs As New ADODB.Recordset
    Dim sGroupName As String
    
    sGroupName = InputBox("Enter the name of the new group", "Group Name", "[New Group]")
    
    If sGroupName = "[New Group]" Or sGroupName = "" Then
        MsgBox "No Group Added", vbInformation, "No Action"
    Else
        rs.Open "Select * from Groups", cn, adOpenDynamic, adLockOptimistic
        
        rs.AddNew "GroupText", sGroupName
        rs.Update
        rs.Close
        
        LoadSources
        
    End If
    

End Sub

Private Sub mnuGroupAddFeed_Click()

    mnuFilemnuFileNewFeed_Click

End Sub

Private Sub mnuGroupDelete_Click()

    Dim sGroupName As String
    Dim sGroupId As Long

    If tvFeeds.SelectedItem Is Nothing Then Exit Sub
    
    If tvFeeds.SelectedItem.Parent Is Nothing Then Exit Sub
    
    If tvFeeds.SelectedItem.Tag = "" Then
        sGroupName = tvFeeds.SelectedItem.Text
        sGroupId = Mid$(tvFeeds.SelectedItem.Key, 2)
    Else
        sGroupName = tvFeeds.SelectedItem.Parent.Text
        sGroupId = Mid$(tvFeeds.SelectedItem.Parent.Key, 2)
    End If
    
    If MsgBox("****************WARNING******************" & vbCrLf & vbCrLf & "All feeds in this group will also be deleted.  Are you sure you want to do this?", vbExclamation + vbYesNoCancel, "Delete group?") = vbYes Then
        cn.Execute "Delete from feeds where groupid = " & sGroupId
        cn.Execute "delete from Groups where groupid = " & sGroupId
        LoadSources
    End If

End Sub

Private Sub mnuGroupEdit_Click()

    Dim sGroupName As String
    Dim sGroupId As Long

    If tvFeeds.SelectedItem Is Nothing Then Exit Sub
    
    If tvFeeds.SelectedItem.Parent Is Nothing Then Exit Sub
    
    If tvFeeds.SelectedItem.Tag = "" Then
        sGroupName = tvFeeds.SelectedItem.Text
        sGroupId = Mid$(tvFeeds.SelectedItem.Key, 2)
    Else
        sGroupName = tvFeeds.SelectedItem.Parent.Text
        sGroupId = Mid$(tvFeeds.SelectedItem.Parent.Key, 2)
    End If
    
    sGroupName = InputBox("Change group name", "Group Edit", sGroupName)
    
    cn.Execute "UPDATE Groups set grouptext = '" & sGroupName & "' where groupid = " & sGroupId
    LoadSources
    

End Sub

Private Sub mnuGroupEditFeed_Click()

    Dim frm As frmAddContent
    
    If tvFeeds.SelectedItem Is Nothing Then Exit Sub
    
    If tvFeeds.SelectedItem.Tag = "" Then Exit Sub
    
    Set frm = New frmAddContent
    
    Load frm
    
    frm.LoadFeed tvFeeds.SelectedItem.Tag
    
    frm.Show vbModal
    
    If frm.Cancel = False Then
        LoadSources
    End If
    
    Unload frm
    
    Set frm = Nothing

End Sub

Private Sub mnuGroupUnsubscribe_Click()

    If tvFeeds.SelectedItem Is Nothing Then Exit Sub
    
    If MsgBox("Are you sure you want to unsubscribe from " & tvFeeds.SelectedItem.Text & "?", vbQuestion + vbYesNo, "Unsubscribe") = vbYes Then
        cn.Execute "UPDATE FEEDS set Subscribed = 0 where feedid = " & tvFeeds.SelectedItem.Tag
        LoadSources
    End If
    

End Sub

Private Sub mnuHistoryClear_Click()
    
    If MsgBox("Are you sure you want to clear the history?", vbQuestion + vbYesNo, "Clear History?") = vbYes Then
        cn.Execute "Delete from History"
    End If
    
End Sub

Private Sub mnuHistoryDelete_Click()

    Dim vArray As Variant
    Dim sSQL As String
    Dim i As Long

    If SelectedNode Is Nothing Then Exit Sub
    If SelectedNode.children = 0 Then
        cn.Execute "Delete from History where Historyid = " & Mid$(SelectedNode.Key, 3)
        tvFeeds.Nodes.Remove SelectedNode.Key
        Set SelectedNode = Nothing
        Exit Sub
    End If
    
    vArray = Split(SelectedNode.Key, "|")
    
    For i = 1 To UBound(vArray)
        If i = 1 Then
            sSQL = "DateViewed = #" & vArray(i) & "# "
        ElseIf i = 2 Then
            sSQL = sSQL & " and GroupID = " & vArray(i)
        ElseIf i = 3 Then
            sSQL = sSQL & " and FeedID = " & vArray(i)
        End If
        
    Next
    
    cn.Execute "Delete from History where " & sSQL
    
    LoadSources
   'Stop

End Sub

Private Sub picFeedHeader_Resize()
    On Error Resume Next
    lblDescription.Move lblDescription.Left, 0, picFeedHeader.ScaleWidth - lblDescription.Left, picFeedHeader.ScaleWidth
End Sub

Private Sub picFeeds_Resize()
    On Error Resume Next
    Dim iCount As Long
    
    For iCount = 1 To Feed.Count - 1
        Feed(iCount).Width = picFeeds.ScaleWidth
    Next

End Sub

Private Sub tbToolBar_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    Call Form_OLEDragDrop(Data, Effect, Button, Shift, x, y)

End Sub

Private Sub tvFeeds_DragDrop(Source As Control, x As Single, y As Single)
    If Source = imgSplitter Then
        SizeControls x
    End If
End Sub


Sub SizeControls(x As Single)
    On Error Resume Next
    

    'set the width
    If x < 1500 Then x = 1500
    If x > (Me.Width - 1500) Then x = Me.Width - 1500
    tvFeeds.Width = x
    imgSplitter.Left = x
    picRight.Left = x + 40
    picRight.Width = Me.Width - (tvFeeds.Width + 140)

    'set the top

    If tbToolBar.Visible Then
        tvFeeds.Top = CoolBar1.Height
    Else
        tvFeeds.Top = 0
    End If

    picRight.Top = tvFeeds.Top
    

    'set the height
    If sbStatusBar.Visible Then
        tvFeeds.Height = Me.ScaleHeight - (sbStatusBar.Height + CoolBar1.Height)
    Else
        tvFeeds.Height = Me.ScaleHeight
    End If
    

    picRight.Height = tvFeeds.Height
    
    imgSplitter.Top = tvFeeds.Top
    
    imgSplitter.Height = tvFeeds.Height
    
End Sub



Private Sub lblLink_Click()
    
    If bInBrowser Then
        StartNewBrowser (lblLink.Tag)
    Else
        wbHeader.Visible = False
        wb.Visible = True
        wb.Navigate lblLink.Tag
        lblLink.Visible = False
    End If
    
End Sub





Private Sub mnuFeedOpenInBrowser_Click()

    If mOpenURL = "" Then Exit Sub
    
    StartNewBrowser (mOpenURL)
    
    mOpenURL = ""

End Sub

Private Sub mnuViewHistory_Click()

    If mnuViewHistory.Checked = False Then
        mnuViewHistory.Checked = True
    Else
        mnuViewHistory.Checked = False
    End If
    
    bViewHistory = mnuViewHistory.Checked
    SizeSubControls imgSplitH.Top
    LoadSources
    
    
End Sub


Private Sub picBrowser_Resize()

    On Error Resume Next
    wbHeader.Move 0, 0, picBrowser.ScaleWidth, picBrowser.ScaleHeight
    wb.Move 0, 0, picBrowser.ScaleWidth, picBrowser.ScaleHeight

End Sub

Private Sub picRight_Resize()
    On Error Resume Next
    SizeSubControls imgSplitH.Top
   ' lvAutosizeMax lvFeeds
End Sub

Private Sub SizeSubControls(y As Single)
    
    On Error Resume Next
    
  '  wbHeader.Top = 0
    imgSplitH.Height = 100
    imgSplitH.Width = picRight.ScaleWidth
    'set the width
    If y < 1500 Then y = 1500
    If y > (picRight.ScaleHeight - (1500)) Then y = picRight.ScaleHeight - (1500)
    
    If Not bViewHistory Then
        lvFeeds.Visible = True
        If bInBrowser Then
            imgSplitH.Visible = False
            picBrowser.Visible = False
            wb.Visible = False
            lvFeeds.Top = 0
            lvFeeds.Height = picRight.ScaleHeight
        Else
            imgSplitH.Visible = True
            picBrowser.Visible = True
            wb.Visible = True
            lvFeeds.Top = 0
            lvFeeds.Height = y
        '    wbHeader.Height = Y
            
            imgSplitH.Top = y
            PicSplitH.Top = y
            PicSplitH.Height = 100
            PicSplitH.Width = picRight.ScaleWidth
            picBrowser.Top = y + 40
            picBrowser.Height = picRight.ScaleHeight - ((lvFeeds.Height) + 140)
            picBrowser.Left = 0
            lvFeeds.Left = 0 'wbHeader.Width + 40
            lvFeeds.Width = picRight.ScaleWidth '- (wbHeader.Width + 40)
            picBrowser.Width = picRight.ScaleWidth
        End If
    Else
        picBrowser.Visible = True
        lvFeeds.Visible = False
        If bInBrowser Then
            wb.Visible = False
            wbHeader.Visible = True
            wbHeader.Move 0, 0, picBrowser.ScaleWidth, picBrowser.ScaleHeight
            picBrowser.Move 0, 0, picRight.ScaleWidth, picRight.ScaleHeight
        Else
            wb.Visible = True
            wbHeader.Visible = False
            picBrowser.Move 0, 0, picRight.ScaleWidth, picRight.ScaleHeight
        End If
    End If
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            'ToDo: Add 'New' button code.
            mnuFilemnuFileNewFeed_Click
        Case "Delete"
            mnuFileDelete_Click
        Case "Properties"
            mnuFileProperties_Click

        Case "refresh"
            RefreshFeeds
            
        Case "history"
            mnuViewHistory_Click
    End Select
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hWnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hWnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub


Private Sub mnuToolsOptions_Click()
    frmOptions.Show vbModal, Me
    bInBrowser = GetOption("openinbrowser", 0)
    checkInBrowser
End Sub




Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
    SizeControls imgSplitter.Left
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    'tbToolBar.Visible = mnuViewToolbar.Checked
    CoolBar1.Bands(1).Visible = mnuViewToolbar.Checked
    SizeControls imgSplitter.Left
End Sub

Private Sub mnuFileClose_Click()
    'unload the form
    Unload Me

End Sub

Private Sub mnuFileProperties_Click()

    Dim frm As frmFeeds
    
    Set frm = New frmFeeds
    
    Load frm
    
    frm.Show vbModal
    
    If Not frm.Cancel Then
        LoadSources
    End If
    
    Unload frm
    
    Set frm = Nothing

End Sub



Private Sub mnuFileDelete_Click()
    'ToDo: Add 'mnuFileDelete_Click' code.
    MsgBox "Add 'mnuFileDelete_Click' code."
End Sub


Private Sub mnuFilemnuFileNewFeed_Click()

    Dim frm As frmAddContent
    
    Set frm = New frmAddContent
    
    Load frm
    
    frm.Show vbModal
    
    If Not frm.Cancel Then
        LoadSources
    End If

    Unload frm
    
    Set frm = Nothing
    
End Sub

Public Function LoadSources()

    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim xmlServer As New XMLHTTP40
    Dim iFreeFile As Long
    Dim sPicture As String
    Dim oNode As MSComctlLib.Node
    Dim sParent As String
    Dim sKey As String
    Dim mCol As Collection
    Dim obj As Object
    Dim sDateKey As String
    Dim sHistKey As String
    

    Set mCol = New Collection
    

        If tvFeeds.Nodes.Count > 0 Then
            For Each oNode In tvFeeds.Nodes
                If oNode.Visible Then
                    mCol.Add oNode.Key, oNode.Key
                End If
            Next
        End If

    
    tvFeeds.Nodes.Clear
        
    If bViewHistory = False Then
        tvFeeds.Nodes.Add , , "Personal", "Personal", 7, 7
        
        rs.Open "SELECT Groups.GroupId, Groups.GroupText FROM Groups WHERE CustomGroup = -1 ORDER BY Groups.GroupText", cn
        
        Do Until rs.EOF
            sKey = "C" & rs("Groupid")
            tvFeeds.Nodes.Add "Personal", tvwChild, sKey, rs("GroupText"), 1, 1
                    
                    rs2.Open "SELECT Feeds.FeedID, Feeds.FeedName, feeds.feedimageurl FROM Feeds " & _
                            "WHERE Feeds.CustomId=" & rs("Groupid") & " AND Feeds.Subscribed=-1", cn, adOpenDynamic, adLockOptimistic
                    Do Until rs2.EOF
                        Set oNode = tvFeeds.Nodes.Add(sKey, tvwChild, "CF" & rs2("Feedid"), rs2("FeedName"), 8, 8)
                        oNode.Tag = rs2("FeedID")
                        oNode.EnsureVisible
                        rs2.MoveNext
                    Loop
                    
                    rs2.Close
    
            
            rs.MoveNext
        Loop
        
        rs.Close
        
        tvFeeds.Nodes.Add , , "Standard", "Standard Feeds", 4, 4
        
        rs.Open "SELECT Groups.GroupId, Groups.GroupText FROM Groups WHERE CustomGroup = 0 ORDER BY Groups.GroupText", cn
        
        Do Until rs.EOF
            
            sKey = "S" & rs("Groupid")
            tvFeeds.Nodes.Add "Standard", tvwChild, sKey, rs("GroupText"), 5, 5
                    
                    rs2.Open "SELECT Feeds.FeedID, Feeds.FeedName, feeds.feedimageurl FROM Feeds " & _
                            "WHERE Feeds.GroupID=" & rs("Groupid") & " AND Feeds.Subscribed=-1", cn, adOpenDynamic, adLockOptimistic
                    Do Until rs2.EOF
                        Set oNode = tvFeeds.Nodes.Add(sKey, tvwChild, "SF" & rs2("Feedid"), rs2("FeedName"), 8, 8)
                        oNode.Tag = rs2("FeedID")
                        rs2.MoveNext
                    Loop
                    
                    rs2.Close
                
            
            rs.MoveNext
        Loop
        
        rs.Close
        
        Dim i As Long
    

    
    Else
        
        tvFeeds.Nodes.Add , , "History", "History", 4, 4
        
        rs.Open "SELECT History.DateViewed, Groups.GroupId, Groups.GroupText " & _
                "FROM History INNER JOIN Groups ON History.GroupID = Groups.GroupId " & _
                "GROUP BY History.DateViewed, Groups.GroupId, Groups.GroupText " & _
                "ORDER BY History.DateViewed, Groups.GroupText", cn
        
        Do Until rs.EOF
                
            If sDateKey <> "d" & "|" & rs("DateViewed") Then
                sDateKey = "d" & "|" & rs("DateViewed")
                tvFeeds.Nodes.Add "History", tvwChild, sDateKey, Format(rs("DateViewed"), "mm/dd/yyyy"), 5, 5
            End If
            If sKey <> sDateKey & "|" & rs("Groupid") Then
                sKey = sDateKey & "|" & rs("Groupid")
                tvFeeds.Nodes.Add sDateKey, tvwChild, sKey, rs("GroupText"), 5, 5
            End If
            
                    rs2.Open "SELECT History.HistoryID, History.ArticleURL, " & _
                            "History.FeedName, History.Title, History.feedid " & _
                            " From History " & _
                            "WHERE History.GroupID= " & rs("Groupid") & " AND History.DateViewed=#" & rs("DateViewed") & "# ORDER BY DateViewed, Groupid, Feedid", cn, adOpenDynamic, adLockOptimistic
                    Do Until rs2.EOF
                        If sHistKey <> sKey & "|" & rs2("FeedID") Then
                            sHistKey = sKey & "|" & rs2("FeedID")
                            Set oNode = tvFeeds.Nodes.Add(sKey, tvwChild, sHistKey, rs2("FeedName"), 8, 8)
                        End If
                            Set oNode = tvFeeds.Nodes.Add(sHistKey, tvwChild, "HF" & rs2("HistoryID"), rs2("Title"), 8, 8)
                            oNode.Tag = rs2("ArticleURL")
                        rs2.MoveNext
                    Loop
                    
                    rs2.Close
                
            
            rs.MoveNext
        Loop
        
        rs.Close
        
    End If
    
        For i = 1 To mCol.Count
            On Error Resume Next
            tvFeeds.Nodes(mCol.Item(i)).EnsureVisible
        Next
    
End Function

Private Sub imgSplith_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    With imgSplitH
        PicSplitH.Move .Left - 40, .Top, .Width, .Height / 2
    End With
    PicSplitH.Visible = True
    mbMoving = True
End Sub


Private Sub imgSplith_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sglPos As Single
    

    If mbMoving Then
        sglPos = y + imgSplitH.Top
        If sglPos < sglSplitLimit Then
            PicSplitH.Top = sglSplitLimit
        ElseIf sglPos > picRight.ScaleHeight - sglSplitLimit Then
            PicSplitH.Top = picRight.ScaleHeight - sglSplitLimit
        Else
            PicSplitH.Top = sglPos
        End If
    End If
End Sub


Private Sub imgSplith_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    SizeSubControls PicSplitH.Top
    PicSplitH.Visible = False
    mbMoving = False
End Sub



Private Function StartNewBrowser(sURL As String) As Boolean
       
  'start a new instance of the user's browser
  'at the page passed as sURL
   Dim success As Long
   Dim hProcess As Long
   Dim sBrowser As String
   Dim start As STARTUPINFO
   Dim proc As PROCESS_INFORMATION
   Dim sCmdLine As String
   
   sBrowser = GetBrowserName(success)
   
  'did sBrowser get correctly filled?
   If success >= ERROR_FILE_SUCCESS Then
   
      sCmdLine = BuildCommandLine(sBrowser)
      
     'prepare STARTUPINFO members
      With start
         .cb = Len(start)
         .dwFlags = STARTF_USESHOWWINDOW
         .wShowWindow = SW_SHOWNORMAL
      End With
      
     'start a new instance of the default
     'browser at the specified URL. The
     'lpCommandLine member (second parameter)
     'requires a leading space or the call
     'will fail to open the specified page.
      success = CreateProcess(sBrowser, _
                              sCmdLine & sURL, _
                              0&, 0&, 0&, _
                              NORMAL_PRIORITY_CLASS, _
                              0&, 0&, start, proc)
                                  
     'if the process handle is valid, return success
      StartNewBrowser = proc.hProcess <> 0
     
     'don't need the process
     'handle anymore, so close it
      Call CloseHandle(proc.hProcess)

     'and close the handle to the thread created
      Call CloseHandle(proc.hThread)

   End If

End Function


Private Function GetBrowserName(dwFlagReturned As Long) As String

  'find the full path and name of the user's
  'associated browser
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


Private Function BuildCommandLine(ByVal sBrowser As String) As String

  'just in case the returned string is mixed case
   sBrowser = LCase$(sBrowser)
   
  'try for internet explorer
   If InStr(sBrowser, "iexplore.exe") > 0 Then
      BuildCommandLine = " -nohome "
   
  'try for netscape 4.x
   ElseIf InStr(sBrowser, "netscape.exe") > 0 Then
      BuildCommandLine = " "
   
  'try for netscape 7.x
   ElseIf InStr(sBrowser, "netscp.exe") > 0 Then
      BuildCommandLine = " -url "
   
   Else
   
     'not one of the usual browsers, so
     'either determine the appropriate
     'command line required through testing
     'and adding to ElseIf conditions above,
     'or just return a default 'empty'
     'command line consisting of a space
     '(to separate the exe and command line
     'when CreateProcess assembles the string)
      BuildCommandLine = " "
      
   End If
   
End Function


Private Function TrimNull(Item As String)

  'remove string before the terminating null(s)
   Dim pos As Integer
   
   pos = InStr(Item, Chr$(0))
   
   If pos Then
      TrimNull = Left$(Item, pos - 1)
   Else
      TrimNull = Item
   End If
   
End Function


Public Function GetTempDir() As String

  'retrieve the user's system temp folder
   Dim tmp As String
   
   tmp = Space$(MAX_PATH)
   Call GetTempPath(Len(tmp), tmp)
   
   GetTempDir = TrimNull(tmp)
    
End Function

Private Sub tvFeeds_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Not tvFeeds.HitTest(x, y) Is Nothing Then
        Set SelectedNode = tvFeeds.HitTest(x, y)
    End If
    
    If Button = 2 Then
        If bViewHistory Then
            If SelectedNode Is Nothing Then
                mnuHistoryDelete.Visible = False
            Else
                mnuHistoryDelete.Visible = True
            End If
            PopupMenu mnuHistory
        Else
            PopupMenu mnuGroup
        End If
    End If

End Sub

Private Sub tvFeeds_NodeClick(ByVal Node As MSComctlLib.Node)

    Dim rs As New ADODB.Recordset
    Dim rsview As New ADODB.Recordset
    Dim oDom As New FreeThreadedDOMDocument30
    Dim xmlServer As New XMLHTTP40

    Dim oElement As IXMLDOMElement
    Dim oNodelist As IXMLDOMNodeList
    Dim xNode As IXMLDOMNode
    Dim sTemp As String
    Dim mli As MSComctlLib.ListItem
    Dim iFreeFile As Long
    Dim sPicture As String
    Dim sFile As String
    Dim vDate As Variant
    Dim lMonth As Integer
    Dim iCount As Long
    Dim cTotal As Double
    
    If SelectedNode Is Nothing Then Exit Sub
    
    If Node.Tag = "" Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    cTotal = Feed.Count - 1
    
    For iCount = 1 To cTotal
        Unload Feed(iCount)
    Next
    
    picFeeds.Top = VScroll1.min
    VScroll1.Value = VScroll1.min
    
    If Left(Node.Key, 2) = "HF" Then
       ' wb.Visible = True
       ' wbHeader.Visible = True
        If bInBrowser Then
            StartNewBrowser Node.Tag
        Else
            wb.Navigate Node.Tag
        End If

    Else
    
    'lvFeeds.ListItems.Clear
    lvFeeds.Tag = Node.Tag
    rs.Open "SELECT * FROM Feeds " & _
            "WHERE Feeds.FeedID=" & Node.Tag, cn, adOpenDynamic, adLockOptimistic

    If Not (rs.EOF And rs.BOF) Then
        
        sbStatusBar.Panels(1).Text = "Loading Feed " & rs("FeedURL")
        Set oDom = LoadFeed(rs("FeedURL"))
        If Len(oDom.xml) > 0 Then
            sFile = CreateTempFile("RSS")
            oDom.Save sFile
            FileToColumn rs("Feedxml"), sFile
            rs.Update
            Kill sFile
        End If
            If Not oDom.selectSingleNode("//channel/description") Is Nothing Then
                rs("FeedDescription") = oDom.selectSingleNode("//channel/description").Text
                rs.Update
            End If
            
            If IsNull(rs("feedimageurl")) Then
ReloadImage:
                If Not oDom.selectSingleNode("//image") Is Nothing Then
                    If FileExists(App.path & "\images\" & "Feed-" & rs("feedid") & Mid$(oDom.selectSingleNode("//image/url").Text, InStrRev(oDom.selectSingleNode("//image/url").Text, "."))) Then
                        rs("feedImageUrl") = App.path & "\images\" & "Feed-" & rs("feedid") & Mid$(oDom.selectSingleNode("//image/url").Text, InStrRev(oDom.selectSingleNode("//image/url").Text, "."))
                    Else
                        DownloadFile oDom.selectSingleNode("//image/url").Text, App.path & "\images\" & "Feed-" & rs("feedid") & Mid$(oDom.selectSingleNode("//image/url").Text, InStrRev(oDom.selectSingleNode("//image/url").Text, "."))
                        rs("feedImageUrl") = App.path & "\images\" & "Feed-" & rs("feedid") & Mid$(oDom.selectSingleNode("//image/url").Text, InStrRev(oDom.selectSingleNode("//image/url").Text, "."))
                    End If
                    rs.Update
                End If
            ElseIf Not FileExists(rs("feedimageurl")) Then
                GoTo ReloadImage
            End If
            Dim perChange As Double
            If FileExists(rs("feedimageurl")) Then
                imgFeed.Stretch = False
                Set imgFeed.Picture = LoadPicture(rs("Feedimageurl"))
                If imgFeed.Height > picFeedHeader.ScaleHeight Then
                    imgFeed.Stretch = True
                    perChange = (imgFeed.Height - picFeedHeader.ScaleHeight) / imgFeed.Height
                    imgFeed.Height = imgFeed.Height * (1 - perChange)
                    imgFeed.Width = imgFeed.Width * (1 - perChange)
                End If
                lblDescription.Move imgFeed.Width + 50, 0, picFeedHeader.ScaleWidth - (100 + imgFeed.Width), picFeedHeader.ScaleHeight
            Else
                lblDescription.Move 0, 0, picFeedHeader.ScaleWidth, picFeedHeader.ScaleHeight
            End If
            
            lblDescription.Text = rs("Feeddescription")
            
            Set oNodelist = oDom.selectNodes("//item")
            For Each xNode In oNodelist
                If Not xNode.selectSingleNode("title") Is Nothing Then
                    iCount = Feed.Count
                    Load Feed(iCount)
                    Feed(iCount).Top = (Feed.Count - 2) * Feed(iCount).Height
                    Feed(iCount).Left = 0
                    Feed(iCount).Width = picFeeds.ScaleWidth
                    'Set mli = lvFeeds.ListItems.Add()
                    Feed(iCount).URL = xNode.selectSingleNode("link").Text
                    If Not xNode.selectSingleNode("pubDate") Is Nothing Then
                        vDate = Split(Mid$(xNode.selectSingleNode("pubDate").Text, 6), " ")
                        
                        Select Case vDate(1)
                            Case "Jan": lMonth = 1
                            Case "Feb": lMonth = 2
                            Case "Mar": lMonth = 3
                            Case "Apr": lMonth = 4
                            Case "May": lMonth = 5
                            Case "Jun": lMonth = 6
                            Case "Jul": lMonth = 7
                            Case "Aug": lMonth = 8
                            Case "Sep": lMonth = 9
                            Case "Oct": lMonth = 10
                            Case "Nov": lMonth = 11
                            Case "Dec": lMonth = 12
                        End Select
                        Feed(iCount).FeedDate = Format(vDate(0) & "/" & lMonth & "/" & vDate(2) & " " & vDate(3), "mm/dd/yyyy hh:mm") & " " & vDate(4)
                        'mli.Text = xNode.selectSingleNode("pubDate").Text
                    Else
                        Feed(iCount).FeedDate = "Unknown"
                    End If
                    
                    Feed(iCount).Caption = xNode.selectSingleNode("title").Text
                    Feed(iCount).Description = xNode.selectSingleNode("description").Text
                    rsview.Open "Select * from History where ArticleURL = '" & Feed(iCount).URL & "'", cn
                    If Not (rsview.EOF And rsview.BOF) Then
                        Feed(iCount).Read = True
                    Else
                        Feed(iCount).Read = False
                    End If

                    rsview.Close
                    Feed(iCount).Visible = True
                End If
            Next
        'End If
        lvFeeds_Resize
        
        sbStatusBar.Panels(1).Text = "Done..."
    
    End If
    
    rs.Close
    
    End If
    
    If lvFeeds.Visible Then lvFeeds.SetFocus

    Screen.MousePointer = vbDefault
    
End Sub

Private Sub tvFeeds_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

        Dim rss As String
    Dim frm As frmAddContent
    
    rss = Data.Getdata(1)
    
    Set frm = New frmAddContent
    
    Load frm
    
    frm.txtURL.Text = rss
    frm.Getdata
    frm.Show vbModal
    
    If Not frm.Cancel Then
        LoadSources
    End If
    
    Unload frm
    
    Set frm = Nothing
    


End Sub

Private Sub VScroll1_Change()

    picFeeds.Top = VScroll1.Value

End Sub

Private Sub wb_StatusTextChange(ByVal Text As String)

    sbStatusBar.Panels(1).Text = Text

End Sub

Private Sub wbHeader_Resize()
    On Error Resume Next
    
    Image1.Move (wbHeader.ScaleWidth / 2) - (Image1.Width / 2), (wbHeader.ScaleHeight / 2) - (Image1.Height / 2)

End Sub

Private Sub RefreshFeeds()

    Dim oDom As New FreeThreadedDOMDocument30
    Dim rs As New ADODB.Recordset
    Dim sFile As String
    
    rs.Open "Select * from feeds where subscribed = -1", cn, adOpenDynamic, adLockOptimistic
    
    Do Until rs.EOF
        sbStatusBar.Panels(1).Text = "Now retrieving " & rs("feedname")
        
        Set oDom = LoadFeed(rs("feedurl"))
        
        If Len(oDom.xml) > 0 Then
            sFile = CreateTempFile("rss")
            oDom.Save sFile
            
            If FileExists(sFile) Then
                FileToColumn rs("FeedXML"), sFile
                rs.Update
            End If
            
            Kill sFile
            
        End If
        
        rs.MoveNext
    Loop
    
    rs.Close
    
    Set rs = Nothing

End Sub
