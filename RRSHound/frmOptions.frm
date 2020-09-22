VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Options"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   6150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Options"
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2490
      TabIndex        =   1
      Tag             =   "OK"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Tag             =   "Cancel"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Tag             =   "&Apply"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   2022
         Left            =   505
         TabIndex        =   11
         Tag             =   "Sample 4"
         Top             =   502
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   2022
         Left            =   406
         TabIndex        =   10
         Tag             =   "Sample 3"
         Top             =   403
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   2022
         Left            =   307
         TabIndex        =   8
         Tag             =   "Sample 2"
         Top             =   305
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   0
      Left            =   210
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample1 
         Height          =   3705
         Left            =   0
         TabIndex        =   4
         Tag             =   "Sample 1"
         Top             =   0
         Width           =   5640
         Begin VB.CheckBox chkHistory 
            Caption         =   "History"
            Height          =   195
            Left            =   240
            TabIndex        =   14
            Top             =   660
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.Frame fmHistory 
            Caption         =   "History"
            Height          =   1155
            Left            =   120
            TabIndex        =   13
            Top             =   660
            Width           =   5415
            Begin MSComCtl2.UpDown UpDown1 
               Height          =   285
               Left            =   2940
               TabIndex        =   19
               Top             =   720
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   393216
               Value           =   1
               AutoBuddy       =   -1  'True
               BuddyControl    =   "txtHistory"
               BuddyDispid     =   196619
               OrigLeft        =   3000
               OrigTop         =   720
               OrigRight       =   3255
               OrigBottom      =   975
               Max             =   31
               Min             =   1
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.TextBox txtHistory 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   18
               Text            =   "20"
               Top             =   720
               Width           =   420
            End
            Begin VB.CommandButton cmdHistory 
               Caption         =   "Delete History"
               Height          =   315
               Left            =   3720
               TabIndex        =   17
               Top             =   720
               Width           =   1515
            End
            Begin VB.Image Image1 
               Height          =   480
               Left            =   120
               Picture         =   "frmOptions.frx":0000
               Top             =   360
               Width           =   480
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Days to keep in histroy"
               Height          =   195
               Left            =   720
               TabIndex        =   16
               Top             =   780
               Width           =   1605
            End
            Begin VB.Label Label1 
               Caption         =   "The history keeps track of feeds that you have viewed.  This information is used to search previously viewed news feeds."
               Height          =   495
               Left            =   720
               TabIndex        =   15
               Top             =   240
               Width           =   4635
            End
         End
         Begin VB.CheckBox chkBrowser 
            Caption         =   "Open Feeds in default browser"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   5235
         End
      End
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   4245
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7488
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Group 2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Group 3"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Group 4"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bDirty As Boolean

Public Property Let Dirty(vData As Boolean)
    bDirty = vData
    cmdApply.Enabled = vData
End Property

Public Property Get Dirty() As Boolean
    Dirty = bDirty
End Property

Private Sub chkBrowser_Click()
    Dirty = True
End Sub

Private Sub chkHistory_Click()

    fmHistory.Enabled = chkHistory.Value = vbChecked
    Dirty = True
End Sub

Private Sub cmdApply_Click()
    'ToDo: Add 'cmdApply_Click' code.
    SaveDefaults
    
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdHistory_Click()

    If MsgBox("Are you sure you want to delete all history?", vbQuestion + vbYesNo, "Delete?") = vbYes Then
        cn.Execute "Delete from ViewedFeeds"
    End If

End Sub

Private Sub cmdOK_Click()
    'ToDo: Add 'cmdOK_Click' code.
    SaveDefaults
    Unload Me
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    i = tbsOptions.SelectedItem.Index
    'handle ctrl+tab to move to the next tab
    If (Shift And 3) = 2 And KeyCode = vbKeyTab Then
        If i = tbsOptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    ElseIf (Shift And 3) = 3 And KeyCode = vbKeyTab Then
        If i = 1 Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(tbsOptions.Tabs.Count)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i - 1)
        End If
    End If
End Sub


Private Sub Form_Load()
    cmdApply.Enabled = False
    LoadDefaults
End Sub

Private Sub tbsOptions_Click()
    

    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tbsOptions.SelectedItem.Index - 1 Then
            picOptions(i).Left = 210
            picOptions(i).Enabled = True
        Else
            picOptions(i).Left = -20000
            picOptions(i).Enabled = False
        End If
    Next
    

End Sub


Private Sub LoadDefaults()

    chkBrowser.Value = GetOption("openinbrowser", 0)
    chkHistory.Value = GetOption("keephistory", 1)
    txtHistory.Text = GetOption("daysinhistory", "20")
    
End Sub

Private Sub SaveDefaults()

    WriteOption "openinbrowser", chkBrowser.Value
    WriteOption "keephistory", chkHistory.Value
    WriteOption "daysinhistory", txtHistory.Text
    
    Dirty = False
    
End Sub

Private Sub txtHistory_Change()
    Dirty = True
End Sub
