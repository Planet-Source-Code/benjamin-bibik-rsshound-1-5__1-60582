VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmhistory 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "History"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Delete Checked"
      Height          =   435
      Left            =   60
      TabIndex        =   3
      Top             =   4740
      Width           =   1455
   End
   Begin VB.TextBox txtDescription 
      Height          =   3855
      Left            =   4920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   720
      Width           =   1995
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   435
      Left            =   5700
      TabIndex        =   1
      Top             =   4740
      Width           =   1155
   End
   Begin MSComctlLib.ListView lvHistory 
      Height          =   3855
      Left            =   60
      TabIndex        =   0
      Top             =   720
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Viewed Feed"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   6180
      Picture         =   "frmhistory.frx":0000
      Top             =   60
      Width           =   720
   End
End
Attribute VB_Name = "frmhistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()

    Me.Hide

End Sub

Private Sub Form_Load()

    Dim rs As New ADODB.Recordset
    Dim mli As MSComctlLib.ListItem
    
    
    
    rs.Open "SELECT ViewedFeeds.FeedURL, ViewedFeeds.FeedName, ViewedFeeds.FeedDescription, ViewedFeeds.ViewDate " & _
            "FROM ViewedFeeds ORDER BY ViewedFeeds.ViewDate", cn
            
    Do Until rs.EOF
        Set mli = lvHistory.ListItems.Add()
        mli.Text = rs("FeedName")
        mli.Tag = rs("FeedURL")
        mli.ToolTipText = rs("FeedDescription")
        rs.MoveNext
    Loop
    
    rs.Close

    lvAutosizeMax lvHistory

End Sub

Private Sub lvHistory_DblClick()

    If Not lvHistory.SelectedItem Is Nothing Then
        Me.Hide
        fMainForm.wb.Navigate lvHistory.SelectedItem.Tag

    End If

End Sub

Private Sub lvHistory_ItemClick(ByVal Item As MSComctlLib.ListItem)

    txtDescription.Text = Item.ToolTipText

End Sub
