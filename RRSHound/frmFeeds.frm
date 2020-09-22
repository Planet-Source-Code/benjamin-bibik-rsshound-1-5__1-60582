VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFeeds 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Feed Properties"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Selected"
      Enabled         =   0   'False
      Height          =   375
      Left            =   60
      TabIndex        =   3
      Top             =   2640
      Width           =   1395
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   2640
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6420
      TabIndex        =   1
      Top             =   2640
      Width           =   915
   End
   Begin MSComctlLib.ListView lvFeeds 
      Height          =   2175
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   3836
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Feed Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Feed URL"
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Checked items are currently subscribed.  Double click to edit."
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   60
      Width           =   7155
   End
End
Attribute VB_Name = "frmFeeds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bCancel As Boolean

Public Property Let Cancel(vdata As Boolean)
    bCancel = vdata
End Property

Public Property Get Cancel() As Boolean
    Cancel = bCancel
End Property

Private Sub cmdCancel_Click()
    bCancel = True
    Me.Hide
End Sub

Private Sub cmdDelete_Click()
    
    If MsgBox("Are you sure you want to delete:" & vbCrLf & vbCrLf & lvFeeds.SelectedItem.Text, vbQuestion + vbYesNo, "Delete Selected?") = vbYes Then
        cn.Execute "Delete from feeds where feedid = " & lvFeeds.SelectedItem.Tag
        LoadFeeds
        cmdDelete.Enabled = False
    End If
    
End Sub

Private Sub cmdOK_Click()
    bCancel = False
    Me.Hide
End Sub

Private Sub Form_Load()
    bCancel = True
    LoadFeeds

End Sub

Public Sub LoadFeeds()

    Dim mli As MSComctlLib.ListItem
    Dim rs As New ADODB.Recordset
    
    lvFeeds.ListItems.Clear
    
    rs.Open "Select Feedid, feedname, feedurl, subscribed from feeds", cn
    
    Do Until rs.EOF
        Set mli = lvFeeds.ListItems.Add()
        mli.Tag = rs("feedid")
        mli.Text = rs("feedName")
        mli.SubItems(1) = rs("feedurl")
        mli.Checked = rs("subscribed")
        rs.MoveNext
    Loop
    
    rs.Close
    
End Sub

Private Sub lvFeeds_DblClick()

    If lvFeeds.SelectedItem Is Nothing Then Exit Sub
    
    Dim frm As frmAddContent
    
    Set frm = New frmAddContent
    
    Load frm
    
    frm.LoadFeed lvFeeds.SelectedItem.Tag
    
    frm.Show vbModal
    
    If Not frm.Cancel Then
        LoadFeeds
    End If
    
    Unload frm
    
    Set frm = Nothing

End Sub

Private Sub lvFeeds_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    cn.Execute "UPDATE FEEDS set Subscribed = " & Item.Checked & " where feedid = " & Item.Tag

End Sub

Private Sub lvFeeds_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cmdDelete.Enabled = True
End Sub
