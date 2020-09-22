VERSION 5.00
Begin VB.Form frmAddContent 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Content"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   7650
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox lblDescription 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1140
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Text            =   "frmAddContent.frx":0000
      Top             =   1860
      Width           =   6375
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6480
      TabIndex        =   12
      Top             =   3660
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5460
      TabIndex        =   11
      Top             =   3660
      Width           =   975
   End
   Begin VB.ComboBox cboGroup 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1140
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   900
      Width           =   5475
   End
   Begin VB.CheckBox chkShow 
      Caption         =   "Show in Favorites"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1140
      TabIndex        =   3
      Top             =   600
      Width           =   1875
   End
   Begin VB.CommandButton cmdGet 
      Caption         =   "Get Data"
      Height          =   315
      Left            =   6660
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtURL 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1140
      TabIndex        =   1
      Top             =   120
      Width           =   5475
   End
   Begin VB.Label lblImage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Left            =   1140
      TabIndex        =   10
      Top             =   2580
      Width           =   60
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   2940
      Top             =   2940
      Width           =   1695
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Left            =   1140
      TabIndex        =   9
      Top             =   1440
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Image"
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
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   2580
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   1860
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
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
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List Under"
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
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RSS URL"
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
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   795
   End
End
Attribute VB_Name = "frmAddContent"
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

Private Sub cboGroup_Click()

    If cboGroup.ListIndex = -1 Then Exit Sub
    
    cmdOK.Enabled = True

End Sub

Private Sub cmdCancel_Click()
    bCancel = True
    Me.Hide

End Sub

Private Sub cmdGet_Click()
    Getdata
    

End Sub

Private Sub cmdOK_Click()

    Dim rs As New ADODB.Recordset
    
    If Me.Tag = "" Then
        rs.Open "Select * from Feeds", cn, adOpenDynamic, adLockOptimistic
        rs.AddNew
    Else
        rs.Open "Select * from Feeds WHERE Feedid =" & Me.Tag, cn, adOpenDynamic, adLockOptimistic
    End If
    
    rs("GroupID") = cboGroup.ItemData(cboGroup.ListIndex)
    If chkShow.Value = vbChecked Then
        rs("CustomId") = 17
    Else
        rs("customid") = vbNull
    End If
    
    rs("FeedName") = lblName.Caption
    rs("FeedDescription") = lblDescription.Text
    rs("FeedURL") = txtURL.Text
    rs("FeedImageURL") = lblImage.Caption
    rs("Subscribed") = -1
    rs.Update
    rs.Close

    bCancel = False
    Me.Hide

End Sub

Private Sub Form_Load()

    Dim rs As New ADODB.Recordset
    
    bCancel = True
    
    rs.Open "select * from groups", cn
    
    cboGroup.Clear
    
    Do Until rs.EOF
        cboGroup.AddItem rs("Grouptext")
        cboGroup.ItemData(cboGroup.NewIndex) = rs("Groupid")
        rs.MoveNext
    Loop

    rs.Close
    
    Set rs = Nothing
    
End Sub

Public Sub LoadFeed(feedid As Long)

    Dim rs As New ADODB.Recordset
    Dim iCount As Long
    
    rs.Open "SELECT Feeds.FeedID, Feeds.GroupID, Feeds.CustomId, " & _
            "Feeds.FeedName, Feeds.FeedDescription, Feeds.FeedURL, " & _
            "Feeds.FeedImageURL, Feeds.CheckInterval, " & _
            "Feeds.Subscribed, Feeds.LastCheck " & _
            "From Feeds WHERE Feeds.FeedID=" & feedid, cn
    
    If Not (rs.EOF And rs.BOF) Then
        Me.Tag = feedid
        For iCount = 0 To cboGroup.ListCount - 1
            If cboGroup.ItemData(iCount) = rs("GroupID") Then
                cboGroup.ListIndex = iCount
                Exit For
            End If
        Next
        If IsNull(rs("CustomId")) Then
            chkShow.Value = vbUnchecked
        Else
            chkShow.Value = vbChecked
        End If
        
        lblName.Caption = rs("FeedName")
        lblDescription.Text = rs("FeedDescription")
        txtURL.Text = rs("FeedURL")
        If IsNull(rs("Feedimageurl")) Then
            lblImage.Caption = ""
        Else
            lblImage.Caption = rs("FeedImageURL")
        End If
    End If
    
    rs.Close
    
    Set rs = Nothing

End Sub


Public Sub Getdata()
    Dim oDom As New FreeThreadedDOMDocument30
    
    Set oDom = modMain.LoadFeed(txtURL.Text)

    If Not oDom.selectSingleNode("//channel/title") Is Nothing Then
        lblName.Caption = oDom.selectSingleNode("//channel/title").Text

    End If


    If Not oDom.selectSingleNode("//channel/description") Is Nothing Then
        lblDescription.Text = oDom.selectSingleNode("//channel/description").Text
    End If
    
    If Not oDom.selectSingleNode("//image") Is Nothing Then
        lblImage.Caption = oDom.selectSingleNode("//image/url").Text
    
        If FileExists(App.path & "\images\" & "Temp" & Mid$(oDom.selectSingleNode("//image/url").Text, InStrRev(oDom.selectSingleNode("//image/url").Text, "."))) Then
            Kill App.path & "\images\" & "Temp" & Mid$(oDom.selectSingleNode("//image/url").Text, InStrRev(oDom.selectSingleNode("//image/url").Text, "."))
        End If
        
        DownloadFile oDom.selectSingleNode("//image/url").Text, App.path & "\images\" & "Temp" & Mid$(oDom.selectSingleNode("//image/url").Text, InStrRev(oDom.selectSingleNode("//image/url").Text, "."))
    
        Set Image1.Picture = LoadPicture(App.path & "\images\" & "Temp" & Mid$(oDom.selectSingleNode("//image/url").Text, InStrRev(oDom.selectSingleNode("//image/url").Text, ".")))
    End If
End Sub
