VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl FeedList 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   KeyPreview      =   -1  'True
   ScaleHeight     =   870
   ScaleWidth      =   7500
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1260
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FeedList.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FeedList.ctx":0452
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgClick 
      Height          =   195
      Left            =   180
      Top             =   780
      Width           =   195
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   60
      Picture         =   "FeedList.ctx":08A4
      Stretch         =   -1  'True
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   195
      Left            =   7080
      TabIndex        =   2
      Top             =   60
      Width           =   345
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   660
      TabIndex        =   1
      Top             =   60
      Width           =   6075
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   $"FeedList.ctx":0CE6
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   780
      TabIndex        =   0
      Top             =   300
      Width           =   6180
   End
End
Attribute VB_Name = "FeedList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const cSelected = &HFFC0C0
Const UnSelected = &HFFFFFF
'Default Property Values:
Const m_def_Read = 0
Const m_def_Caption = ""
Const m_def_Description = ""
Const m_def_Date = ""
Const m_def_URL = ""
'Property Variables:
Dim m_Read As Boolean
Dim m_Caption As String
Dim m_Description As String
Dim m_Date As String
Dim m_URL As String
'Event Declarations:
Event ItemClick(URL As String) 'MappingInfo=UserControl,UserControl,-1,Click
Dim m_Selected As Boolean

Public Property Get Selected() As Boolean
    Selected = m_Selected
End Property

Public Property Let Selected(vData As Boolean)
    m_Selected = vData
    If m_Selected Then
        UserControl.BackColor = &HFFFFC0
    Else
        UserControl.BackColor = &HFFFFFF
    End If
End Property
Private Sub imgClick_Click()
    RaiseEvent ItemClick(m_URL)
End Sub

Private Sub UserControl_Initialize()
    lblCaption.Caption = ""
    lblDescription.Caption = ""
    lblDate.Caption = ""
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function Add(Caption As String, Description As String, FeedDate As String, URL As String) As Variant

End Function

Private Sub UserControl_Click()
    RaiseEvent ItemClick(m_URL)
End Sub

Public Property Get URL() As String
    URL = m_URL
End Property

Public Property Let URL(ByVal New_URL As String)
    m_URL = New_URL
    PropertyChanged "URL"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    lblCaption.Caption = New_Caption
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Description() As String
    Description = m_Description
End Property

Public Property Let Description(ByVal New_Description As String)
    m_Description = New_Description
    lblDescription.Caption = New_Description
    PropertyChanged "Description"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get FeedDate() As String
    FeedDate = m_Date
End Property

Public Property Let FeedDate(ByVal New_Date As String)
    m_Date = New_Date
    lblDate.Caption = New_Date
    PropertyChanged "Date"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Caption = m_def_Caption
    m_Description = m_def_Description
    m_Date = m_def_Date
    m_URL = m_def_URL
    m_Read = m_def_Read
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    m_Description = PropBag.ReadProperty("Description", m_def_Description)
    m_Date = PropBag.ReadProperty("Date", m_def_Date)
    m_URL = PropBag.ReadProperty("URL", m_def_URL)
    m_Read = PropBag.ReadProperty("Read", m_def_Read)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    lblDate.Left = UserControl.ScaleWidth - lblDate.Width
    lblCaption.Width = UserControl.ScaleWidth - (lblDate.Width + lblCaption.Left)
    lblDescription.Width = UserControl.ScaleWidth - lblDescription.Left
    UserControl.Height = lblDescription.Top + lblDescription.Height
    imgClick.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("Description", m_Description, m_def_Description)
    Call PropBag.WriteProperty("Date", m_Date, m_def_Date)
    Call PropBag.WriteProperty("URL", m_URL, m_def_URL)
    
    Call PropBag.WriteProperty("Read", m_Read, m_def_Read)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Read() As Boolean
    Read = m_Read
End Property

Public Property Let Read(ByVal New_Read As Boolean)
    m_Read = New_Read
    
    If m_Read Then
        Set Image1.Picture = ImageList1.ListImages(1).Picture
    Else
        Set Image1.Picture = ImageList1.ListImages(2).Picture
    End If
    
    PropertyChanged "Read"
End Property

