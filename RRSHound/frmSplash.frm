VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8970
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0442
   ScaleHeight     =   5985
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   240
      Top             =   1500
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4860
      TabIndex        =   0
      Top             =   1320
      Width           =   720
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'API Delcares
    Private Declare Function SetWindowPos _
               Lib "user32" (ByVal hWnd As Long, _
                             ByVal hWndInsertAfter As Long, _
                             ByVal x As Long, _
                             ByVal y As Long, _
                             ByVal cx As Long, _
                             ByVal cy As Long, _
                             ByVal wFlags As Long) As Long

Dim bytRegion(5551) As Byte
Dim nBytes As Long

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    Call SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, &H2 Or &H1)
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & " Build " & App.Revision
    
End Sub

Private Sub Timer1_Timer()
    Unload Me
End Sub
