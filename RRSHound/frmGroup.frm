VERSION 5.00
Begin VB.Form frmGroup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Groups Edit"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtGroupName 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   300
      Width           =   4515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group Name"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   900
   End
End
Attribute VB_Name = "frmGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

