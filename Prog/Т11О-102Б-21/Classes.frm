VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form1"
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12735
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   12735
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox list1 
      BackColor       =   &H00FFC0FF&
      Columns         =   3
      BeginProperty Font 
         Name            =   "Minion Pro"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      ItemData        =   "Classes.frx":0000
      Left            =   4800
      List            =   "Classes.frx":000D
      TabIndex        =   2
      Top             =   1080
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Minion Pro"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1560
      TabIndex        =   1
      Text            =   "AAAAAAAAAAAAA"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select"
      BeginProperty Font 
         Name            =   "Minion Pro"
         Size            =   20.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   0
      Top             =   3120
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim j As New Progs
j.start list1.ListIndex + 1


End Sub


