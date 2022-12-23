VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF00FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OCTAHEDRON"
   ClientHeight    =   12000
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   19275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "okt.frx":0000
   ScaleHeight     =   12000
   ScaleWidth      =   19275
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Roboto"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   11535
      TabIndex        =   27
      Text            =   "20"
      Top             =   8175
      Width           =   1365
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Roboto"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   14895
      TabIndex        =   26
      Text            =   "50"
      Top             =   8205
      Width           =   1395
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Roboto"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   13335
      TabIndex        =   25
      Text            =   "30"
      Top             =   8190
      Width           =   1320
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H00FFC0FF&
      Caption         =   ".rtf"
      BeginProperty Font 
         Name            =   "Lexend"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      Picture         =   "okt.frx":2DDC
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton Command18 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "Photoshop"
      BeginProperty Font 
         Name            =   "Lexend"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1320
      Picture         =   "okt.frx":7026
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00000080&
      Caption         =   "SVG"
      BeginProperty Font 
         Name            =   "Lexend"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00000080&
      Caption         =   "BMP"
      BeginProperty Font 
         Name            =   "Lexend"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Command15 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Hide"
      BeginProperty Font 
         Name            =   "Lexend"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17385
      MaskColor       =   &H00FF00FF&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   690
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command14 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "Draft"
      BeginProperty Font 
         Name            =   "Lexend"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2535
      Picture         =   "okt.frx":B270
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8760
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF80FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   6015
      Left            =   10665
      Picture         =   "okt.frx":F4BA
      ScaleHeight     =   8597.148
      ScaleMode       =   0  'User
      ScaleWidth      =   8674.804
      TabIndex        =   17
      Top             =   660
      Visible         =   0   'False
      Width           =   7875
   End
   Begin VB.CommandButton Command13 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "Ruby"
      BeginProperty Font 
         Name            =   "Lexend"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4920
      Picture         =   "okt.frx":14465
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   ".zip"
      BeginProperty Font 
         Name            =   "Lexend"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7305
      MaskColor       =   &H0000C000&
      Picture         =   "okt.frx":186AF
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8760
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2175
      Left            =   5745
      TabIndex        =   14
      Top             =   5400
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   3836
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      BackColor       =   64
      ForeColor       =   16777215
      BackColorFixed  =   12632319
      ForeColorFixed  =   16777215
      BackColorSel    =   -2147483638
      ForeColorSel    =   16777215
      BackColorBkg    =   255
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "side | volume"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lexend"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command11 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Add new value"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Lexend"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      MaskColor       =   &H00C000C0&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton Command10 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "Printer (PDF)"
      BeginProperty Font 
         Name            =   "Lexend"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3720
      Picture         =   "okt.frx":1C8F9
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "Copy"
      BeginProperty Font 
         Name            =   "Lexend"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      Picture         =   "okt.frx":20B43
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "Volume"
      BeginProperty Font 
         Name            =   "Lexend"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6135
      OLEDropMode     =   1  'Manual
      Picture         =   "okt.frx":24D8D
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8760
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   14040
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command7 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "Save .txt"
      BeginProperty Font 
         Name            =   "Lexend"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8505
      MaskColor       =   &H00FF00FF&
      Picture         =   "okt.frx":28FD7
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "Open .txt"
      BeginProperty Font 
         Name            =   "Lexend"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7320
      Picture         =   "okt.frx":2D221
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "SOLIDWORKS"
      BeginProperty Font 
         Name            =   "Lexend"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6120
      Picture         =   "okt.frx":3146B
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "IExplorer"
      BeginProperty Font 
         Name            =   "Lexend"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4920
      Picture         =   "okt.frx":356B5
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "Powerpoint"
      BeginProperty Font 
         Name            =   "Lexend"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3720
      Picture         =   "okt.frx":398FF
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7545
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "Excel"
      DisabledPicture =   "okt.frx":3DB49
      DownPicture     =   "okt.frx":4659A
      BeginProperty Font 
         Name            =   "Lexend"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   2505
      Picture         =   "okt.frx":4EFEB
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7560
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "Word"
      BeginProperty Font 
         Name            =   "Lexend"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1260
      Picture         =   "okt.frx":53235
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000040&
      Height          =   2175
      Left            =   90
      TabIndex        =   0
      Top             =   5385
      Width           =   5655
      Begin VB.TextBox Text4 
         BackColor       =   &H00000080&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Lexend"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   1935
         Left            =   2205
         TabIndex        =   2
         Text            =   "20"
         Top             =   150
         Width           =   3375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF80FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Cuboid side"
         BeginProperty Font 
            Name            =   "Lexend"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1245
         Left            =   195
         TabIndex        =   1
         Top             =   825
         Width           =   2055
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Printer Mode"
      BeginProperty Font 
         Name            =   "Lexend"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   7800
      TabIndex        =   22
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Menu mmSort 
      Caption         =   "Sort"
      Begin VB.Menu mSortA 
         Caption         =   "SortA"
      End
      Begin VB.Menu mSortV 
         Caption         =   "SortV"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim print_mode As Integer
Dim draw_mode As Boolean

Sub draw() ' Drafting

If draw_mode = True Then
    Me.Picture1.BackColor = &HFFFFFF
    Me.Picture1.ForeColor = &H0&
Else
    Me.Picture1.BackColor = &HFF80FF
    Me.Picture1.ForeColor = &HFFFFFF
End If

    
    
    
Me.Picture1.Visible = True
Me.Command15.Visible = True
Dim matr(7, 2) As Double
Dim batr
Dim batr_dash

matr(0, 0) = 8.001
matr(0, 1) = 4.028

matr(1, 0) = 0
matr(1, 1) = 4.028

matr(2, 0) = 1.732
matr(2, 1) = 5.087

matr(3, 0) = 9.732
matr(3, 1) = 5.087

matr(4, 0) = 8.001
matr(4, 1) = 8.028

matr(5, 0) = 9.732
matr(5, 1) = 9.028

matr(6, 0) = 0
matr(6, 1) = 8.028

matr(7, 0) = 1.732
matr(7, 1) = 9.028



batr = Array(0, 1, 1, 2, 2, 3, 3, 0, 1, 6, 6, 7, 2, 7, 7, 5, 3, 5)
batr_dash = Array(0, 4, 6, 4, 5, 4)
Me.Picture1.Cls
Me.Picture1.Scale (-20, -15)-(80, 65)
'Me.Picture1.draw
Me.Picture1.DrawStyle = 0

For k = 0 To UBound(batr) Step 2
    mn = Int(batr(k))
    bn = Int(batr(k + 1))
    ff = matr()
'    Me.Picture1
    Me.Picture1.Line (matr(mn, 0) * 5, matr(mn, 1) * 5)-(matr(bn, 0) * 5, matr(bn, 1) * 5)
Next

Me.Picture1.DrawStyle = 2
For l = 0 To UBound(batr_dash) Step 2
    mn = Int(batr_dash(l))
    bn = Int(batr_dash(l + 1))
    ff = matr()
    
    Me.Picture1.Line (matr(mn, 0) * 5, matr(mn, 1) * 5)-(matr(bn, 0) * 5, matr(bn, 1) * 5)
Next
Me.Picture1.DrawStyle = 0


vinos_x = matr(2, 0) * 5 + ((matr(1, 0) - matr(2, 0)) * 5) / 2
vinos_y = matr(2, 1) * 5 + ((matr(1, 1) - matr(2, 1)) * 5) / 2

Me.Picture1.Line (vinos_x, vinos_y)-(vinos_x - 7, vinos_y - 7)
Me.Picture1.Line (vinos_x - 7, vinos_y - 7)-(vinos_x - 16, vinos_y - 7)
Me.Picture1.CurrentX = vinos_x - 15
Me.Picture1.CurrentY = vinos_y - 12
Me.Picture1.FontSize = 13
Me.Picture1.Print Format(CDbl(Me.Text1.Text), "0.000")

Me.Picture1.CurrentX = 4
Me.Picture1.CurrentY = 22
a = (Me.Text1 ^ 3 * Sqr(2)) / 3
Me.Picture1.Print "Объем параллелепипеда:" & Format(CDbl(Me.Text1.Text), "0.000") & vbCrLf & "   равен " & Format(a, "0.000")


End Sub
'MS Word
Private Sub Command1_Click()
Dim w As Object

On Error GoTo noword
Set w = CreateObject("word.application")
 On Error GoTo 0
 w.Visible = True
 w.Documents.Add
 w.selection.typetext "Объём заданной фигуры: " & Me.Text1 * Me.Text2 * Me.Text3 & " m"
 w.Activate
 Set w = Nothing
 Exit Sub
noword:
    MsgBox "noword"
End Sub

Private Sub Command11_Click() 'заполнение таблица объемами
Me.MSFlexGrid1.Rows = Me.MSFlexGrid1.Rows + 1
Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Rows - 1, 0) = Me.Text1.Text
Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Rows - 1, 1) = Format((Me.Text1 ^ 3 * Sqr(2)) / 3, "0.000")
End Sub

Private Sub Command12_Click()
Call arhiv
End Sub
'Vector
Sub prr()

Dim matr(5, 2) As Double
Dim batr

matr(0, 0) = 8.001
matr(0, 1) = 4.028
matr(1, 0) = 0
matr(1, 1) = 4.028
matr(2, 0) = 1.732
matr(2, 1) = 5.087
matr(3, 0) = 9.732
matr(3, 1) = 5.087
matr(4, 0) = 3.366
matr(4, 1) = 8.586
matr(5, 0) = 3.366
matr(5, 1) = 0.53
batr = Array(0, 1, 1, 2, 2, 3, 3, 0, 1, 4, 2, 4, 0, 4, 3, 4, 1, 5, 2, 5, 0, 5, 3, 5)



Printer.ScaleMode = 0
Printer.DrawWidth = 10



Printer.Scale (-15, -20)-(45, 60)
For k = 0 To UBound(batr) Step 2
    mn = Int(batr(k))
    bn = Int(batr(k + 1))
    ff = matr()
    Printer.DrawStyle = vbDash
    Printer.Line (matr(mn, 0) * 5, matr(mn, 1) * 5)-(matr(bn, 0) * 5, matr(bn, 1) * 5)
Next

vinos_x = matr(2, 0) * 5 + ((matr(1, 0) - matr(2, 0)) * 5) / 2
vinos_y = matr(2, 1) * 5 + ((matr(1, 1) - matr(2, 1)) * 5) / 2

Printer.Line (vinos_x, vinos_y)-(vinos_x - 7, vinos_y - 7)
Printer.Line (vinos_x - 7, vinos_y - 7)-(vinos_x - 16, vinos_y - 7)
Printer.CurrentX = vinos_x - 15
Printer.CurrentY = vinos_y - 10
Printer.FontSize = 13
Printer.Print Format(CDbl(Me.Text1.Text), "0.000")

Printer.CurrentX = -10
Printer.CurrentY = -10
a = (Me.Text1 ^ 3 * Sqr(2)) / 3
Printer.Print "Объем параллелепипеда" & Format(CDbl(Me.Text1.Text), "0.000") & vbCrLf & "   равен " & Format(a, "0.000")
Printer.EndDoc


End Sub
'Arch
Sub arhiv()
Me.CommonDialog1.FileName = ""
Me.CommonDialog1.ShowSave

If Me.CommonDialog1.FileName <> "" Then
    Dim ShellApp As Object
    Open Me.CommonDialog1.FileName For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
    
    Set ShellApp = CreateObject("shell.application")
    filetozip = "C:\1\фыр.txt" 'путь к файлу
    ShellApp.namespace(Me.CommonDialog1.FileName).copyhere filetozip
End If

End Sub
'Ruby compiler
Private Sub Command13_Click()
Shell "C:\Program Files\Google\Chrome\Application\chrome.exe https://onlinegdb.com/aCyYDN7d4/"
End Sub

Private Sub Command14_Click()
Call draw
End Sub

Private Sub Command15_Click()
Me.Picture1.Visible = False
Me.Command15.Visible = False
End Sub

Private Sub Command16_Click()
print_mode = 1
Me.Command16.BackColor = &HFF00&
Me.Command16.Enabled = False
Me.Command17.BackColor = &HC0FFC0
Me.Command17.Enabled = True




End Sub

Private Sub Command17_Click()
print_mode = 2
Me.Command17.BackColor = &HFF00&
Me.Command17.Enabled = False
Me.Command16.Enabled = True
Me.Command16.BackColor = &HC0FFC0

End Sub
'Photoshop
Private Sub Command18_Click()

Set appRef = CreateObject("photoshop.application")
Set docref = appRef.Documents.Add(1920, 1080, 72)
docref.activelayer.kind = 2

Set artlayerref = docref.ArtLayers.Add
artlayerref.isbackgroundlayer = False

Set textItemRef = artlayerref.TextItem

a = (Me.Text1 ^ 3 * Sqr(2)) / 3

docref.activelayer.TextItem.Contents = "Объем параллелепипеда" & Format(CDbl(Me.Text1.Text), "0.000") & vbCrLf & "   равен " & Format(a, "0.000")
docref.activelayer.TextItem.Size = 30

End Sub
'MS Excel
Private Sub Command2_Click()
Dim e As Object

On Error GoTo noexcel
Set e = CreateObject("excel.application")
 On Error GoTo 0
 e.Visible = True
 e.workbooks.Add
 e.ActiveSheet.Range("B7").Value = Me.Text1 * Me.Text2 * Me.Text3 & "m^3"
 Set e = Nothing
 Exit Sub
noexcel:
    MsgBox "noexcel"

End Sub
'Powerpoint
Private Sub Command3_Click()
Dim p As Object

On Error GoTo nopowerpoint
Set p = CreateObject("powerpoint.application")
 On Error GoTo 0
 p.Visible = True
 p.Presentations.Add
 Set newslide = p.activepresentation.slides.Add(1, 11)
 Set mydocument = p.activepresentation.slides(1)
 mydocument.shapes(1).textframe.textrange.Text = "Объем параллелепипеда: " & Me.Text1 * Me.Text2 * Me.Text3 & "m^3"
 p.Activate
 Set p = Nothing
 Exit Sub
nopowerpoint:
    MsgBox "nopowerpoint"

End Sub
'IE
Private Sub Command4_Click()
Dim IExplorer As Object
Set IExplorer = CreateObject("InternetExplorer.Application")
IExplorer.Visible = True
'IExplorer.Navigate "объём параллелепипеда = " & (Me.Text1 ^ 3 * Sqr(2)) / 3
IExplorer.Navigate "https://www-formula.ru/2011-09-21-10-52-19"   ' Ссылка на сайт с ртасчетом объема фигуры
Set IExplorer = Nothing

'                                                                   Нерабочая часть кода с сайтоми заполнением переменных.
'Dim IExplorer As Object

'Set IExplorer = CreateObject("InternetExplorer.Application")
       ' IExplorer.Visible = True
       ' IExplorer.Navigate "https://www-formula.ru/2011-09-21-10-52-19"
          '  Do While IExplorer.busy = True And Not IExplorer.readystate = 4
             '   DoEvents
           ' Loop
   ' With IExplorer

           ' IExplorer.Document.getelementsbyclassname("val_a").Item(0).Value = Me.Text1.Text
           ' IExplorer.Document.getelementsbyclassname("val_b").Item(0).Value = Me.Text2.Text
           ' IExplorer.Document.getelementsbyclassname("val_c").Item(0).Value = Me.Text3.Text
            '.Document.All("calc_button71").Click

    'End With

'Set IExplorer = Nothing
End Sub

'SolidWorks''''''''''''''''''''''''

'Dim Part As Object
'Dim boolstatus As Boolean
'Dim longstatus As Long, longwarnings As Long
   ' Dim swApp As SldWorks.SldWorks
    'Dim swmodel As SldWorks.ModelDoc2
    'Dim swSelMgr As SldWorks.SelectionMgr
   ' Dim swComp As SldWorks.Component2
   ' Dim swCompModel As SldWorks.ModelDoc2
   ' Dim swCompBody As SldWorks.Body2
    'Dim vMassProps As Variant
    'Dim nDenesity As Double
    'Dim bRet As Boolean
Private Sub Command5_Click()
Dim Part As Object
Dim longstatus As Long
On Error Resume Next
Set swApp = CreateObject("sldworks.application")
swApp.Visible = True
Set Part = swApp.NewDocument("C:\Program Files\SolidWorks Corp\SolidWorks\lang\russian\Tutorial\part.prtdot", 0, 0, 0)
swApp.ActivateDoc2 "Деталь2", False, longstatus
Set Part = swApp.ActiveDoc
Dim myModelView As Object
Set myModelView = Part.ActiveView
myModelView.FrameState = swWindowState_e.swWindowMaximized
boolstatus = Part.Extension.SelectByID2("Сверху", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
Part.SketchManager.InsertSketch True
Part.ClearSelection2 True
Dim vSkLines As Variant
vSkLines = Part.SketchManager.CreateCenterRectangle(0, 0, 0, Me.Text1.Text / 1000, Me.Text1.Text / 1000, 0)
Part.ClearSelection2 True
Part.SketchManager.InsertSketch True
Part.ShowNamedView2 "*Триметрия", 8
Part.ClearSelection2 True
boolstatus = Part.Extension.SelectByID2("Line5", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
boolstatus = Part.Extension.SelectByID2("Line6", "SKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0)
boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", 0, 0, 0, True, 0, Nothing, 0)
boolstatus = Part.Extension.SelectByID2("Line2", "SKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0)
boolstatus = Part.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0)
boolstatus = Part.Extension.SelectByID2("Line4", "SKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0)
boolstatus = Part.Extension.SelectByID2("Line3", "SKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0)
Dim myFeature As Object
Set myFeature = Part.FeatureManager.FeatureExtrusion2(True, False, False, 0, 0, Me.Text3.Text / 1000, 0.01, False, False, False, False, 0.5235987755983, 0.5235987755983, False, False, False, False, True, True, True, 0, 0, False)
Part.SelectionManager.EnableContourSelection = False

Set swmass = Part.Extension.CreateMassProperty
dvolume = swmass.Volume
swApp.SendMsgToUser ("Объем параллелепипеда: " & dvolume * 10 ^ 9 & "m^3")
Me.Label1.Caption = dvolume * 10 ^ 9 & "mm^3"
End Sub
'Open button
Private Sub Command6_Click()

Dim inData

Me.CommonDialog1.Filter = "Tекстовый файл (.txt)|*.txt"    '   C:\Documents\
Me.CommonDialog1.ShowOpen
  Open Me.CommonDialog1.FileName For Input As #1
  Input #1, inData
  bb = inData
  Close #1

For i = 1 To Len(bb)
    yy = Mid(bb, i, 1) Like "#"
    If yy = True Then kk = i
Next

Me.Text1.Text = kk

End Sub
'Save button
Private Sub Command7_Click()
s = ((Me.Text1.Text ^ 3) * Sqr(2)) / 3
Me.CommonDialog1.Filter = "TXT|*.txt"
Me.CommonDialog1.ShowSave
Open Me.CommonDialog1.FileName For Output As #2
  Print #2, "объем параллелепипеда " & " " & Me.Text1.Text & " " & " мм равен   " & s
  Close #2
End Sub
'Volume button
Private Sub Command8_Click()
s = ((Me.Text1.Text ^ 3) * Sqr(2)) / 3
Dim sss As Object
Set sss = CreateObject("SAPI.SpVoice")
sss.Speak "volume of cuboid is " & (Replace(Format(s, "0.0"), ",", "."))

Const SAFT48kHz16BitStereo = 39
Const SSFMCreateForWrite = 3
Dim oFileStream, oVoice

Set oFileStream = CreateObject("SAPI.SpFileStream")
oFileStream.Format.Type = SAFT48kHz16BitStereo
oFileStream.Open "C:\Test\Sample.wav", SSFMCreateForWrite

Set oVoice = CreateObject("SAPI.SpVoice")
Set oVoice.AudioOutputStream = oFileStream
oVoice.Speak "volume of cuboid is " & (Replace(Format(s, "0.0"), ",", ".")) & "                                                                                         "

oFileStream.Close

End Sub
'Copy button
Private Sub Command9_Click()
s = ((Me.Text1.Text ^ 3) * Sqr(2)) / 3
a = "volume of cuboid " & " " & Me.Text1.Text & " " & " mm is " & s

Clipboard.SetText (a)
tet = Clipboard.GetText()
MsgBox (tet)
End Sub
'Print
Private Sub Command10_Click()
If print_mode <> 0 Then
    If print_mode = 1 Then
        Call printr
    End If
    If print_mode = 2 Then
        Call prr
    End If
Else
    MsgBox "Choose Printer Mode", vbCritical, "ERROR"
End If

End Sub


Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    Me.PopupMenu mmSort
End If
End Sub
'Printer
Sub printr()
draw_mode = True
Call draw
Printer.ScaleMode = vbCentimeters
Printer.PrintQuality = 10
Printer.PaperSize = 5
Printer.Orientation = 1
Printer.PaintPicture Me.Picture1.Image, 0, 0, (Printer.ScaleWidth * 4) / 4, (Printer.ScaleWidth * 3) / 4
draw_mode = False
Call draw
Printer.EndDoc


End Sub
Private Sub mSortV_Click()
Me.MSFlexGrid1.Sort = 1
End Sub
