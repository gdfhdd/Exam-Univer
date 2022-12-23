VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form1 
   Caption         =   "окоргир"
   ClientHeight    =   8730
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   13095
   LinkTopic       =   "Form1"
   Picture         =   "forma.frx":0000
   ScaleHeight     =   8730
   ScaleWidth      =   13095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command13 
      Caption         =   "энтернэт"
      Height          =   1215
      Left            =   2400
      Picture         =   "forma.frx":141D62
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton Command12 
      Caption         =   "орхив"
      Height          =   1215
      Left            =   9240
      Picture         =   "forma.frx":144DA4
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5040
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1695
      Left            =   7080
      TabIndex        =   14
      Top             =   1800
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2990
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      AllowUserResizing=   1
      FormatString    =   "Сторона | Объем"
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Add new value"
      Default         =   -1  'True
      Height          =   615
      Left            =   7080
      TabIndex        =   13
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton Command10 
      Caption         =   "принтер"
      Height          =   1215
      Left            =   240
      TabIndex        =   12
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Копи"
      Height          =   1215
      Left            =   120
      Picture         =   "forma.frx":147DE6
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "волуме"
      Height          =   1335
      Left            =   6960
      OLEDropMode     =   1  'Manual
      Picture         =   "forma.frx":14AE28
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4800
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10560
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command7 
      Caption         =   "save "
      Height          =   1455
      Left            =   10680
      Picture         =   "forma.frx":14DE6A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "open"
      Height          =   1335
      Left            =   9000
      Picture         =   "forma.frx":150EAC
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "solidworks"
      Height          =   975
      Left            =   7320
      Picture         =   "forma.frx":153EEE
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "exployer"
      Height          =   1095
      Left            =   5760
      Picture         =   "forma.frx":155A30
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PP"
      Height          =   1095
      Left            =   4200
      Picture         =   "forma.frx":157572
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EXCEL"
      Height          =   1095
      Left            =   2880
      Picture         =   "forma.frx":1590B4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "WORD"
      Height          =   975
      Left            =   1440
      Picture         =   "forma.frx":15ABF6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   1440
      TabIndex        =   0
      Top             =   1200
      Width           =   5655
      Begin VB.TextBox Text1 
         Height          =   1935
         Left            =   2040
         TabIndex        =   2
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "сторона октаэдра"
         Height          =   975
         Left            =   480
         TabIndex        =   1
         Top             =   840
         Width           =   1935
      End
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
Dim vlist()



Private Sub Command1_Click()
Dim w As Object

Debug.Print Me.Text1 ^ 3

On Error GoTo noword
Set w = CreateObject("word.application")
 On Error GoTo 0
 w.Visible = True
 w.documents.Add
 w.selection.typetext "объём октаэдра = " & (Me.Text1 ^ 3 * Sqr(2)) / 3
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
Me.CommonDialog1.FileName = ""
Me.CommonDialog1.ShowSave

If Me.CommonDialog1.FileName <> "" Then
    Dim ShellApp As Object
    Open Me.CommonDialog1.FileName For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
    
    Set ShellApp = CreateObject("shell.application")
    filetozip = "C:\Temp\фыр.txt"
    ShellApp.namespace(Me.CommonDialog1.FileName).copyhere filetozip
End If

End Sub


Private Sub Command13_Click()
Shell "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe https://onlinegdb.com/OeTv1VrMWG/"
End Sub

Private Sub Command2_Click()
Dim e As Object

On Error GoTo noexcel
Set e = CreateObject("excel.application")
 On Error GoTo 0
 e.Visible = True
 e.workbooks.Add
 e.ActiveSheet.Range("A1").Value = "объём октаэдра = " & (Me.Text1 ^ 3 * Sqr(2)) / 3
 Set e = Nothing
 Exit Sub
noexcel:
    MsgBox "noexcel"

End Sub

Private Sub Command3_Click()
Dim p As Object


On Error GoTo nopowerpoint
Set p = CreateObject("powerpoint.application")
 On Error GoTo 0
 p.Visible = True
 p.Presentations.Add
 Set newslide = p.activepresentation.slides.Add(1, 11)
 Set bb = p.activepresentation.slides(1)
 bb.shapes(1).textframe.textrange.Text = "объём октаэдра = " & (Me.Text1 ^ 3 * Sqr(2)) / 3
 p.Activate
 Set p = Nothing
 Exit Sub
nopowerpoint:
    MsgBox "nopowerpoint"

End Sub

Private Sub Command4_Click()
Dim IExplorer As Object
Set IExplorer = CreateObject("InternetExplorer.Application")
IExplorer.Visible = True
IExplorer.Navigate "объём октаэдра = " & (Me.Text1 ^ 3 * Sqr(2)) / 3
Set IExplorer = Nothing
End Sub

Private Sub Command5_Click()
Dim part As Object
Dim longstatus As Long
On Error Resume Next
Set swapp = CreateObject("sldworks.application")
swapp.Visible = True
Set part = swapp.NewDocument("C:\Program Files\SolidWorks Corp\SolidWorks\lang\russian\Tutorial\part.prtdot", 0, 0, 0)
swapp.ActivateDoc2 "Деталь2", False, longstatus
Set part = swapp.ActiveDoc
Dim myModelView As Object
Set myModelView = part.ActiveView
myModelView.FrameState = swWindowState_e.swWindowMaximized
boolstatus = part.Extension.SelectByID2("Сверху", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
part.SketchManager.InsertSketch True
part.ClearSelection2 True
Dim vSkLines As Variant
vSkLines = part.SketchManager.CreateCenterRectangle(0, 0, 0, Me.Text1.Text / 1000, Me.Text1.Text / 1000, 0)
part.ClearSelection2 True
part.SketchManager.InsertSketch True
part.ShowNamedView2 "*Триметрия", 8
part.ClearSelection2 True
boolstatus = part.Extension.SelectByID2("Эскиз1", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
Dim myFeature As Object
Set myFeature = part.FeatureManager.FeatureExtrusion2(False, False, False, 0, 0, Me.Text1.Text / 1000, Me.Text1.Text / 1000, True, True, False, False, 0.78539816339745, 0.78539816339745, False, False, False, False, True, True, True, 0, 0, False)
part.SelectionManager.EnableContourSelection = False



Set swmass = part.Extension.createmassproperty
dvolume = swmass.Volume
swapp.sendmsgtouser ("объём октаэдра = " & ((Me.Text1.Text ^ 3) * Sqr(2))) / 3
Me.Label2.Caption = (((Me.Text1.Text ^ 3) * Sqr(2)) / 3)
End Sub

Private Sub Command6_Click()
Dim inData
Me.CommonDialog1.Filter = "Tекстовый файл (.txt)|*.txt"
Me.CommonDialog1.ShowOpen
  Open Me.CommonDialog1.FileName For Input As #1
  Input #1, inData
  bb = inData
  Close #1

For i = 1 To Len(bb)
    yy = Mid(bb, i, 1) Like "#"
    If yy = True Then kk = i
'    For j = 0 To 9
'        If Str(Mid(bb, i, 1)) = Str(j) Then
'        kk = Mid(bb, i, 1)
'        End If
'    Next
Next


'ff = InStr(1, bb, )
'bb = Mid(bb, ff, 3)
Me.Text1.Text = kk
End Sub

Private Sub Command7_Click()
s = ((Me.Text1.Text ^ 3) * Sqr(2)) / 3
Me.CommonDialog1.Filter = "TXT|*.txt"
Me.CommonDialog1.ShowSave
Open Me.CommonDialog1.FileName For Output As #2
  Print #2, "объем октаэдра с длиной ребра " & " " & Me.Text1.Text & " " & " мм равен   " & s
  Close #2
End Sub



Private Sub Command8_Click()
s = ((Me.Text1.Text ^ 3) * Sqr(2)) / 3
Dim sss As Object
Set sss = CreateObject("SAPI.SpVoice")
sss.Speak "volume of octahedron is " & (Replace(Format(s, "0.0"), ",", "."))

Const SAFT48kHz16BitStereo = 39
Const SSFMCreateForWrite = 3 ' Creates file even if file exists and so destroys or overwrites the existing file

Dim oFileStream, oVoice

Set oFileStream = CreateObject("SAPI.SpFileStream")
oFileStream.Format.Type = SAFT48kHz16BitStereo
oFileStream.Open "C:\Test\Sample.wav", SSFMCreateForWrite

Set oVoice = CreateObject("SAPI.SpVoice")
Set oVoice.AudioOutputStream = oFileStream
oVoice.Speak "volume of octahedron is " & (Replace(Format(s, "0.0"), ",", ".")) & "                                                                                         "

oFileStream.Close

End Sub


Private Sub Command9_Click()
s = ((Me.Text1.Text ^ 3) * Sqr(2)) / 3
a = "volume of octahedron so storonoy " & " " & Me.Text1.Text & " " & " mm is " & s

Clipboard.SetText (a)
tet = Clipboard.GetText()
MsgBox (tet)
End Sub

Private Sub Command10_Click()
   ss = ((Me.Text1.Text ^ 3) * Sqr(2)) / 3
a = "volume of octahedron so storonoy " & " " & Me.Text1.Text & " " & " mm is " & ss
    
    Printer.Print a
    Printer.EndDoc
    
    
End Sub


Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    Me.PopupMenu mmSort
End If
End Sub

Private Sub mSortV_Click()
Me.MSFlexGrid1.Sort = 1
End Sub

