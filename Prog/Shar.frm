VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.Ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Расчёт объёма шара"
   ClientHeight    =   4104
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   342
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   910
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1572
      Left            =   8160
      ScaleHeight     =   1524
      ScaleWidth      =   2604
      TabIndex        =   22
      Top             =   120
      Width           =   2652
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1572
      Left            =   5160
      TabIndex        =   21
      Top             =   120
      Width           =   2892
      _ExtentX        =   5101
      _ExtentY        =   2773
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      FormatString    =   "N      | Radius mm | Volume mm^3       "
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   492
      Left            =   3600
      TabIndex        =   20
      Top             =   120
      Width           =   1452
      _ExtentX        =   2561
      _ExtentY        =   868
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Shar.frx":0000
   End
   Begin VB.CommandButton Command16 
      Height          =   612
      Left            =   10200
      Picture         =   "Shar.frx":008D
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "artificial intelligence"
      Top             =   1800
      Width           =   612
   End
   Begin VB.CommandButton Command15 
      Height          =   612
      Left            =   9480
      Picture         =   "Shar.frx":1FE3
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "server CS: 1.6"
      Top             =   1800
      Width           =   612
   End
   Begin VB.CommandButton Command14 
      Height          =   612
      Left            =   8760
      Picture         =   "Shar.frx":3F39
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "crypto"
      Top             =   1800
      Width           =   612
   End
   Begin VB.CommandButton Command13 
      Height          =   612
      Left            =   8040
      Picture         =   "Shar.frx":5E8F
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1800
      Width           =   612
   End
   Begin VB.CommandButton Command12 
      Height          =   612
      Left            =   7320
      Picture         =   "Shar.frx":7DE5
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1800
      Width           =   612
   End
   Begin VB.CommandButton Command11 
      Height          =   612
      Left            =   6600
      Picture         =   "Shar.frx":9D3B
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "text to speach"
      Top             =   1800
      Width           =   612
   End
   Begin VB.CommandButton Command10 
      Height          =   612
      Left            =   5880
      Picture         =   "Shar.frx":BC91
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "clipboard"
      Top             =   1800
      Width           =   612
   End
   Begin VB.CommandButton Command9 
      Height          =   612
      Left            =   5160
      Picture         =   "Shar.frx":DBE7
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "print"
      Top             =   1800
      Width           =   612
   End
   Begin VB.CommandButton Command8 
      Height          =   612
      Left            =   4440
      Picture         =   "Shar.frx":FB3D
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "open"
      Top             =   1800
      Width           =   612
   End
   Begin VB.CommandButton Command7 
      DisabledPicture =   "Shar.frx":11A93
      DownPicture     =   "Shar.frx":139E9
      Height          =   612
      Left            =   3720
      Picture         =   "Shar.frx":1593F
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1800
      Width           =   612
   End
   Begin VB.CommandButton Command6 
      DisabledPicture =   "Shar.frx":17895
      DownPicture     =   "Shar.frx":197EB
      Height          =   612
      Left            =   3000
      Picture         =   "Shar.frx":1B741
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1800
      Width           =   612
   End
   Begin VB.CommandButton Command5 
      Height          =   612
      Left            =   120
      Picture         =   "Shar.frx":1D697
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "word"
      Top             =   1800
      Width           =   612
   End
   Begin VB.CommandButton Command4 
      Height          =   612
      Left            =   840
      Picture         =   "Shar.frx":1F1D9
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   612
   End
   Begin VB.CommandButton Command3 
      Height          =   612
      Left            =   1560
      Picture         =   "Shar.frx":20D1B
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1800
      Width           =   612
   End
   Begin VB.CommandButton Command2 
      Height          =   612
      Left            =   2280
      Picture         =   "Shar.frx":2285D
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   612
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Расчёт"
      Height          =   492
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3372
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Left            =   1560
      TabIndex        =   0
      Text            =   "50"
      Top             =   120
      Width           =   1932
   End
   Begin VB.Label Label3 
      Caption         =   "Объём:"
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   612
   End
   Begin VB.Label Label2 
      Caption         =   "Введите радиус шара, мм"
      Height          =   492
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1332
   End
   Begin VB.Label Label1 
      Height          =   372
      Left            =   840
      TabIndex        =   2
      Top             =   1320
      Width           =   2652
   End
   Begin VB.Menu sorting 
      Caption         =   "sorting"
      Visible         =   0   'False
      Begin VB.Menu SortV 
         Caption         =   "Sort by Volume"
      End
      Begin VB.Menu sortR 
         Caption         =   "Sort by Radius"
      End
      Begin VB.Menu sortN 
         Caption         =   "Sort by Number"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Part As Object
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long
    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swSelMgr As SldWorks.SelectionMgr
    Dim swComp As SldWorks.Component2
    Dim swCompModel As SldWorks.ModelDoc2
    Dim swCompBody As SldWorks.Body2
    Dim vMassProps As Variant
    Dim nDensity As Double
    Dim bRet As Boolean
    Const pi = 3.141592
    

Private Sub Command1_Click()

Set swApp = CreateObject("SldWorks.Application")
swApp.Visible = True
Set Part = swApp.NewDocument("C:\Program Files\SolidWorks Corp\SolidWorks\lang\russian\Tutorial\part.prtdot", 0, 0, 0)
swApp.ActivateDoc2 "Деталь1", False, longstatus
Set Part = swApp.ActiveDoc
Dim myModelView As Object
Set myModelView = Part.ActiveView
'myModelView.FrameState = swWindowState_e.swWindowMaximized
boolstatus = Part.Extension.SelectByID2("Front", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
Part.SketchManager.InsertSketch True
Part.ClearSelection2 True
Dim skSegment As Object
Set skSegment = Part.SketchManager.CreateArc(0#, 0#, 0#, (-Me.Text1.Text / 1000), 0#, 0#, (Me.Text1.Text / 1000), 0#, 0#, -1)
Part.ClearSelection2 True
Set skSegment = Part.SketchManager.CreateLine((-Me.Text1.Text / 1000), 0#, 0#, (Me.Text1.Text / 1000), 0#, 0#)
Part.ClearSelection2 True
boolstatus = Part.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", -5.00603044909766E-02, 7.53258500177705E-03, 0, False, 0, Nothing, 0)
'Dim myDisplayDim As Object
'Set myDisplayDim = Part.AddDimension2(2.71486917772381E-02, 2.10284664632942E-02, 0)
'Part.ClearSelection2 True
'Dim myDimension As Object
'Set myDimension = Part.Parameter("D1@Эскиз1")
'myDimension.SystemValue = Me.Text1.Text / 100
Part.ClearSelection2 True
Part.SketchManager.InsertSketch True
Part.ShowNamedView2 "*Триметрия", 8
boolstatus = Part.Extension.SelectByID2("Line1@Эскиз1", "EXTSKETCHSEGMENT", -3.21848014731358E-02, 0, 0, True, 0, Nothing, 0)
Part.ClearSelection2 True
boolstatus = Part.Extension.SelectByID2("Эскиз1", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
boolstatus = Part.Extension.SelectByID2("Line1@Эскиз1", "EXTSKETCHSEGMENT", -3.21848014731358E-02, 0, 0, True, 16, Nothing, 0)
Dim myFeature As Object
Set myFeature = Part.FeatureManager.FeatureRevolve2(True, True, False, False, False, False, 0, 0, 6.2831853071796, 0, False, False, 0.01, 0.01, 0, 0, 0, True, True, True)
Part.SelectionManager.EnableContourSelection = False

Set swMass = Part.Extension.CreateMassProperty
dVolume = swMass.Volume
swApp.SendMsgToUser ("Объем: " & dVolume * 10 ^ 9 & " мм^3")
Me.Label1.Caption = dVolume * 10 ^ 9 & " мм^3"

    n = Me.MSFlexGrid1.Rows
    Me.MSFlexGrid1.Rows = Me.MSFlexGrid1.Rows + 1
    
    Me.MSFlexGrid1.TextMatrix(n, 0) = n
    Me.MSFlexGrid1.TextMatrix(n, 1) = Me.Text1.Text
    Me.MSFlexGrid1.TextMatrix(n, 2) = dVolume * 10 ^ 9

'Set ppoint = CreateObject("PowerPoint.Aplication")
'ppoint.Visible = True
'PowerPoint.
End Sub

Private Sub Command10_Click()
'Clipboard.SetText "The volume of this sphere, with radius of " & Me.Text1.Text & " millimeters, is " & Format((Me.Text1.Text ^ 3) * pi * (4 / 3), "#.##") & " cubic millimeters"
Me.RichTextBox1.Text = "Объём шара радиуса " & Me.Text1.Text & " (мм) = " & Format((Me.Text1.Text ^ 3) * pi * (4 / 3), "#.##") & " (куб. мм)"
Clipboard.SetText Me.RichTextBox1.TextRTF, vbCFRTF
End Sub

Private Sub Command11_Click()
Dim s As New SpVoice, fileStream As New SpFileStream
Set s.Voice = s.GetVoices.Item(0)
s.Rate = 0.0000000000001
s.Volume = 70
s.Speak Replace("The volume of this sphere, with radius of " & Me.Text1.Text & " millimeters, is " & Format((Me.Text1.Text ^ 3) * pi * (4 / 3), "#.##") & " cubic millimeters", ",", "."), 0
FileName = "c:\Temp\ttstemp.wav"
fileStream.Open FileName, SSFMCreateForWrite, False
Set s.AudioOutputStream = fileStream
s.Speak Replace("The volume of this sphere, with radius of " & Me.Text1.Text & " millimeters, is " & 523598.8 & " cubic millimeters", ",", "."), 0
'Do
'   DoEvents
'Loop While Not s.Status.RunningState = 1
fileStream.Close
Set fileStream = Nothing
Set s = Nothing
End Sub

Private Sub Command13_Click()
    Dim np As Object
    Open "C:\Temp\SHAR-радиус-объём.txt" For Output As #1
    Print #1, "Объём шара радиуса " & Me.Text1.Text & " (мм) = " & Format((Me.Text1.Text ^ 3) * pi * (4 / 3), "#.##") & " (куб. мм)"
    Close #1
    Shell "notepad.exe " & "C:\Temp\SHAR-радиус-объём.txt", vbNormalFocus
    
End Sub

Private Sub Command2_Click()
Dim xl As Object
Dim wb As Object
Set xl = CreateObject("Excel.Application")
On Error GoTo 0
xl.Visible = True
Set wb = xl.Workbooks.Add
wb.ActiveSheet.Range("B7").Value = (Me.Text1 ^ 3) * pi * (4 / 3)
Set xl = Nothing
Set wb = Nothing
Exit Sub
noexcel:
MsgBox "NoExcel"
End Sub

Private Sub Command3_Click()
Dim ppoint As Object
On Error GoTo nopowerpoint
Set ppoint = CreateObject("PowerPoint.Application")
On Error GoTo 0
ppoint.Visible = True
ppoint.Presentations.Add
Set newslide = ppoint.ActivePresentation.Slides.Add(slidecount + 1, 11)
Set mydocument = ppoint.ActivePresentation.Slides(1)
mydocument.Shapes(1).TextFrame.TextRange.Text = "Объём шара " & (Me.Text1 ^ 3) * pi * (4 / 3) & " (куб. мм)"
ppoint.Activate
Set ppoint = Nothing
Exit Sub
nopowerpoint:
MsgBox "no powerpoint"

End Sub

Private Sub Command4_Click()
Dim IExplorer As Object

Set IExplorer = CreateObject("InternetExplorer.Application")
        IExplorer.Visible = True
        IExplorer.Navigate "https://ru.onlinemschool.com/math/assistance/figures_volume/sphere/"
            Do While IExplorer.busy = True And Not IExplorer.readystate = 4
                DoEvents
            Loop
Set IExplorer = Nothing

End Sub

Private Sub Command5_Click()
Dim wr As Object
On Error GoTo noword
Set wr = CreateObject("Word.Application")
On Error GoTo 0
wr.Visible = True
wr.Documents.Add
wr.Selection.TypeText "Объём шара " & (Me.Text1 ^ 3) * pi * (4 / 3) & " куб. мм)"
wr.Activate
Set wr = Nothing
Exit Sub
noword:
MsgBox "no word"
End Sub

Private Sub Command9_Click()
Printer.ScaleMode = vbMillimeters
Printer.Print "Объём шара радиуса " & Me.Text1.Text & " (мм) = " & Format((Me.Text1.Text ^ 3) * pi * (4 / 3), "#.##") & " (куб. мм)"
Printer.Circle (105, 297 / 2), Me.Text1.Text, vbBlack
Printer.EndDoc

End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu sorting
    End If
    
End Sub

Private Sub sortN_Click()
    Me.MSFlexGrid1.Col = 0
    Me.MSFlexGrid1.Sort = flexSortGenericAscending

End Sub

Private Sub sortR_Click()
    Me.MSFlexGrid1.Col = 1
    Me.MSFlexGrid1.Sort = flexSortGenericAscending

End Sub

Private Sub SortV_Click()
    Me.MSFlexGrid1.Col = 2
    Me.MSFlexGrid1.Sort = flexSortGenericAscending

End Sub
