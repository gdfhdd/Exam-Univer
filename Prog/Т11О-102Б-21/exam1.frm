VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   Caption         =   "Form1"
   ClientHeight    =   10020
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18360
   FillStyle       =   0  'Solid
   ForeColor       =   &H8000000C&
   LinkTopic       =   "Form1"
   Picture         =   "exam1.frx":0000
   ScaleHeight     =   10020
   ScaleWidth      =   18360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command12 
      Caption         =   "Command12"
      Height          =   1440
      Left            =   6555
      TabIndex        =   15
      Top             =   7845
      Width           =   1635
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Chrome"
      Height          =   1695
      Left            =   7680
      Picture         =   "exam1.frx":5ED5C2
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3600
      Width           =   2895
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Command10"
      Height          =   975
      Left            =   10440
      TabIndex        =   13
      Top             =   8520
      Width           =   3615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   1215
      Left            =   12480
      TabIndex        =   12
      Top             =   6720
      Width           =   2535
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   1095
      Left            =   9360
      TabIndex        =   11
      Top             =   6720
      Width           =   2175
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   1215
      Left            =   6240
      TabIndex        =   10
      Top             =   6240
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8640
      Top             =   8160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00004040&
      Caption         =   "Save"
      DownPicture     =   "exam1.frx":5EE89B
      BeginProperty Font 
         Name            =   "GothicE"
         Size            =   26.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   3120
      MaskColor       =   &H0080FFFF&
      Picture         =   "exam1.frx":608EAB
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6360
      Width           =   2415
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00000080&
      Caption         =   "SolidWorks"
      BeginProperty Font 
         Name            =   "GothicG"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   14160
      Picture         =   "exam1.frx":609DEF
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C000&
      Caption         =   "IE"
      BeginProperty Font 
         Name            =   "GothicG"
         Size            =   36
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   8640
      Picture         =   "exam1.frx":60B931
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000080FF&
      Caption         =   "PowerPoint"
      BeginProperty Font 
         Name            =   "GothicG"
         Size            =   15.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   6000
      Picture         =   "exam1.frx":60D473
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "Excel"
      BeginProperty Font 
         Name            =   "GothicG"
         Size            =   26.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3120
      Picture         =   "exam1.frx":60EFB5
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   18
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9120
      TabIndex        =   3
      Text            =   "50"
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   18
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6960
      TabIndex        =   2
      Text            =   "30"
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   18
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4800
      TabIndex        =   1
      Text            =   "20"
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF0000&
      Caption         =   "Word"
      BeginProperty Font 
         Name            =   "GothicE"
         Size            =   30
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   360
      Picture         =   "exam1.frx":610AF7
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   2400
      Picture         =   "exam1.frx":612639
      Top             =   3240
      Visible         =   0   'False
      Width           =   4710
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Volume"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   21.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   10920
      TabIndex        =   8
      Top             =   4680
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim w As Object
On Error GoTo Noword
Set w = CreateObject("Word.Application")

On Error GoTo 0

w.Visible = True
w.documents.Add
w.Selection.TypeText "Объём заданной фигуры: " & Me.Text1 * Me.Text2 * Me.Text3 & " m"
w.Selection.Font.Superscript = wdToggle
w.Selection.TypeText "3"
w.Activate
    
Set w = Nothing
Exit Sub

Noword:
    MsgBox "No word"


End Sub


Private Sub Command2_Click()

Dim x1 As Object, wb As Object

On Error GoTo noExcel
Set x1 = CreateObject("Excel.Application")

On Error GoTo 0

x1.Visible = True
Set wb = x1.workbooks.Add
wb.ActiveSheet.Range("B7").Value = Me.Text1 * Me.Text2 * Me.Text3 & "m^3"

Set x1 = Nothing
Set wb = Nothing

noExcel:

    
End Sub

Private Sub Command3_Click()

Dim p As Object
On Error GoTo nopowerpoint
Set p = CreateObject("powerpoint.Application")
On Error GoTo 0
p.Visible = True

p.presentations.Add
Set newslide = p.activepresentation.slides.Add(slidecount + 1, 11)
Set mydocument = p.activepresentation.slides(1)
mydocument.shapes(1).textframe.textrange.Text = "Объем параллелепипеда: " & Me.Text1 * Me.Text2 * Me.Text3 & "m^3"

Set mydocument = Nothing
Set p = Nothing
Exit Sub
nopowerpoint:
    MsgBox "nopowerpoint"
End Sub

Private Sub Command4_Click()

Dim IExplorer As Object

Set IExplorer = CreateObject("InternetExplorer.Application")
        IExplorer.Visible = True
        IExplorer.Navigate "https://www-formula.ru/2011-09-21-10-52-19"
            Do While IExplorer.busy = True And Not IExplorer.readystate = 4
                DoEvents
            Loop
    With IExplorer

            IExplorer.Document.getelementsbyclassname("val_a").Item(0).Value = Me.Text1.Text
            IExplorer.Document.getelementsbyclassname("val_b").Item(0).Value = Me.Text2.Text
            IExplorer.Document.getelementsbyclassname("val_c").Item(0).Value = Me.Text3.Text
            .Document.All("calc_button71").Click

    End With

Set IExplorer = Nothing
End Sub
 
Private Sub Command5_Click()
Dim Part As Object
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long
    Dim swApp As SldWorks.SldWorks
    Dim swmodel As SldWorks.ModelDoc2
    Dim swSelMgr As SldWorks.SelectionMgr
    Dim swComp As SldWorks.Component2
    Dim swCompModel As SldWorks.ModelDoc2
    Dim swCompBody As SldWorks.Body2
    Dim vMassProps As Variant
    Dim nDenesity As Double
    Dim bRet As Boolean
Set swApp = CreateObject("sldworks.application")
swApp.Visible = True

Set Part = swApp.NewDocument("C:\ProgramData\SolidWorks\SolidWorks 2011\templates\Деталь.prtdot", 0, 0, 0)
Me.Image1.Visible = True

Set Part = swApp.ActiveDoc
Dim myModelView As Object
Set myModelView = Part.ActiveView
myModelView.FrameState = swWindowState_e.swWindowMaximized
Part.SketchManager.InsertSketch True
boolstatus = Part.Extension.SelectByID2("Спереди", "PLANE", Me.Text1.Text / 2000, Me.Text2.Text / 2000, 0, False, 0, Nothing, 0)
Part.ClearSelection2 True
Dim vSkLines As Variant
vSkLines = Part.SketchManager.CreateCenterRectangle(0, 0, 0, Me.Text1.Text / 2000, Me.Text2.Text / 2000, 0)
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
    End If
    break
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
  Print #2, "объем параллелепипеда с длиной ребра " & " " & Me.Text1.Text & " " & " мм равен " & s
  Close #2
End Sub

Private Sub Command8_Click()
s = ((Me.Text1.Text ^ 3) * Sqr(2)) / 3
Dim sss As Object
Set sss = CreateObject("SAPI.SpVoice")
sss.Speak "volume of cuboid is " & (Replace(Format(s, "0.0"), ",", "."))

Const SAFT48kHz16BitStereo = 39
Const SSFMCreateForWrite = 3 ' Creates file even if file exists and so destroys or overwrites the existing file

Dim oFileStream, oVoice

Set oFileStream = CreateObject("SAPI.SpFileStream")
oFileStream.Format.Type = SAFT48kHz16BitStereo
oFileStream.Open "C:\Test\Sample.wav", SSFMCreateForWrite

Set oVoice = CreateObject("SAPI.SpVoice")
Set oVoice.AudioOutputStream = oFileStream
oVoice.Speak "volume of cupoid is " & (Replace(Format(s, "0.0"), ",", ".")) & "                                                                                         "

oFileStream.Close

End Sub

Private Sub Command9_Click()
s = ((Me.Text1.Text ^ 3) * Sqr(2)) / 3
a = "volume of cuboid so storonoy " & " " & Me.Text1.Text & " " & " mm is " & s

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


Private Sub command11_click()
'Shell "C:\Program Files\Google\Chrome\Application\chrome.exe https://www.google.com/"
Shell "C:\Program Files\Google\Chrome\Application\chrome.exe https://onlinegdb.com/aCyYDN7d4/"
End Sub

