VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   ScaleHeight     =   475
   ScaleMode       =   3  '�������
   ScaleWidth      =   709
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "������"
      Height          =   495
      Left            =   6960
      TabIndex        =   29
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "���������"
      Height          =   495
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "����������"
      Height          =   495
      Left            =   4320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '���������
      Caption         =   "��� ""��������� ������"""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   4215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '���������
      Caption         =   "����� �����"
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   11
      Left            =   4560
      TabIndex        =   30
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   3975
      Left            =   240
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800080&
      BorderStyle     =   2  'Dash
      BorderWidth     =   2
      X1              =   408
      X2              =   408
      Y1              =   120
      Y2              =   408
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800080&
      BorderWidth     =   2
      X1              =   368
      X2              =   624
      Y1              =   88
      Y2              =   88
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '���������
      Caption         =   "�������"
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   10
      Left            =   4560
      TabIndex        =   28
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '���������
      Caption         =   "��� ������"
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   9
      Left            =   4560
      TabIndex        =   27
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '���������
      Caption         =   "�����������"
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   8
      Left            =   4560
      TabIndex        =   26
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '���������
      Caption         =   "������"
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   7
      Left            =   4560
      TabIndex        =   25
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '���������
      Caption         =   "���"
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   6
      Left            =   4560
      TabIndex        =   24
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '���������
      Caption         =   "�������"
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   5
      Left            =   4560
      TabIndex        =   23
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '���������
      Caption         =   "�����������"
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   4
      Left            =   4560
      TabIndex        =   22
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '���������
      Caption         =   "��� �������"
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   3
      Left            =   4560
      TabIndex        =   21
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '���������
      Caption         =   "��������� ������"
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   20
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '���������
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   615
      Index           =   12
      Left            =   1680
      TabIndex        =   19
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '���������
      Caption         =   "Label9"
      Height          =   255
      Index           =   11
      Left            =   6240
      TabIndex        =   18
      Top             =   5760
      Width           =   4215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '���������
      Caption         =   "Label9"
      Height          =   255
      Index           =   10
      Left            =   6240
      TabIndex        =   17
      Top             =   5400
      Width           =   4215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '���������
      Caption         =   "Label9"
      Height          =   255
      Index           =   9
      Left            =   6240
      TabIndex        =   16
      Top             =   5040
      Width           =   4215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '���������
      Caption         =   "Label9"
      Height          =   855
      Index           =   8
      Left            =   6240
      TabIndex        =   15
      Top             =   4080
      Width           =   4215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '���������
      Caption         =   "Label9"
      Height          =   255
      Index           =   7
      Left            =   6240
      TabIndex        =   14
      Top             =   3720
      Width           =   4215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '���������
      Caption         =   "Label9"
      Height          =   255
      Index           =   6
      Left            =   6240
      TabIndex        =   13
      Top             =   3360
      Width           =   4215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '���������
      Caption         =   "Label9"
      Height          =   255
      Index           =   5
      Left            =   6240
      TabIndex        =   12
      Top             =   2640
      Width           =   4215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '���������
      Caption         =   "Label9"
      Height          =   255
      Index           =   4
      Left            =   6240
      TabIndex        =   11
      Top             =   3000
      Width           =   4215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '���������
      Caption         =   "Label9"
      Height          =   255
      Index           =   3
      Left            =   6240
      TabIndex        =   10
      Top             =   2280
      Width           =   4215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '���������
      Caption         =   "Label9"
      Height          =   255
      Index           =   2
      Left            =   6240
      TabIndex        =   9
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Label Label9 
      Alignment       =   2  '���������
      BackStyle       =   0  '���������
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   375
      Index           =   1
      Left            =   8040
      TabIndex        =   8
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '���������
      Caption         =   "������� ����� �������� �����"
      Height          =   255
      Left            =   7920
      TabIndex        =   7
      Top             =   0
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800080&
      BorderWidth     =   3
      Height          =   5655
      Left            =   4320
      Top             =   720
      Width           =   6135
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '���������
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   6840
      TabIndex        =   6
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  '���������
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '���������
      Caption         =   "���."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  '���������
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '���������
      Caption         =   "��������� �������"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   3735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '���������
      Caption         =   "����������� ��������������"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '���������
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   3750
      Left            =   360
      Top             =   1920
      Width           =   3750
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Dim MassiveData(12) As String
Dim FileName As String
Dim F As Long
Dim X As Long
Dim QuantityFonts As Long

Private Sub Form_Load()
FileName = "1"
DataLoading
End Sub

Private Sub DataLoading()
Image1.Picture = LoadPicture(App.Path & "\" & FileName & ".jpg")
Image1.Left = 24
Image1.Top = 128
If Image1.Width > Image1.Height Then Image1.Top = Image1.Top + (250 - Image1.Height) / 2
If Image1.Height > Image1.Width Then Image1.Left = Image1.Left + (250 - Image1.Width) / 2
F = FreeFile
Open App.Path & "\" & FileName & ".txt" For Input As #F
For X = 1 To 12
Line Input #F, MassiveData(X)
Label9(X).Caption = MassiveData(X)
Next X
Close #F

End Sub

Private Sub Command1_Click()
If Val(FileName) > 1 Then
FileName = Str(Val(FileName) - 1)
FileName = Right(FileName, Len(FileName) - 1)
DataLoading
End If
End Sub

Private Sub Command2_Click()
FileName = Str(Val(FileName) + 1)
FileName = Right(FileName, Len(FileName) - 1)
If Dir(App.Path & "\" & FileName & ".txt") <> "" Or Dir(App.Path & "\" & FileName & ".jpg") <> "" Then
DataLoading
Else
FileName = Str(Val(FileName) - 1)
FileName = Right(FileName, Len(FileName) - 1)
DataLoading
End If
End Sub

Private Sub Command3_Click()
Dim OrientTelefona As Single
Dim Fonts() As String
Dim VertCoord As Single
Dim NumberFont As Long
QuantityFonts = 0


'������� ������� ���� ��������������� ������� �� 4
   For NumberFont = 0 To Printer.FontCount - 1
       If Printer.Fonts(NumberFont) = "Arial Cyr" Or Printer.Fonts(NumberFont) = "Times New Roman" _
       Or Printer.Fonts(NumberFont) = "Courier New" Or Printer.Fonts(NumberFont) = "MS Sans Serif" Then
            QuantityFonts = QuantityFonts + 1
            ReDim Preserve Fonts(QuantityFonts)
            Fonts(QuantityFonts) = Printer.Fonts(NumberFont)
            
       End If
    Next NumberFont


'������������� �����
If QuantityFonts > 0 Then Printer.FontName = Fonts(1)
'������������� ����������� � �����������
Printer.ScaleMode = vbCentimeters

'������ ������
Printer.FontSize = 12
'������������� �������������
Printer.Font.Underline = True
'��������������� �������� ������ - �������
Printer.PrintQuality = 3
'��������� ������ ������ ������� � ��������� �����
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth(Label1.Caption)
'����������, ������� ������ ������ ������ ������
VertCoord = Printer.TextHeight(Label1.Caption)
            '���������� � �������� ������ �������
            Printer.Print Label1.Caption
'������������� ������� ������ ������ ��� �������� ��������
Printer.FontSize = 20
'�������� �������������
Printer.Font.Underline = False
'���� ������ ������
Printer.Font.Bold = True
'������ ������ ���� �� 1,5 �� �� ������ �������, ����� ������������ ����������
Printer.CurrentY = VertCoord + 1.5
'����������� ������ ������� ��-��������
Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(Label9(1).Caption & "����� ������ - ������� ")) / 2
'����������, ������� ������ �� ������������ ��  ��� ������� ������ � ���������
VertCoord = VertCoord + Printer.TextHeight(Label9(1).Caption) + 1.5
            '���������� � �������� ������ �������
            Printer.Print "����� ������ - ������� " & Label9(1).Caption
            

 
'������������� ������ ������ �������� ��� �������� ����������� ������������� ��������
Printer.FontSize = 14
'��������� ������
Printer.Font.Italic = True
'��������� ���� ��� �� 2 ��
Printer.CurrentY = VertCoord + 2
'����������, ������� ������ �� ������������ ��  ��� ������� ������ � ������������������ ������ ������ ������
VertCoord = VertCoord + Printer.TextHeight(Label4.Caption) + 2
'��������� 8 �� ����� ��� ��������, � ��������� ������ - ��� �����
'������� ���������� �� �� Printer.ScaleWidth - 8 ��
Printer.CurrentX = (Printer.ScaleWidth - 8 - Printer.TextWidth(Label4.Caption)) / 2
            '���������� � �������� ������ �������
            Printer.Print Label4.Caption
            

'Printer.ScaleWidth-8 � ����� ������� ���������� �� ��������� VertCoord
'�������� ��������, ��������� �� �� ����� �������
If Image1.Height > Image1.Width Then
OrientTelefona = 2.25
Else
OrientTelefona = 0.5
End If

Printer.PaintPicture Image1.Picture, Printer.ScaleWidth - ScaleX(Image1.Width, vbPixels, vbCentimeters) - OrientTelefona, VertCoord
            
'�������� ������ � ��������
Printer.Font.Italic = False
Printer.Font.Bold = False
'������������� ������ ������ ��� ������
Printer.FontSize = 12
'��������� ���� �� 3-�� ������ �� 1 ��
Printer.CurrentY = VertCoord + 1
VertCoord = VertCoord + 1
'������ � ����� ������������� ����������� ������ (����� 10 �����)

For X = 2 To 11
'��������� ����� �� ����������
Printer.CurrentX = 1
Printer.CurrentY = VertCoord
Printer.Print Label8(X).Caption

Printer.CurrentX = 5
Printer.CurrentY = VertCoord
Printer.Print Label9(X).Caption
VertCoord = VertCoord + 0.6

Next X
'��������� ��� 1 �� � �������� �����
VertCoord = VertCoord + 1
Printer.Line (1, VertCoord)-(17.5, VertCoord + 0.7), vbBlack, BF

Printer.CurrentY = VertCoord + 0.1
Printer.CurrentX = 6
'������ ����� ����
Printer.Font.Bold = True
Printer.Print "����� ������ ����: " & Label9(12).Caption & " ���."
Printer.Font.Bold = False
Printer.EndDoc

End Sub



