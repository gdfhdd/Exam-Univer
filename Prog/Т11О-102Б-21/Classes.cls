VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Progs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum software
    sExcel = 1
    sWord = 2
    sPower = 3
End Enum


Private soft As Object, doc As Object, apptype As software

Public Property Get info() As Variant

End Property

Public Property Let info(ByVal vNewValue As Variant)

End Property

Public Sub start(ByVal app1 As software)
app1name = Array("", "Excel", "Word", "PowerPoint")

    Set soft = CreateObject(app1name(app1) & ".application")
    soft.Visible = True
    
    apptype = app1
    
    Select Case apptype
        Case sExcel
            Set doc = soft.workbooks.Add
            soft.ActiveSheet.Range("B7").Value = Me.Text1
            
        Case sWord
            Set doc = soft.documents.Add
            soft.Selection.TypeText = Me.Text1
        Case sPower
            soft.presentations.Add
            Set newslide = soft.activepresentation.slides.Add(slidecount + 1, 11)
            
            Set mydocument = soft.activepresentation.slides(1)
            mydocument.shapes(1).textframe.textrange.Text = Me.Text1
            
    End Select
    
    

End Sub











