VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Elevator"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   7140
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   525
      Left            =   120
      TabIndex        =   2
      Text            =   "1"
      Top             =   3240
      Width           =   1995
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   3810
      Width           =   2685
   End
   Begin VB.TextBox Text1 
      Height          =   2985
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   4485
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim AutoCAD As acadapplication
Dim Thisdrawing As AcadDocument  'as Object
 Set acadapp = GetObject(, "autocad.application")
    
Set Thisdrawing = acadapp.activedocument
Dim activedocument As Object
Dim modelspace As AcadModelSpace
Dim paperspace As AcadPaperSpace
Dim SSET As Object
Set acadapp = GetObject(, "AutoCAD.Application")
Set Thisdrawing = acadapp.activedocument
Set modelspace = Thisdrawing.modelspace
Set paperspace = Thisdrawing.paperspace
 Dim insertionPoint(0 To 2) As Variant
 Dim x
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
Apptivate
On Error Resume Next
   Set SSET2 = Thisdrawing.SelectionSets.Add("strSet")
'On Error Resume Next
SSET2.SelectOnScreen
    
    For Each ent In SSET2
    
    
    List1.AddItem ent.entityname & " " & Val(ent.Elevation)
    x = ent.StartPoint
    List1.AddItem ent.entityname & " " & Val(x(2))
   Text1.text = Text1.text & ent.entityname & " " & ent.Elevation & vbCrLf
     newValue = Val(ent.Elevation) * Val(Text2.text)
     
     ent.Elevation = newValue
     
     Text1.text = Text1.text & ent.entityname & " " & ent.Elevation & "New Elevation: " & newValue & vbCrLf
     
    Next ent
    
    SSET2.Delete
End Sub
