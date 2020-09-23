Attribute VB_Name = "Module1"

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Function Apptivate()
Dim Thisdrawing As Object
Dim modelspace As Object
Dim strCaption As String
'Dim modelspace As ACADMODELSPACE
'Dim paperspace As AcadPaperSpace


On Error GoTo ZOOP:

Set acadapp = GetObject(, "AutoCAD.Application")
        If Err Then
            Err.Clear
        Set acadapp = CreateObject("AutoCAD.Application")
            acadapp.Visible = True
                If Err Then
            
       End If
    End If
   strCaption = acadapp.Caption
    
ipos = InStr(1, strCaption, " -")
ipos2 = InStr(1, strCaption, " [")
ipos3 = InStr(1, strCaption, " 2000")


If ipos > 0 Then
strStripCaption = Mid(strCaption, 1, ipos)
End If
If ipos2 > 0 Then
strStripCaption = Mid(strCaption, 1, ipos2)
End If
If ipos3 > 0 Then
strStripCaption = Mid(strCaption, 1, ipos3)
End If



    
    
    Set acadapp = GetObject(, "autocad.application")
        Set Thisdrawing = acadapp.activedocument
            acadapp.Visible = True
    'AppActivate "autocad"
    AppActivate Trim(strStripCaption)
    
    Exit Function
ZOOP:

End Function
Private Sub Command51_Click()
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
 Dim insertionPoint2  '(0 To 2) 'As Variant
 Dim movex(0 To 2) As Double
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
strSet = Trim("SSET" & Trim(ripTime))



Apptivate
'On Error Resume Next
Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))
'Apptivate
SSET2.SelectOnScreen
    
    For Each ent In SSET2
    
    If ent.EntityType = acText Then
    cnt = cnt + 1
    exerx1 = ent.insertionPoint
    newtext = ent.TextString
    For v = 0 To 4 'set this to the copies needed
      
        With Thisdrawing.Utility
        .InitializeUserInput 1
    
        insertionPoint2 = Thisdrawing.Utility.GetPoint(, vbCr & "insertion: ")
        End With
   
    movex(0) = insertionPoint2(0)
    movex(1) = insertionPoint2(1)
   
    ent.Copy
    ent.TextString = newtext & Str(v) ' set this with an increment greater than 1 if needed, or -
    ent.insertionPoint = movex
   
    Next
     
    End If
    
   
    Next ent
    SSET2.Delete
End Sub
