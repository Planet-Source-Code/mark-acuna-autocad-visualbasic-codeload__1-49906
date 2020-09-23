VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Generic Modelspace - Artscapers.com"
   ClientHeight    =   12240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12795
   Icon            =   "MSE_Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12240
   ScaleWidth      =   12795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command78 
      Caption         =   "Excel Textcheck"
      Height          =   495
      Left            =   3480
      TabIndex        =   90
      Top             =   7530
      Width           =   1695
   End
   Begin VB.CommandButton Command77 
      Caption         =   "SelectAll to Excel"
      Height          =   495
      Left            =   3480
      TabIndex        =   89
      Top             =   7020
      Width           =   1665
   End
   Begin VB.CommandButton Command76 
      Caption         =   "Collect Insertions"
      Height          =   465
      Left            =   10350
      TabIndex        =   88
      Top             =   11100
      Width           =   1665
   End
   Begin VB.CommandButton Command75 
      Caption         =   "Insert List"
      Height          =   405
      Left            =   10350
      TabIndex        =   87
      Top             =   11580
      Width           =   1665
   End
   Begin VB.CommandButton Command74 
      Caption         =   "Command74"
      Height          =   645
      Left            =   1500
      TabIndex        =   86
      Top             =   11220
      Width           =   1905
   End
   Begin VB.CommandButton Command73 
      Caption         =   "Send List1 to Excel"
      Height          =   465
      Left            =   6810
      TabIndex        =   85
      Top             =   11010
      Width           =   1875
   End
   Begin VB.CommandButton Command72 
      Caption         =   "Load List split spaces"
      Height          =   405
      Left            =   3690
      TabIndex        =   84
      Top             =   6120
      Width           =   1485
   End
   Begin VB.CommandButton Command71 
      Caption         =   "Block Atts"
      Height          =   435
      Left            =   4530
      TabIndex        =   83
      Top             =   11520
      Width           =   2085
   End
   Begin VB.CommandButton Command70 
      Caption         =   "Text with Sequence"
      Height          =   435
      Left            =   4530
      TabIndex        =   82
      Top             =   11070
      Width           =   2085
   End
   Begin VB.CommandButton Command69 
      Caption         =   "Poly Areas to Excel - Seq"
      Height          =   435
      Left            =   4530
      TabIndex        =   81
      Top             =   10170
      Width           =   2085
   End
   Begin VB.CommandButton Command68 
      Height          =   255
      Left            =   4980
      TabIndex        =   80
      Top             =   570
      Width           =   255
   End
   Begin VB.TextBox TextCheck 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6960
      TabIndex        =   79
      Text            =   "CM"
      Top             =   6420
      Width           =   855
   End
   Begin VB.CommandButton Command67 
      Caption         =   "Count Text with String"
      Height          =   345
      Left            =   5250
      TabIndex        =   78
      Top             =   6390
      Width           =   1695
   End
   Begin VB.CommandButton Command66 
      Caption         =   "Check Text"
      Height          =   435
      Left            =   3690
      TabIndex        =   77
      Top             =   5640
      Width           =   1485
   End
   Begin VB.CommandButton Command65 
      Caption         =   "Load List"
      Height          =   405
      Left            =   3690
      TabIndex        =   76
      Top             =   5190
      Width           =   1185
   End
   Begin VB.CommandButton Command64 
      Caption         =   "Numerologist (CharacterJuggle)"
      Height          =   375
      Left            =   10080
      TabIndex        =   75
      Top             =   5910
      Width           =   2445
   End
   Begin VB.CommandButton Command63 
      Caption         =   "revNumerologist (Sequential)"
      Height          =   375
      Left            =   10080
      TabIndex        =   74
      Top             =   5130
      Width           =   2445
   End
   Begin VB.CommandButton Command62 
      Caption         =   "Numerologist (Ascii#)"
      Height          =   375
      Left            =   10080
      TabIndex        =   73
      Top             =   5520
      Width           =   2445
   End
   Begin VB.CommandButton Command61 
      Caption         =   "Numerologist (Sequential)"
      Height          =   375
      Left            =   10080
      TabIndex        =   72
      Top             =   4740
      Width           =   2445
   End
   Begin VB.CommandButton Command60 
      Caption         =   "Text Trim"
      Height          =   375
      Left            =   150
      TabIndex        =   71
      Top             =   9330
      Width           =   1815
   End
   Begin VB.CommandButton Command59 
      Caption         =   "Mtext to Excel"
      Height          =   435
      Left            =   4530
      TabIndex        =   70
      Top             =   10620
      Width           =   2085
   End
   Begin VB.TextBox TextDisperse 
      Height          =   405
      Left            =   9000
      TabIndex        =   69
      Text            =   "Text2"
      Top             =   5760
      Width           =   1035
   End
   Begin VB.CommandButton Command58 
      Caption         =   "Insert Disperse"
      Height          =   435
      Left            =   7230
      TabIndex        =   68
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton Command57 
      Caption         =   "List Blocks"
      Height          =   375
      Left            =   7230
      TabIndex        =   67
      Top             =   5250
      Width           =   1725
   End
   Begin VB.CommandButton Command56 
      Caption         =   "Loop Views"
      Height          =   315
      Left            =   5250
      TabIndex        =   66
      Top             =   5460
      Width           =   1905
   End
   Begin VB.TextBox TextView 
      Height          =   435
      Left            =   7200
      TabIndex        =   65
      Top             =   4770
      Width           =   2325
   End
   Begin VB.CommandButton Command55 
      Caption         =   "SetView"
      Height          =   315
      Left            =   5250
      TabIndex        =   64
      Top             =   5100
      Width           =   1905
   End
   Begin VB.CommandButton Command54 
      Caption         =   "ListViews"
      Height          =   315
      Left            =   5250
      TabIndex        =   63
      Top             =   4770
      Width           =   1905
   End
   Begin VB.CommandButton Command53 
      Caption         =   "Poly Areas to Excel"
      Height          =   435
      Left            =   4530
      TabIndex        =   62
      Top             =   9720
      Width           =   2085
   End
   Begin VB.TextBox TextFillet 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   2100
      TabIndex        =   61
      Text            =   "100"
      Top             =   10530
      Width           =   1605
   End
   Begin VB.CommandButton Command52 
      Caption         =   "Fillet Poly"
      Height          =   435
      Left            =   2100
      TabIndex        =   60
      Top             =   9930
      Width           =   1635
   End
   Begin VB.TextBox TextSpace 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2460
      TabIndex        =   59
      Top             =   4080
      Width           =   945
   End
   Begin VB.CommandButton Command51 
      Caption         =   "Copy IncrementText"
      Height          =   375
      Left            =   150
      TabIndex        =   58
      Top             =   10170
      Width           =   1815
   End
   Begin VB.CommandButton Command50 
      Caption         =   "Scale till cad crash"
      Height          =   525
      Left            =   2160
      TabIndex        =   57
      Top             =   7620
      Width           =   1185
   End
   Begin VB.CommandButton Command49 
      Caption         =   "Font Inf"
      Height          =   375
      Left            =   150
      TabIndex        =   56
      Top             =   10620
      Width           =   1815
   End
   Begin VB.CommandButton Command48 
      Caption         =   "Copy Block"
      Height          =   375
      Left            =   150
      TabIndex        =   55
      Top             =   9780
      Width           =   1815
   End
   Begin VB.CommandButton Command47 
      Caption         =   "Text Justification"
      Height          =   375
      Left            =   150
      TabIndex        =   54
      Top             =   8940
      Width           =   1815
   End
   Begin VB.CommandButton Command46 
      Caption         =   "TextStyle"
      Height          =   375
      Left            =   150
      TabIndex        =   53
      Top             =   8550
      Width           =   1815
   End
   Begin VB.CommandButton Command45 
      Caption         =   "Label Block"
      Height          =   375
      Left            =   150
      TabIndex        =   52
      Top             =   8160
      Width           =   1815
   End
   Begin VB.CommandButton Command18 
      Height          =   135
      Index           =   0
      Left            =   1770
      Picture         =   "MSE_Form1.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "This Form Always On Top"
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command18 
      Height          =   120
      Index           =   1
      Left            =   1770
      Picture         =   "MSE_Form1.frx":0C1D
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "This Form Not On Top"
      Top             =   285
      Width           =   255
   End
   Begin VB.CommandButton Command44 
      Caption         =   "Align Text"
      Height          =   345
      Left            =   2130
      TabIndex        =   49
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton Command43 
      Caption         =   "Hatch Pattern/Scale"
      Height          =   375
      Left            =   150
      TabIndex        =   48
      Top             =   7770
      Width           =   1815
   End
   Begin VB.CommandButton Command42 
      Caption         =   "% Text"
      Height          =   345
      Left            =   2130
      TabIndex        =   47
      Top             =   1020
      Width           =   1185
   End
   Begin VB.CommandButton Command41 
      Caption         =   "Toolbar Lab"
      Height          =   345
      Left            =   2160
      TabIndex        =   46
      Top             =   6750
      Width           =   1185
   End
   Begin VB.CommandButton Command40 
      Caption         =   "List Toolbar Macro"
      Height          =   345
      Left            =   2160
      TabIndex        =   45
      Top             =   6390
      Width           =   1485
   End
   Begin VB.CommandButton Command39 
      Caption         =   "List Menus"
      Height          =   345
      Left            =   2160
      TabIndex        =   44
      Top             =   6030
      Width           =   1185
   End
   Begin VB.CommandButton Command38 
      Caption         =   "List Menu"
      Height          =   345
      Left            =   2160
      TabIndex        =   43
      Top             =   5670
      Width           =   1185
   End
   Begin VB.CommandButton Command37 
      Caption         =   "Object ID"
      Height          =   375
      Left            =   150
      TabIndex        =   42
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton Command36 
      Caption         =   "Elevator"
      Height          =   345
      Left            =   2130
      TabIndex        =   41
      Top             =   9120
      Width           =   1095
   End
   Begin VB.CommandButton Command35 
      Caption         =   "ParseTest"
      Height          =   345
      Left            =   2160
      TabIndex        =   40
      Top             =   7260
      Width           =   1185
   End
   Begin VB.CommandButton Command34 
      Caption         =   "Layer Freezer"
      Height          =   375
      Left            =   150
      TabIndex        =   39
      Top             =   7350
      Width           =   1815
   End
   Begin VB.CommandButton Command33 
      Caption         =   "Erase Poly"
      Height          =   375
      Left            =   150
      TabIndex        =   38
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton Command32 
      Caption         =   "Replace"
      Height          =   405
      Left            =   3480
      TabIndex        =   37
      Top             =   960
      Width           =   1185
   End
   Begin VB.TextBox Textrep 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3480
      TabIndex        =   36
      Text            =   "_AREA"
      Top             =   660
      Width           =   1455
   End
   Begin VB.CommandButton Command31 
      Caption         =   "Round Text"
      Height          =   345
      Left            =   2130
      TabIndex        =   35
      Top             =   660
      Width           =   1185
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Rotate"
      Height          =   345
      Left            =   2130
      TabIndex        =   34
      Top             =   9480
      Width           =   1125
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Text Sum"
      Height          =   435
      Left            =   2160
      TabIndex        =   33
      Top             =   8670
      Width           =   1485
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Parcel Numberer"
      Height          =   375
      Left            =   150
      TabIndex        =   32
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton Command27 
      Caption         =   "AlignSpace Text"
      Height          =   375
      Left            =   2130
      TabIndex        =   31
      Top             =   3330
      Width           =   1575
   End
   Begin VB.CommandButton Command26 
      Caption         =   "AlignSpace Block"
      Height          =   375
      Left            =   2160
      TabIndex        =   30
      Top             =   4410
      Width           =   1575
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Align Block"
      Height          =   375
      Left            =   2160
      TabIndex        =   29
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton Command24 
      Caption         =   "SortPolys"
      Height          =   375
      Left            =   150
      TabIndex        =   28
      Top             =   2070
      Width           =   1815
   End
   Begin VB.CommandButton Command23 
      Caption         =   "List1 to Text"
      Height          =   465
      Left            =   3330
      TabIndex        =   27
      Top             =   9270
      Width           =   1095
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Layer and Color"
      Height          =   465
      Left            =   2160
      TabIndex        =   26
      Top             =   8190
      Width           =   1485
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Hatch2Back"
      Height          =   405
      Left            =   2160
      TabIndex        =   25
      Top             =   5250
      Width           =   1125
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Copy to Front"
      Height          =   345
      Left            =   2130
      TabIndex        =   24
      Top             =   2100
      Width           =   1575
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Erase Block"
      Height          =   375
      Left            =   150
      TabIndex        =   23
      Top             =   3690
      Width           =   1815
   End
   Begin VB.CommandButton Command17 
      Caption         =   "PolyWidth Erase"
      Height          =   375
      Left            =   150
      TabIndex        =   22
      Top             =   5670
      Width           =   1815
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Erase Solid"
      Height          =   375
      Left            =   150
      TabIndex        =   21
      Top             =   5250
      Width           =   1815
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Erase Hatch"
      Height          =   375
      Left            =   150
      TabIndex        =   20
      Top             =   4080
      Width           =   1815
   End
   Begin VB.TextBox TextDirectory 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5250
      TabIndex        =   19
      Text            =   "Directory"
      Top             =   60
      Width           =   7515
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Directory"
      Height          =   345
      Left            =   2130
      TabIndex        =   18
      Top             =   1740
      Width           =   1575
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Erase Lines"
      Height          =   375
      Left            =   150
      TabIndex        =   17
      Top             =   6060
      Width           =   1815
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Send Command"
      Height          =   345
      Left            =   2130
      TabIndex        =   16
      Top             =   2460
      Width           =   1575
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Offset"
      Height          =   345
      Left            =   2130
      TabIndex        =   15
      Top             =   2820
      Width           =   1575
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Erase Layer"
      Height          =   375
      Left            =   150
      TabIndex        =   14
      Top             =   4860
      Width           =   1815
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Erase Text"
      Height          =   375
      Left            =   150
      TabIndex        =   13
      Top             =   4470
      Width           =   1815
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Add List2Text"
      Height          =   525
      Left            =   2130
      TabIndex        =   12
      Top             =   120
      Width           =   1905
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2835
      Left            =   5250
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   6720
      Width           =   7515
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Text Inventory"
      Height          =   375
      Left            =   150
      TabIndex        =   10
      Top             =   2850
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Poly Ready for Code"
      Height          =   375
      Left            =   150
      TabIndex        =   9
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "PolyClose"
      Height          =   375
      Left            =   150
      TabIndex        =   8
      Top             =   1290
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Change Text"
      Height          =   375
      Left            =   150
      TabIndex        =   7
      Top             =   2460
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SortPolys"
      Height          =   375
      Left            =   150
      TabIndex        =   6
      Top             =   900
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "All Ents"
      Height          =   375
      Left            =   150
      TabIndex        =   5
      Top             =   510
      Width           =   1815
   End
   Begin VB.ListBox ListInsert 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Index           =   2
      Left            =   10470
      TabIndex        =   4
      Top             =   9720
      Width           =   1815
   End
   Begin VB.ListBox ListInsert 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Index           =   1
      Left            =   8640
      TabIndex        =   3
      Top             =   9720
      Width           =   1815
   End
   Begin VB.ListBox ListInsert 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Index           =   0
      Left            =   6810
      TabIndex        =   2
      Top             =   9720
      Width           =   1815
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4155
      Left            =   5250
      TabIndex        =   1
      Top             =   570
      Width           =   7545
   End
   Begin VB.CommandButton Command1 
      Caption         =   "sset"
      Height          =   375
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Various AutoCAD VB Functions for use while designing in AutoCAD

Public varBefore As Integer
Public varBefore2 As Integer


Private Sub Command1_Click()
Dim AutoCAD As acadapplication
Dim Thisdrawing As AcadDocument  'as Object
 Set acadapp = GetObject(, "autocad.application")
  AutoCAD.activedocument.WindowTitle
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))

Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    If ent.EntityType = acText Then
    'set ent = autocad
    x = ent.insertionPoint
    'If ent.TEXTSTRING = " " Then
    List1.AddItem ent.entityname & " " & ent.TextString
    
    ListInsert(0).AddItem x(0)
    ListInsert(1).AddItem x(1)
    ListInsert(2).AddItem x(2)
    
    Dim x2(0 To 2) As Double
    Dim Y(0 To 2) As Double
    Dim y2(0 To 2) As Double
    dblMove = 2
    y2(0) = x(0) '+ dblMove
     y2(1) = x(0) + dblMove
      'x2(0) = x(0) '+ dblMove
    
    x2(0) = CDbl(x(0))
    x2(1) = CDbl(x(1))
    Y(0) = CDbl(y2(0))
    Y(1) = CDbl(y2(1))
        
    
    
    ' Move the circle
   ent.Move x2, Y
    
    
    
    'ent.move (x,y)
    'ent.insertionpoint(1) = 0
    'ent.TextString = "111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111"
   End If
   ' End If
    Next ent
End Sub



Private Sub Command10_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))



Apptivate
On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    If ent.Layer = "p-lotlines" Then
    'set ent = autocad
    ent.Erase
    End If
    
    
        
    
   
   
    Next ent
    SSET2.Delete
End Sub

Private Sub Command11_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))



Apptivate
On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    ''If ent.EntityType = acHatch Then
    'set ent = autocad
    'ent.Erase
    ''End If
    ent.Offset (5)
    'ent.Layer = "p-sidewalk"
    ''If ent.EntityType = acCircle Then
    'set ent = autocad
    'ent.Erase
    ''End If
    
    Dim ssetObj As AcadSelectionSet
Dim ssetObj2 As AcadSelectionSet
    Set ssetObj = Thisdrawing.SelectionSets.Add("SSETm")
    'Set ssetObj2 = Thisdrawing.SelectionSets.Add("SSET2l")
 
 Mode = acSelectionSetLast
 
    ssetObj.Select Mode
    'For Each ent2 In ssetObj
    'ssetObj2.AddItems ent
    ''MsgBox ssetObj.
    'ssetObj.Layer = strLayer
ent2.Layer = "p-sidewalk"
'Next


ssetObj.Delete
    
   
   
    Next ent
    SSET2.Delete
End Sub

Private Sub Command12_Click()
Dim AutoCAD As acadapplication
Dim Thisdrawing As AcadDocument  'as Object
 Set acadapp = GetObject(, "autocad.application")
    
Set Thisdrawing = acadapp.activedocument
Dim activedocument As Object

Apptivate
CmdGetSetVar_Click
Thisdrawing.SendCommand "vbaload" & vbCr & "p" & vbCr & "P" & vbCr & "la" & vbCr & "p-sidewalk" & vbCr
Thisdrawing.SendCommand "p" & vbCr

CmdResetVar_Click

End Sub

Private Sub Command13_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))



Apptivate
On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    If ent.EntityType = acLine Then
    'set ent = autocad
    ent.Erase
    End If
    
    
        
    
   
   
    Next ent
    SSET2.Delete
End Sub

Private Sub Command14_Click()
FmTVremote.Show
End Sub

Private Sub Command15_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))



Apptivate
On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    If ent.EntityType = acHatch Then
    'set ent = autocad
    ent.Erase
    End If
    
    
        
    
   
   
    Next ent
    SSET2.Delete
End Sub

Private Sub Command16_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))



Apptivate
On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    If ent.EntityType = acSolid Then
    'set ent = autocad
    ent.Erase
    End If
    
    
        
    
   
   
    Next ent
    SSET2.Delete
End Sub

Private Sub Command17_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))
Dim StartWidth As Double
    Dim EndWidth As Double

Apptivate
On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    If ent.EntityType = acPolylineLight Then
    Dim lwpl As AutoCAD.AcadLWPolyline
            
            If (TypeOf ent Is AutoCAD.AcadLWPolyline) Then
            
                Set lwpl = ent
                
        Index = 1
        ent.GetWidth Index, StartWidth, EndWidth
                    
                                      

            
        
    
    
    
    If StartWidth = 10 Then
    ent.Erase
    End If
    End If
    
    End If
    
    
        
    
   
   
    Next ent
    SSET2.Delete
End Sub

Private Sub Command18back_Click()
On Error Resume Next
List2.Clear
Dim AutoCAD As acadapplication
Dim Thisdrawing As AcadDocument
 Set acadapp = GetObject(, "autocad.application")
   Apptivate
Set Thisdrawing = acadapp.activedocument
Dim activedocument As Object
Dim acadapplication As Object
Dim modelspace As Object
Dim ent 'As AutoCAD.AcadEntity
Dim vtxs As Variant
Dim idx As Long
Dim SSET2 As Object
   
List1(0).Clear
List2.Clear
Me.Hide

Set SSET2 = Thisdrawing.SelectionSets.Add("strSet2")
  
    acadapp.Visible = True
    SSET2.SelectOnScreen
   
        For Each ent In SSET2
        If Not TypeOf ent Is AutoCAD.AcadLWPolyline Then
        ''MsgBox "Select a LightWeightPolyline."
        End If
        Next
        For Each ent In SSET2
        Dim lwpl As AutoCAD.AcadLWPolyline
            
            If (TypeOf ent Is AutoCAD.AcadLWPolyline) Then
            
                Set lwpl = ent
                
               
                    ' Loop through the exploded objects
                    Dim i As Integer
                    

            End If
        
            Next ent
    SSET2.Delete
    List2.AddItem dist1
    Me.Show
     
Exit Sub

hu2:
MsgBox Err.Description
Me.Show
End Sub

Private Sub Command18_Click(Index As Integer)
On Error Resume Next
 If Command18(0) Then AlwaysOnTop Me, True
 If Command18(1) Then AlwaysOnTop Me, False
End Sub
Private Sub AlwaysOnTop(FrmID As Form, OnTop As Integer)
    Const SWP_NOMOVE = 2
    Const SWP_NOSIZE = 1
    Const Flags = SWP_NOMOVE Or SWP_NOSIZE
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    On Error Resume Next
    If OnTop = -1 Then
        OnTop = SetWindowPos(FrmID.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
    Else
        OnTop = SetWindowPos(FrmID.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
    End If
End Sub




Private Sub Command19_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))



Apptivate
On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    If ent.EntityType = acBlockReference Then
    If ent.Name <> "RS2" And ent.Name <> "RS" Then
    
    'set ent = autocad
    ent.Erase
    End If
    End If
    
    
        
    
   
   
    Next ent
    SSET2.Delete
End Sub

Private Sub Command2_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch




Apptivate
On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add("strSet")

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    List1.AddItem Val(ent.TextString)
    'list2.AddItem ent.TextString
    ''tot = tot + Val(ent.TextString)
    
    ''If ent.TextString = "91 Acres" Then
    Text1.text = Text1.text & ent.TextString & vbCrLf
    ''End If
     If ent.TextString = "1QQ" Then
     ent.HIGHLIGTH = True
     End If
    
   
   
    Next ent
    'Text1.text = Str(tot)
    SSET2.Delete
End Sub

Private Sub Command20_Click()

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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))



Apptivate
'On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    If ent.EntityType <> acHatch Then
    'Thisdrawing.SendCommand "DRAWORDER" & vbCr & "p" & vbCr & vbCr & vbCr
    ent.Copy
    ent.Erase
    End If
    ''If ent.EntityType = acText Then
    'Thisdrawing.SendCommand "DRAWORDER" & vbCr & "p" & vbCr & vbCr & vbCr
    ''ent.Copy
    ''ent.Erase
    ''End If
    ''If ent.EntityType = acBlockReference Then
    ''ent.Copy
    ''ent.Erase
    ''End If
    
    
        
    
   
   
    Next ent
    SSET2.Delete


End Sub

Private Sub Command21_Click()

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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))



Apptivate
'On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    If ent.EntityType = acHatch Then
    Thisdrawing.SendCommand "DRAWORDER" & vbCr & "p" & vbCr & vbCr & vbCr
    Thisdrawing.SendCommand "DRAWORDER" & vbCr & "p" & vbCr & vbCr & vbCr
    Thisdrawing.SendCommand "DRAWORDER" & vbCr & "p" & vbCr & vbCr & vbCr
    
    
    End If
    ''If ent.EntityType = acText Then
    'Thisdrawing.SendCommand "DRAWORDER" & vbCr & "p" & vbCr & vbCr & vbCr
    ''ent.Copy
    ''ent.Erase
    ''End If
    ''If ent.EntityType = acBlockReference Then
    ''ent.Copy
    ''ent.Erase
    ''End If
    
    
        
    
   
   
    Next ent
    SSET2.Delete

End Sub

Private Sub Command22_Click()
Dim Thisdrawing As Object
Dim obstyls
Dim obstyl
On Error Resume Next

Set acadapp = GetObject(, "autocad.application")
Set Thisdrawing = acadapp.activedocument
Set obstyls = Thisdrawing.Layers
For Each obstyl In obstyls
lcolor = obstyl.Color
List1.text = lcolor
    
    If List1.ListIndex = -1 Then
    List1.AddItem lcolor
    End If
Text1.text = Text1.text & obstyl.Color & vbCrLf

Next

Exit Sub
zerp:
MsgBox Err.Description
Exit Sub
End Sub
Private Sub AddUnique(StringToAdd As String, lst As ListBox)
    List.text = StringToAdd


    If List.ListIndex = -1 Then
        'it does not exist, so add it..
        List.AddItem StringToAdd
    End If
End Sub
Private Sub Command23_Click()
For x = 0 To List1.ListCount - 1
Text1.text = Text1.text & List1.List(x) & vbCrLf
Next x
End Sub

Private Sub Command24_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))

Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    If (TypeOf ent Is AutoCAD.AcadLWPolyline) Then
    If ent.Closed = True Then
    'If ent.Layer = "P-HYDRO" Then
    ent.Layer = "P-LANDSCAPE"
    'End If
    End If
    
   End If
   ' End If
    Next ent
End Sub

Private Sub Command25_Click()
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
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    If ent.EntityType = acBlockReference Then
    cnt = cnt + 1
    exerx1 = ent.insertionPoint
    If cnt = 1 Then
    exerX2 = ent.insertionPoint
    End If
    movex(0) = exerX2(0)
    movex(1) = exerx1(1)
    
    ent.insertionPoint = movex
    
     'MsgBox SSET2.Item()
    End If
    
   
    Next ent
    SSET2.Delete
End Sub

Private Sub Command26_Click()
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
 Dim movex(0 To 2) As Double
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))

Dim minPoint As Variant
Dim maxPoint As Variant

Apptivate
'On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    If ent.EntityType = acBlockReference Then
    
    
    'do offset
    
    
    exerx1 = ent.insertionPoint
    If cnt = 0 Then
    ent.GetBoundingBox minPoint, maxPoint
    MoveY = maxPoint(1) - minPoint(1) + Val(TextSpace.text)
    mY = MoveY
    exerX2 = ent.insertionPoint
    
    End If
    movex(0) = exerX2(0)
    movex(1) = exerX2(1) - (MoveY * cnt)
    If cnt > 0 Then
    ent.insertionPoint = movex
    
     'MsgBox SSET2.Item()
     End If
     cnt = cnt + 1
    'MoveY = MoveY + MoveY
    mY = mY + mY
    List1.AddItem MoveY
    End If
    
    
        
    
   
   
    Next ent
    SSET2.Delete
End Sub

Private Sub Command27_Click()
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
 Dim movex(0 To 2) As Double
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))

Dim minPoint As Variant
Dim maxPoint As Variant

Apptivate
'On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    'do offset
    
    
    exerx1 = ent.insertionPoint
    If cnt = 0 Then
    ent.GetBoundingBox minPoint, maxPoint
    MoveY = maxPoint(1) - minPoint(1) + Val(TextSpace.text)
    mY = MoveY
    exerX2 = ent.insertionPoint
    
    End If
    movex(0) = exerX2(0)
    movex(1) = exerX2(1) - (MoveY * cnt)
    If cnt > 0 Then
    ent.insertionPoint = movex
    
     'MsgBox SSET2.Item()
     End If
     cnt = cnt + 1
    'MoveY = MoveY + MoveY
    mY = mY + mY
    List1.AddItem MoveY
    
    Next ent
    SSET2.Delete
End Sub

Private Sub Command28_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch




Apptivate
On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add("strSet")

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    nHood = "2"
    For Each ent In SSET2
    cntx = cntx + 1
    If cntx < 27 Then
    ent.TextString = nHood & Chr(64 + cntx)
    
    List1.AddItem Val(ent.TextString) & " " & cntx
    'list2.AddItem ent.TextString
    ''tot = tot + Val(ent.TextString)
    
    ''If ent.TextString = "91 Acres" Then
    Text1.text = Text1.text & ent.TextString & vbCrLf
    End If
    ''End If
     If cntx >= 27 Then
    ent.TextString = nHood & Chr(64 + (cntx - 26)) & Chr(64 + (cntx - 26))
    
    List1.AddItem Val(ent.TextString) & " " & cntx
    'list2.AddItem ent.TextString
    ''tot = tot + Val(ent.TextString)
    
    ''If ent.TextString = "91 Acres" Then
    Text1.text = Text1.text & ent.TextString & vbCrLf
    End If
    
   
   
    Next ent
    'Text1.text = Str(tot)
    SSET2.Delete
End Sub

Private Sub Command29_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch




Apptivate
On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add("strSet")

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    tval = tval + Val(ent.TextString)
    List1.AddItem Val(ent.TextString)
    'list2.AddItem ent.TextString
    ''tot = tot + Val(ent.TextString)
    
    ''If ent.TextString = "91 Acres" Then
    'Text1.text = Text1.text & ent.TextString & vbCrLf
    ''End If
     'If ent.TextString = "1QQ" Then
     'ent.HIGHLIGTH = True
     'End If
    
   
   
    Next ent
    Text1.text = tval
    'Text1.text = Str(tot)
    SSET2.Delete
End Sub

Private Sub Command3_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))



Apptivate
'On Error Resume Next

    For Each ent In Thisdrawing.paperspace
    
    If ent.EntityType = acText Then
    'set ent = autocad
    x = ent.insertionPoint
    'If ent.TEXTSTRING = " " Then
    List1.AddItem ent.entityname & " " & ent.TextString
    
    ListInsert(0).AddItem x(0)
    ListInsert(1).AddItem x(1)
    ListInsert(2).AddItem x(2)
    
    Dim x2(0 To 2) As Double
    Dim Y(0 To 2) As Double
    Dim y2(0 To 2) As Double
    dblMove = 2
    y2(0) = x(0) '+ dblMove
     y2(1) = x(0) + dblMove
      'x2(0) = x(0) '+ dblMove
    
    x2(0) = CDbl(x(0))
    x2(1) = CDbl(x(1))
    Y(0) = CDbl(y2(0))
    Y(1) = CDbl(y2(1))
        
    
    
    ' Move the circle
   'ent.Move x2, y
    
    
    
    'ent.move (x,y)
    'ent.insertionpoint(1) = 0
    'ent.TextString = "111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111"
   End If
   ' End If
    Next ent
End Sub

Private Sub Command30_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))



Apptivate
On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    If ent.EntityType <> acNothing Then
    
    ent.Rotate
    'End If
    End If
    
    
        
    
   
   
    Next ent
    SSET2.Delete
End Sub

Private Sub Command31_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch




Apptivate
On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add("strSet")

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    List1.AddItem Val(ent.TextString)
    'list2.AddItem ent.TextString
    ''tot = tot + Val(ent.TextString)
    
    ''If ent.TextString = "91 Acres" Then
    Text1.text = Text1.text & ent.TextString & vbCrLf
    ''End If
    textval = Val(ent.TextString)
    newtext = Round(textval, 2)
     
     ent.TextString = Trim(newtext)
     
    
   
   
    Next ent
    'Text1.text = Str(tot)
    SSET2.Delete
End Sub

Private Sub Command32_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch




Apptivate
On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add("strSet")

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    List1.AddItem Val(ent.TextString)
    'list2.AddItem ent.TextString
    ''tot = tot + Val(ent.TextString)
    
    ''If ent.TextString = "91 Acres" Then
    Text1.text = Text1.text & ent.TextString & vbCrLf
    ''End If
    textval = Replace(ent.TextString, Textrep.text, "")
    newtext = textval
     
     ent.TextString = Trim(newtext)
     
    
   
   
    Next ent
    'Text1.text = Str(tot)
    SSET2.Delete
End Sub

Private Sub Command33_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))



Apptivate
On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    If ent.EntityType = acPolylineLight Then
    'set ent = autocad
    ent.Erase
    End If
    
    
        
    
   
   
    Next ent
    SSET2.Delete
End Sub

Private Sub Command34_Click()
'layers for 14
Dim acadapp As Object
Dim AutoCAD As Object
Dim activedocument As Object
Dim Thisdrawing As Object
Dim objlayers As AcadLayers
Dim objlayer As AcadLayer
Dim objlayer2 As String
Dim objlayers2
Dim strlayernames As String
Dim strlayercols As String
Dim strlayerVIs As String
Dim objlayerstr

On Error Resume Next
Set acadapp = GetObject(, "AutoCAD.Application")
Dim strnamelayers As String

Set Thisdrawing = acadapp.activedocument
Set objlayers2 = Thisdrawing.Layers
ooper = Layer7
strnamelayers = "the l" & vbCrLf
Set objLayerstrx = Thisdrawing.ActiveLayer
TextBefore = ""
TextBefore.text = objLayerstrx.Name 'Thisdrawing.ACTIVELAYER
Thisdrawing.ActiveLayer = Thisdrawing.Layers("0")

For Each objlayerstr In objlayers2
On Error Resume Next
If InStr(1, objlayerstr.Name, "_AREA") > 0 Then
objlayerstr.Freeze = True
End If
List1.AddItem objlayerstr.Name
Next objlayerstr



End Sub

Private Sub Command35_Click()
strRet = Replace("\\infinity\i3ms\documentname\.dot filename", "\\", "e:\")
List1.AddItem Trim(strRet)
End Sub

Private Sub Command36_Click()
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
     newValue = Val(ent.Elevation) * 0.3047851
     
     ent.Elevation = newValue
    Next ent
    
    SSET2.Delete
End Sub

Private Sub Command37_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))



Apptivate
On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    List1.AddItem "Handle: " & ent.Handle & " ObjectID: " & ent.ObjectID & " OwnerID: " & ent.OwnerID
    
        
    
   
   
    Next ent
    SSET2.Delete
End Sub

Private Sub Command38_Click()
Dim Thisdrawing As Object
Dim objToolBars
Dim objToolBar
'On Error Resume Next

Set acadapp = GetObject(, "autocad.application")
Set Thisdrawing = acadapp.activedocument
'Set objToolBars = Thisdrawing.MenuGroups

Dim currMenuGroup As AcadMenuGroup
    Set currMenuGroup = Thisdrawing.application.MenuGroups.Item(1)
    
    ' Create the new toolbar
    Dim newToolBar As AcadToolbar
    'Set newToolBar = currMenuGroup.Toolbars.Add("TestToolbar")
'Set newToolBar = currMenuGroup.Toolbars.Count

For Y = 0 To currMenuGroup.Menus.Count - 1
'Y = 0
xx = currMenuGroup.Menus.Item(Y).Name
'xx = currMenuGroup.Toolbars.Item(2).Name
'For Each newToolBar In currMenuGroup


List1.AddItem xx
Next
x = x + 1

'Next
Exit Sub
zerp:
MsgBox Err.Description
Exit Sub
End Sub

Private Sub Command39_Click()
Dim Thisdrawing As Object
Dim objToolBars
Dim objToolBar
'On Error Resume Next

Set acadapp = GetObject(, "autocad.application")
Set Thisdrawing = acadapp.activedocument
'Set objToolBars = Thisdrawing.MenuGroups

Dim currMenuGroup As AcadMenuGroup
    Set currMenuGroup = Thisdrawing.application.MenuGroups.Item(1)
    
    ' Create the new toolbar
    Dim newToolBar As AcadToolbar
    'Set newToolBar = currMenuGroup.Toolbars.Add("TestToolbar")
'Set newToolBar = currMenuGroup.Toolbars.Count

'For Y = 0 To currMenuGroup.Toolbars.Count - 1
For Y = 0 To Thisdrawing.application.MenuGroups.Count - 1
'Y = 0
''xx = currMenuGroup.Menus.Item(Y).Name
'yy = currMenuGroup.Toolbars.Item(Y).Name
 yy = Thisdrawing.application.MenuGroups.Item(Y).Name

'For Each newToolBar In currMenuGroup


List1.AddItem yy

Next
x = x + 1

'Next
Exit Sub
zerp:
MsgBox Err.Description
Exit Sub
End Sub

Private Sub Command4_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))

Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    If (TypeOf ent Is AutoCAD.AcadLWPolyline) Then
    If ent.Closed = False Then
    If ent.Layer = "P-HYDRO" Then
    ent.Layer = "P-LEGAL2"
    End If
    End If
    
   End If
   ' End If
    Next ent
End Sub

Private Sub Command40_Click()
Dim Thisdrawing As Object
Dim objToolBars
Dim objToolBar
'On Error Resume Next

Set acadapp = GetObject(, "autocad.application")
Set Thisdrawing = acadapp.activedocument
'Set objToolBars = Thisdrawing.MenuGroups

Dim currMenuGroup As AcadMenuGroup
    Set currMenuGroup = Thisdrawing.application.MenuGroups.Item(1)
    
    ' Create the new toolbar
    Dim newToolBar As AcadToolbar
    'Set newToolBar = currMenuGroup.Toolbars.Add("TestToolbar")
'Set newToolBar = currMenuGroup.Toolbars.Count

'For Y = 0 To currMenuGroup.Toolbars.Count - 1
For Y = 0 To currMenuGroup.Menus.Count - 1
'Y = 0
''xx = currMenuGroup.Menus.Item(Y).Name
'xx = currMenuGroup.Toolbars.Item(Y).Name
yy = currMenuGroup.Toolbars.Item(Y).Name
'Set tb = currMenuGroup.Toolbars.Item(Y).Name
'zz = tb.Macro
'For Each newToolBar In currMenuGroup
''yy = currMenuGroup.Toolbars.Item("Layouts").Item(0).Macro

List1.AddItem yy
'List1.AddItem zz

Next
x = x + 1

'Next
Exit Sub
zerp:
MsgBox Err.Description
Exit Sub
End Sub

Private Sub Command41_Click()
Dim Thisdrawing As Object
Dim objToolBars
Dim objToolBar
'On Error Resume Next

Set acadapp = GetObject(, "autocad.application")
Set Thisdrawing = acadapp.activedocument
'Set objToolBars = Thisdrawing.MenuGroups

Dim currMenuGroup As AcadMenuGroup
    Set currMenuGroup = Thisdrawing.application.MenuGroups.Item(1)
    
    ' Create the new toolbar
    Dim newToolBar As AcadToolbar
    'Set newToolBar = currMenuGroup.Toolbars.Add("TestToolbar")
'Set newToolBar = currMenuGroup.Toolbars.Count

'For Y = 0 To currMenuGroup.Toolbars.Count - 1
For Y = 0 To currMenuGroup.Toolbars.Item("text").Count - 1
'Y = 0
''xx = currMenuGroup.Menus.Item(Y).Name
'xx = currMenuGroup.Toolbars.Item(Y).Name
''yy = currMenuGroup.Toolbars.Item(Y).Name
'Set tb = currMenuGroup.Toolbars.Item(Y).Name
'zz = tb.Macro
'For Each newToolBar In currMenuGroup
'yy = currMenuGroup.Toolbars.Item("text").Item(Y).Macro

currMenuGroup.Toolbars.Item("text").Top = 100

List1.AddItem yy
'List1.AddItem zz

Next
x = x + 1

'Next
Exit Sub
zerp:
MsgBox Err.Description
Exit Sub
End Sub

Private Sub Command42_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch




Apptivate
On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add("strSet")

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    List1.AddItem Val(ent.TextString)
    'list2.AddItem ent.TextString
    ''tot = tot + Val(ent.TextString)
    
    ''If ent.TextString = "91 Acres" Then
    Text1.text = Text1.text & ent.TextString & vbCrLf
    ''End If
    ipos = InStr(1, ent.TextString, ".")
    Textx = Mid(ent.TextString, ipos + 1)
    
    newtext = Trim(Textx & "%")
     
     ent.TextString = Trim(newtext)
     
    
   
   
    Next ent
    'Text1.text = Str(tot)
    SSET2.Delete
End Sub

Private Sub Command43_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))



Apptivate
'On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    If ent.EntityType = acHatch Then
    'set ent = autocad
    'ent.PatternType = 2
     'ent.PatternName = "SOLID"
     MsgBox ent.PatternName
    End If
    
    
        
    
   
   
    Next ent
    
    SSET2.Delete
End Sub

Private Sub Command44_Click()
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
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    If ent.EntityType = acText Then
    cnt = cnt + 1
    exerx1 = ent.insertionPoint
    If cnt = 1 Then
    exerX2 = ent.insertionPoint
    End If
    movex(0) = exerX2(0)
    movex(1) = exerx1(1)
    
    ent.insertionPoint = movex
    
     'MsgBox SSET2.Item()
    End If
    
    
        
    
   
   
    Next ent
    SSET2.Delete
End Sub

Private Sub Command45_Click()
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
Dim textObj As AcadText
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))

h = 100

Apptivate
'On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    If ent.EntityType = acBlockReference Then
    TextString = ent.Name
    cnt = cnt + 1
    exerx1 = ent.insertionPoint
    exerx1(0) = exerx1(0) + (h * 5)
    Set textObj = Thisdrawing.modelspace.AddText(TextString, exerx1, h)
    End If
    textObj.StyleName = "SIMPLEX"
   
    Next ent
    SSET2.Delete
End Sub

Private Sub Command46_Click()
Dim AutoCAD As acadapplication
Dim Thisdrawing As AcadDocument  'as Object
Dim objText As AcadText
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))


'objText
Apptivate
On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    If ent.EntityType = acText Then
    'set ent = autocad
    ent.StyleName = "SIMPLEX"
    ent.ScaleFactor = 1
    ent.Height = 1.5
    ent.ObliqueAngle = 0
    End If
    Next ent
    SSET2.Delete
End Sub

Private Sub Command47_Click()

Dim AutoCAD As acadapplication
Dim Thisdrawing As AcadDocument  'as Object
Dim objText As AcadText
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
 'Dim insertionPoint(0 To 2) As Variant
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))


'objText
Apptivate
On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    If ent.EntityType = acText Then
    insertionPoint = ent.insertionPoint
    'set ent = autocad
    ent.StyleName = "SIMPLEX"
    ent.ScaleFactor = 1
    ent.Height = 1.5
    ent.Alignment = acAlignmentRight
    ent.TextAlignmentPoint = insertionPoint

    End If
    
    
        
    
   
   
    Next ent
    SSET2.Delete

'acAlignmentLeft
'acAlignmentCenter
'acAlignmentRight
'acAlignmentAligned
'acAlignmentMiddle
'acAlignmentFit
'acAlignmentTopLeft
'acAlignmentTopCenter
'acAlignmentTopRight
'acAlignmentMiddleLeft
'acAlignmentMiddleCenter
'acAlignmentMiddleRight
'acAlignmentBottomLeft
'acAlignmentBottomCenter
'acAlignmentBottomRight

End Sub

Private Sub Command48_Click()
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
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    If ent.EntityType = acBlockReference Then
    cnt = cnt + 1
    exerx1 = ent.insertionPoint
    exerX2 = exerx1
    
    For v = 0 To 4
    exerX2(0) = exerx1(0) + (10 * v + 1)
    exerX2(1) = exerx1(1) + (10 * v + 1)
    
   
   ent.Copy
   ent.Move exerx1, exerX2
    Next
     'MsgBox SSET2.Item()
    End If
    
   
    Next ent
    SSET2.Delete
End Sub

Private Sub Command49_Click()
Dim AutoCAD As acadapplication
Dim Thisdrawing As AcadDocument  'as Object
Dim objText As AcadText
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
Dim typeFace As String
    Dim Bold As Boolean
    Dim Italic As Boolean
    Dim charSet As Long
    Dim PitchandFamily As Long
    
    'Thisdrawing.ActiveTextStyle.GetFont typeFace, Bold, Italic, charSet, PitchandFamily
    
    

ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))


'objText
Apptivate
On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    If ent.EntityType = acText Then
    'set ent = autocad
    'ent.StyleName = "SIMPLEX"
    'ent.ScaleFactor = 1
    'ent.Height = 1.5
    'ent.ObliqueAngle = 0
    ent.GetFont typeFace, Bold, Italic, charSet, PitchandFamily
    
    MsgBox "The current text style has the following font properties:" & vbCrLf _
            & "Typeface: " & typeFace & vbCrLf _
            & "Bold: " & Bold & vbCrLf _
            & "Italic: " & Italic & vbCrLf _
            & "Character set: " & charSet & vbCrLf _
            & "Pitch and Family: " & PitchandFamily

    
    End If
    Next ent
    SSET2.Delete
End Sub
Sub Example_GetFont()
    ' This example find the font information for the active text style.
    
    Dim typeFace As String
    Dim Bold As Boolean
    Dim Italic As Boolean
    Dim charSet As Long
    Dim PitchandFamily As Long
    
    Thisdrawing.ActiveTextStyle.GetFont typeFace, Bold, Italic, charSet, PitchandFamily
    
    MsgBox "The current text style has the following font properties:" & vbCrLf _
            & "Typeface: " & typeFace & vbCrLf _
            & "Bold: " & Bold & vbCrLf _
            & "Italic: " & Italic & vbCrLf _
            & "Character set: " & charSet & vbCrLf _
            & "Pitch and Family: " & PitchandFamily
    
End Sub

Private Sub Command5_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))



Apptivate
'On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    If ent.EntityType = acText Then
    'set ent = autocad
    If ent.TextString = "L" Then
    ent.TextString = "LOW"
    End If
    If ent.TextString = "M" Then
    ent.TextString = "MED."
    End If
    If ent.TextString = "MH" Then
    ent.TextString = "MED. HIGH"
    End If
    
    
        
    
   End If
   
    Next ent
    SSET2.Delete
End Sub

Private Sub Command50_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))



Apptivate
'On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    If ent.EntityType = acText Then
    
    BasePointX = ent.insertionPoint
    For x = 0 To 100000
    ent.ScaleEntity BasePointX, 1000
    List1.AddItem ent.entityname & x & ent.Visible
    acadapp.ZoomExtents
    Next
    End If
    
    
        
    
   
   
    Next ent
    SSET2.Delete
End Sub

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
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    If ent.EntityType = acText Then
    cnt = cnt + 1
    exerx1 = ent.insertionPoint
    'exerX2 = exerx1
    newtext = ent.TextString
    For v = 0 To 100
    'exerX2(0) = exerx1(0) + (10 * v + 1)
    'exerX2(1) = exerx1(1) + (10 * v + 1)
    'insertionPoint2(0) = insertionPoint2(0) - insertionPoint2(0)
    'insertionPoint2(1) = insertionPoint2(1) - insertionPoint2(1)
    'insertionPoint2(0) = 0
    
   
   
   With Thisdrawing.Utility
    .InitializeUserInput 1
    
    insertionPoint2 = Thisdrawing.Utility.GetPoint(, vbCr & "insertion: ")
   End With
   'ent.Move exerx1, insertionPoint2
movex(0) = insertionPoint2(0)
movex(1) = insertionPoint2(1)
   
    ent.Copy
   
   'ent.insertionPoint = insertionPoint2
   
   ent.TextString = newtext & Str(v)
   ent.insertionPoint = movex
   'ent.Move exerx1, movex  'insertionPoint2
    Next
     'MsgBox SSET2.Item()
    End If
    
   
    Next ent
    SSET2.Delete
End Sub

Private Sub Command52_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))



Apptivate
On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    varFillet = Val(TextFillet.text)
    For Each ent In SSET2
    
    If ent.EntityType = acPolylineLight Then
    ent.Copy
    ent.Erase
    Thisdrawing.SendCommand "fillet" & vbCr & "r" & vbCr & varFillet & vbCr & "p" & vbCr & "l" & vbCr
    'Thisdrawing.SendCommand "fillet" & vbCr & "p" & vbCr & "l" & vbCr '& "la" & vbCr & "p-sidewalk" & vbCr
'Thisdrawing.SendCommand "p" & vbCr
    End If
    
    
        
    
   
   
    Next ent
    SSET2.Delete
End Sub

Private Sub Command53_Click()

Dim AutoCAD As acadapplication
Dim Thisdrawing As AcadDocument  'as Object
Dim activedocument As Object
Dim modelspace As AcadModelSpace 'not used
Dim paperspace As AcadPaperSpace 'not used
Dim SSET2 As Object
Dim insertionPoint(0 To 2) As Variant 'not used
Dim ent As AcadEntity 'as Object
Dim Excel As Object
Dim excelSheet As Object
Dim application As Object
Dim C As Integer, R As Integer
Dim entLayer As String
Dim entTextString As String
Dim entColor As Variant
Dim entName As String
Dim entName2
Dim entHandle As String

    On Error Resume Next
    
    Set Excel = GetObject(, "Excel.Application")
    If Err <> 0 Then
      Err.Clear
        Set Excel = CreateObject("Excel.Application")
        If Err <> 0 Then
            MsgBox "Could not load Excel.", vbExclamation
            End
        End If
    End If
    On Error 'GoTo 0
    
    Excel.Visible = True
    Excel.Sheets("Sheet1").Select
    Set excelSheet = Excel.ActiveWorkbook.Sheets("Sheet1")
    

Set acadapp = GetObject(, "autocad.application")
Set Thisdrawing = acadapp.activedocument
Set acadapp = GetObject(, "AutoCAD.Application")
Set Thisdrawing = acadapp.activedocument

Apptivate

On Error Resume Next
    
   Set SSET2 = Thisdrawing.SelectionSets.Add("strSet1")

R = 1
SSET2.SelectOnScreen

    AppActivate ("Microsoft Excel")
    
    For Each ent In SSET2
    
    
        
        
        entLayer = ent.Layer
        entArea = ent.Area
        entClosed = ent.Closed
        entName = ent.entityname
        'entName2 = ent.Name
        entHandle = (ent.Area / 43560)
            
        'write values to Excel cells
        excelSheet.Cells(R, 1).Value = entLayer
        excelSheet.Cells(R, 2).Value = entArea
        excelSheet.Cells(R, 3).Value = entClosed
        excelSheet.Cells(R, 4).Value = entName
        'excelSheet.Cells(R, 5).Value = entName2
        excelSheet.Cells(R, 5).Value = entHandle
    
    
    
    R = R + 1
    
    entArea = 0
    Next ent
    
    Text1.text = SSET2.Count - 1
    
    SSET2.Delete
    
    
    excelSheet.Range("A1").Select
    
    
    Excel.Selection.EntireRow.Insert
    
    excelSheet.Range("A1:F1").Select
    
    Excel.Selection.AutoFilter
    
    excelSheet.Range("A1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Layer"
   
    excelSheet.Range("B1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Area"
   
    excelSheet.Range("C1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Closed"
    
    excelSheet.Range("D1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Entity"
    
    excelSheet.Range("E1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Acres"
    
    excelSheet.Range("F1").Select
    
    Excel.ActiveCell.FormulaR1C1 = " <<< Select Filter"
    excelSheet.Range("H1").Select
    
    excelSheet.Range("A1:G1").Select
   
    Excel.Selection.Font.Bold = True
    
    
    
    C = 1
    excelSheet.Columns(C).AutoFit
    C = 2
    excelSheet.Columns(C).AutoFit
    C = 3
    excelSheet.Columns(C).AutoFit
    C = 4
    excelSheet.Columns(C).AutoFit
    C = 5
    excelSheet.Columns(C).AutoFit
    C = 6
    excelSheet.Columns(C).AutoFit
    C = 7
    excelSheet.Columns(C).AutoFit
     
    
    excelSheet.Rows("1:1").Select
    excelSheet.Selection.RowHeight = 22.5
    excelSheet.Range("A1:G1").Select
    With excelSheet.Selection.Interior
        .ColorIndex = 37
        .Pattern = xlSolid
    End With
    
   


    excelSheet.Cells.Select
    With excelSheet.Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .MergeCells = False
    End With
    excelSheet.Range("A1").Select



End Sub

Private Sub Command54_Click()
Dim AutoCAD As acadapplication
Dim Thisdrawing As AcadDocument
Set acadapp = GetObject(, "autocad.application")
Set Thisdrawing = acadapp.activedocument
Set obviews = Thisdrawing.Views
For Each obview In obviews
List1.AddItem obview.Name '& vbCrLf

Next

Exit Sub
zerp:
MsgBox Err.Description
Exit Sub
End Sub

Private Sub Command55_Click()
Dim AutoCAD As acadapplication
Dim Thisdrawing As AcadDocument  'as Object
 Set acadapp = GetObject(, "autocad.application")
    
Set Thisdrawing = acadapp.activedocument
    Dim viewObj As AcadView
    Set viewObj = Thisdrawing.Views.Item(TextView.text)
    
    ' Set the view characteristics
    ''viewObj.Center(0) = 374: viewObj.Center(1) = 313
    ''viewObj.Width = 450
    ''viewObj.Height = 354
    
    ' Get the current active viewport
    Dim viewportObj As AcadViewport
    Set viewportObj = Thisdrawing.ActiveViewport
   
       
    ' Set the view in the viewport
    viewportObj.SetView viewObj
    Thisdrawing.ActiveViewport = viewportObj
        
    Thisdrawing.Regen True
End Sub

Private Sub Command56_Click()
Dim AutoCAD As acadapplication
Dim Thisdrawing As AcadDocument  'as Object
 Set acadapp = GetObject(, "autocad.application")
    
Set Thisdrawing = acadapp.activedocument
    Dim viewObj As AcadView
    
    
    ' Set the view characteristics
    ''viewObj.Center(0) = 374: viewObj.Center(1) = 313
    ''viewObj.Width = 450
    ''viewObj.Height = 354
    
    ' Get the current active viewport
    Dim viewportObj As AcadViewport
    Set viewportObj = Thisdrawing.ActiveViewport
   
       
    ' Set the view in the viewport
    For v = 0 To List1.ListCount - 1
    TextView.text = List1.List(v)
    Set viewObj = Thisdrawing.Views.Item(TextView.text)
    viewportObj.SetView viewObj
    Thisdrawing.ActiveViewport = viewportObj
        
    Thisdrawing.Regen True
    Next
End Sub

Private Sub Command57_Click()
Dim AutoCAD As acadapplication
Dim Thisdrawing As AcadDocument
Set acadapp = GetObject(, "autocad.application")
Set Thisdrawing = acadapp.activedocument
Set obblks = Thisdrawing.Blocks
For Each obblk In obblks
List1.AddItem obblk.Name '& vbCrLf

Next

Exit Sub
zerp:
MsgBox Err.Description
Exit Sub
End Sub

Private Sub Command58_Click()
Dim AutoCAD As acadapplication
Dim Thisdrawing As AcadDocument  'as Object
Dim movex(0 To 2) As Double
 Set acadapp = GetObject(, "autocad.application")
    
Set Thisdrawing = acadapp.activedocument
Apptivate
With Thisdrawing.Utility
    .InitializeUserInput 1
    
    insertionPoint2 = Thisdrawing.Utility.GetPoint(, vbCr & "insertion: ")
   End With
'ent.Move exerx1, insertionPoint2
'movex(0) = insertionPoint2(0)
'movex(1) = insertionPoint2(1)
   
'ent.Copy

x = insertionPoint2
NEWNAME = TextView.text
    
    
    Set blockrefobj = Thisdrawing.modelspace.InsertBlock(x, NEWNAME, 1#, 1#, 1#, 0)
    
Dim MyValue
Randomize                       'Initialize random-number generator.
For Y = 0 To 10
ir = Int((6 * Rnd) + 1)        'Generate random value between 1 and 6.
ir2 = Int((-22 * Rnd) + 1)
    blockrefobj.Copy
    movex(0) = insertionPoint2(0) + ir
ir = Int((6 * Rnd) + 1)
    movex(1) = insertionPoint2(1) + ir2 '/ 2
    blockrefobj.Move x, movex
    List1.AddItem x(0) & movex(0)
    Next
    
    'blockrefobj.Copy
    'movex(0) = insertionPoint2(0) + 1
    'movex(1) = insertionPoint2(1) + 1
    'blockrefobj.Move insertionPoint2, movex
    
    'blockrefobj.Copy
    'movex(0) = insertionPoint2(0) + 1
    'movex(1) = insertionPoint2(1) + 1
    'blockrefobj.Move insertionPoint2, movex
    
    'blockrefobj.Copy
    'movex(0) = insertionPoint2(0) + 1
    'movex(1) = insertionPoint2(1) + 1
    'blockrefobj.Move insertionPoint2, movex
    
    'blockrefobj.Copy
    'movex(0) = insertionPoint2(0) + 1
    'movex(1) = insertionPoint2(1) + 1
    'blockrefobj.Move insertionPoint2, movex
    
    'blockrefobj.Copy
    'movex(0) = insertionPoint2(0) + 1
    'movex(1) = insertionPoint2(1) + 1
    'blockrefobj.Move insertionPoint2, movex
    

    
    
End Sub

Private Sub Command59_Click()

Dim AutoCAD As acadapplication
Dim Thisdrawing As AcadDocument  'as Object
Dim activedocument As Object
Dim modelspace As AcadModelSpace 'not used
Dim paperspace As AcadPaperSpace 'not used
Dim SSET2 As Object
Dim insertionPoint(0 To 2) As Variant 'not used
Dim ent As AcadEntity 'as Object
Dim Excel As Object
Dim excelSheet As Object
Dim application As Object
Dim C As Integer, R As Integer
Dim entLayer As String
Dim entTextString As String
Dim entColor As Variant
Dim entName As String
Dim entName2
Dim entHandle As String

    On Error Resume Next
    
    Set Excel = GetObject(, "Excel.Application")
    If Err <> 0 Then
      Err.Clear
        Set Excel = CreateObject("Excel.Application")
        If Err <> 0 Then
            MsgBox "Could not load Excel.", vbExclamation
            End
        End If
    End If
    On Error 'GoTo 0
    
    Excel.Visible = True
    Excel.Sheets("Sheet1").Select
    Set excelSheet = Excel.ActiveWorkbook.Sheets("Sheet1")
    

Set acadapp = GetObject(, "autocad.application")
Set Thisdrawing = acadapp.activedocument
Set acadapp = GetObject(, "AutoCAD.Application")
Set Thisdrawing = acadapp.activedocument

Apptivate

On Error Resume Next
    
   Set SSET2 = Thisdrawing.SelectionSets.Add("strSet1")

R = 1
SSET2.SelectOnScreen

    AppActivate ("Microsoft Excel")
    
    For Each ent In SSET2
    
    If ent.EntityType = acMtext Then
        
        
        entLayer = ent.TextString
        entArea = ent.Area
        entClosed = ent.Closed
        entName = ent.entityname
        'entName2 = ent.Name
        entHandle = (ent.Area / 43560)
            
        'write values to Excel cells
        excelSheet.Cells(R, 1).Value = entLayer
        excelSheet.Cells(R, 2).Value = entArea
        excelSheet.Cells(R, 3).Value = entClosed
        excelSheet.Cells(R, 4).Value = entName
        'excelSheet.Cells(R, 5).Value = entName2
        excelSheet.Cells(R, 5).Value = entHandle
    
    
    
    R = R + 1
    
    entArea = 0
    
    entLayer = ""
    End If
    Next ent
    
    Text1.text = SSET2.Count - 1
    
    SSET2.Delete
    
    
    excelSheet.Range("A1").Select
    
    
    Excel.Selection.EntireRow.Insert
    
    excelSheet.Range("A1:F1").Select
    
    Excel.Selection.AutoFilter
    
    excelSheet.Range("A1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Layer"
   
    excelSheet.Range("B1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Area"
   
    excelSheet.Range("C1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Closed"
    
    excelSheet.Range("D1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Entity"
    
    excelSheet.Range("E1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Acres"
    
    excelSheet.Range("F1").Select
    
    Excel.ActiveCell.FormulaR1C1 = " <<< Select Filter"
    excelSheet.Range("H1").Select
    
    excelSheet.Range("A1:G1").Select
   
    Excel.Selection.Font.Bold = True
    
    
    
    C = 1
    excelSheet.Columns(C).AutoFit
    C = 2
    excelSheet.Columns(C).AutoFit
    C = 3
    excelSheet.Columns(C).AutoFit
    C = 4
    excelSheet.Columns(C).AutoFit
    C = 5
    excelSheet.Columns(C).AutoFit
    C = 6
    excelSheet.Columns(C).AutoFit
    C = 7
    excelSheet.Columns(C).AutoFit
     
    
    excelSheet.Rows("1:1").Select
    excelSheet.Selection.RowHeight = 22.5
    excelSheet.Range("A1:G1").Select
    With excelSheet.Selection.Interior
        .ColorIndex = 37
        .Pattern = xlSolid
    End With
    
   


    excelSheet.Cells.Select
    With excelSheet.Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .MergeCells = False
    End With
    excelSheet.Range("A1").Select
End Sub

Private Sub Command6_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))
On Error Resume Next
Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    If (TypeOf ent Is AutoCAD.AcadLWPolyline) Then
     ent.Closed = True
    
   End If
   ' End If
    Next ent
End Sub

Private Sub Command60_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))



Apptivate
'On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    If ent.EntityType = acText Then
    ent.TextString = Trim(ent.TextString)
    
        
    
   End If
   
    Next ent
    SSET2.Delete
End Sub

Private Sub Command61_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))



Apptivate
'On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    If ent.EntityType = acText Then
    STRWORD = Trim(UCase(ent.TextString))
    strlen = Len(Trim(ent.TextString))
    For v = 0 To strlen - 1
    strLetter = Mid(STRWORD, v + 1, 1)
    'List1.AddItem strLetter
    Select Case strLetter
        Case "A"
        cntLetter = cntLetter + 1
        Case "B"
        cntLetter = cntLetter + 2
        Case "C"
        cntLetter = cntLetter + 3
        Case "D"
        cntLetter = cntLetter + 4
        Case "E"
        cntLetter = cntLetter + 5
        Case "F"
        cntLetter = cntLetter + 6
        Case "G"
        cntLetter = cntLetter + 7
        Case "H"
        cntLetter = cntLetter + 8
        Case "I"
        cntLetter = cntLetter + 9
        Case "J"
        cntLetter = cntLetter + 10
        Case "K"
        cntLetter = cntLetter + 11
        Case "L"
        cntLetter = cntLetter + 12
        Case "M"
        cntLetter = cntLetter + 13
        Case "N"
        cntLetter = cntLetter + 14
        Case "O"
        cntLetter = cntLetter + 15
        Case "P"
        cntLetter = cntLetter + 16
        Case "Q"
        cntLetter = cntLetter + 17
        Case "R"
        cntLetter = cntLetter + 18
        Case "S"
        cntLetter = cntLetter + 19
        Case "T"
        cntLetter = cntLetter + 20
        Case "U"
        cntLetter = cntLetter + 21
        Case "V"
        cntLetter = cntLetter + 22
        Case "W"
        cntLetter = cntLetter + 23
        Case "X"
        cntLetter = cntLetter + 24
        Case "Y"
        cntLetter = cntLetter + 25
        Case "Z"
        cntLetter = cntLetter + 26
    End Select
    Next
    
      'List1.AddItem cntLetter
    'List1.AddItem strWord & " = " & cntLetter
    List1.AddItem cntLetter & " = " & STRWORD
    cntLetter = 0
   End If
   
    Next ent
    SSET2.Delete
End Sub

Private Sub Command62_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))



Apptivate
'On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    If ent.EntityType = acText Then
    STRWORD = Trim(UCase(ent.TextString))
    strlen = Len(Trim(ent.TextString))
    
    For v = 0 To strlen - 1
    strLetter = Trim(UCase(Mid(STRWORD, v + 1, 1)))
    'List1.AddItem strLetter
    Select Case strLetter
        Case "A"
        cntLetter = cntLetter + Asc(strLetter)
        Case "B"
        cntLetter = cntLetter + Asc(strLetter)
        Case "C"
        cntLetter = cntLetter + Asc(strLetter)
        Case "D"
        cntLetter = cntLetter + Asc(strLetter)
        Case "E"
        cntLetter = cntLetter + Asc(strLetter)
        Case "F"
        cntLetter = cntLetter + Asc(strLetter)
        Case "G"
        cntLetter = cntLetter + Asc(strLetter)
        Case "H"
        cntLetter = cntLetter + Asc(strLetter)
        Case "I"
        cntLetter = cntLetter + Asc(strLetter)
        Case "J"
        cntLetter = cntLetter + Asc(strLetter)
        Case "K"
        cntLetter = cntLetter + Asc(strLetter)
        Case "L"
        cntLetter = cntLetter + Asc(strLetter)
        Case "M"
        cntLetter = cntLetter + Asc(strLetter)
        Case "N"
        cntLetter = cntLetter + Asc(strLetter)
        Case "O"
        cntLetter = cntLetter + Asc(strLetter)
        Case "P"
        cntLetter = cntLetter + Asc(strLetter)
        Case "Q"
        cntLetter = cntLetter + Asc(strLetter)
        Case "R"
        cntLetter = cntLetter + Asc(strLetter)
        Case "S"
        cntLetter = cntLetter + Asc(strLetter)
        Case "T"
        cntLetter = cntLetter + Asc(strLetter)
        Case "U"
        cntLetter = cntLetter + Asc(strLetter)
        Case "V"
        cntLetter = cntLetter + Asc(strLetter)
        Case "W"
        cntLetter = cntLetter + Asc(strLetter)
        Case "X"
        cntLetter = cntLetter + Asc(strLetter)
        Case "Y"
        cntLetter = cntLetter + Asc(strLetter)
        Case "Z"
        cntLetter = cntLetter + Asc(strLetter)
    End Select
    Next
    List1.AddItem cntLetter & " = " & STRWORD
      'List1.AddItem cntLetter
    cntLetter = 0
   End If
   
    Next ent
    SSET2.Delete
End Sub

Private Sub Command63_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))



Apptivate
'On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    If ent.EntityType = acText Then
    STRWORD = Trim(UCase(ent.TextString))
    strlen = Len(Trim(ent.TextString))
    For v = 0 To strlen - 1
    strLetter = Mid(STRWORD, v + 1, 1)
    'List1.AddItem strLetter
    Select Case strLetter
        Case "A"
        cntLetter = cntLetter + 26
        Case "B"
        cntLetter = cntLetter + 25
        Case "C"
        cntLetter = cntLetter + 24
        Case "D"
        cntLetter = cntLetter + 23
        Case "E"
        cntLetter = cntLetter + 22
        Case "F"
        cntLetter = cntLetter + 21
        Case "G"
        cntLetter = cntLetter + 20
        Case "H"
        cntLetter = cntLetter + 19
        Case "I"
        cntLetter = cntLetter + 18
        Case "J"
        cntLetter = cntLetter + 17
        Case "K"
        cntLetter = cntLetter + 16
        Case "L"
        cntLetter = cntLetter + 15
        Case "M"
        cntLetter = cntLetter + 14
        Case "N"
        cntLetter = cntLetter + 13
        Case "O"
        cntLetter = cntLetter + 12
        Case "P"
        cntLetter = cntLetter + 11
        Case "Q"
        cntLetter = cntLetter + 10
        Case "R"
        cntLetter = cntLetter + 9
        Case "S"
        cntLetter = cntLetter + 8
        Case "T"
        cntLetter = cntLetter + 7
        Case "U"
        cntLetter = cntLetter + 6
        Case "V"
        cntLetter = cntLetter + 5
        Case "W"
        cntLetter = cntLetter + 4
        Case "X"
        cntLetter = cntLetter + 3
        Case "Y"
        cntLetter = cntLetter + 2
        Case "Z"
        cntLetter = cntLetter + 1
    End Select
    Next
    
      'List1.AddItem cntLetter
    'List1.AddItem strWord & " = " & cntLetter
    List1.AddItem cntLetter & " = " & STRWORD
    cntLetter = 0
   End If
   
    Next ent
    SSET2.Delete
End Sub

Private Sub Command64_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))



Apptivate
'On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    If ent.EntityType = acText Then
    STRWORD = Trim(UCase(ent.TextString))
    strlen = Len(Trim(ent.TextString))
    x = 0
    For v = x To strlen - 1
    'STRWORD = Trim(UCase(ent.TextString))
    strLetter = Mid(STRWORD, x + 1, 1)
    'List1.AddItem strLetter
    List1.AddItem strLetter
    Select Case strLetter
        Case "A"
        STRWORD = Replace(STRWORD, "A", "Z")
        'GoTo blue:
        Case "B"
        STRWORD = Replace(STRWORD, "B", "Y")
        'GoTo blue:
        Case "C"
        STRWORD = Replace(STRWORD, "C", "X")
        'GoTo blue:
        Case "D"
        STRWORD = Replace(STRWORD, "D", "W")
        'GoTo blue:
        Case "E"
        STRWORD = Replace(STRWORD, "E", "V")
        'GoTo blue:
        Case "F"
        STRWORD = Replace(STRWORD, "F", "U")
        'GoTo blue:
        Case "G"
        STRWORD = Replace(STRWORD, "G", "T")
        'GoTo blue:
        Case "H"
        STRWORD = Replace(STRWORD, "H", "S")
        'GoTo blue:
        Case "I"
        STRWORD = Replace(STRWORD, "I", "R")
        'GoTo blue:
        Case "J"
        STRWORD = Replace(STRWORD, "J", "Q")
        'GoTo blue:
        Case "K"
        STRWORD = Replace(STRWORD, "K", "P")
        'GoTo blue:
        Case "L"
        STRWORD = Replace(STRWORD, "L", "O")
        'GoTo blue:
        Case "M"
        STRWORD = Replace(STRWORD, "M", "N")
        'GoTo blue:
        Case "N"
        STRWORD = Replace(STRWORD, "N", "M")
        'GoTo blue:
        Case "O"
        STRWORD = Replace(STRWORD, "O", "L")
        'GoTo blue:
        Case "P"
        STRWORD = Replace(STRWORD, "P", "K")
        'GoTo blue:
        Case "Q"
        STRWORD = Replace(STRWORD, "Q", "J")
        'GoTo blue:
        Case "R"
        STRWORD = Replace(STRWORD, "R", "I")
        'GoTo blue:
        Case "S"
        STRWORD = Replace(STRWORD, "S", "H")
        'GoTo blue:
        Case "T"
        STRWORD = Replace(STRWORD, "T", "G")
        'GoTo blue:
        Case "U"
        STRWORD = Replace(STRWORD, "U", "F")
        'GoTo blue:
        Case "V"
        STRWORD = Replace(STRWORD, "V", "E")
        'GoTo blue:
        Case "W"
        STRWORD = Replace(STRWORD, "W", "D")
        'GoTo blue:
        Case "X"
        STRWORD = Replace(STRWORD, "X", "C")
        'GoTo blue:
        Case "Y"
        STRWORD = Replace(STRWORD, "Y", "B")
        'GoTo blue:
        Case "Z"
        STRWORD = Replace(STRWORD, "Z", "A")
        'GoTo blue:


blue:
x = x + 1

    End Select

'STRWORD = STRWORD
'ent.Copy
    ent.TextString = STRWORD
    List1.AddItem cntLetter & " = " & STRWORD
    Next

      'List1.AddItem cntLetter
    'List1.AddItem strWord & " = " & cntLetter
'blue:
    cntLetter = 0
   'ent.TextString = STRWORD
   End If
   
    Next ent
    SSET2.Delete
End Sub

Private Sub Command65_Click()
strTest1 = Text1.text
intLen = Len(strTest1)
'For x = 1 To intLen
'ipos = InStr(x, strTest1, "-")
'List1.AddItem ipos
'Next
Y = Split(strTest1, vbCrLf)
For C = LBound(Y) To UBound(Y)
List1.AddItem Y(C)
Next

End Sub

Private Sub Command66_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))



Apptivate
On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    If ent.EntityType = acText Then
    For x = 0 To List1.ListCount - 1
    If Trim(ent.TextString) = List1.List(x) Then
    ent.Color = acRed
    End If
    Next x
    End If
    
    
        
    
   
   
    Next ent
    SSET2.Delete
End Sub

Private Sub Command67_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))



Apptivate
On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    strCheck = TextCheck.text
    For Each ent In SSET2
    
    If ent.EntityType = acMtext Or ent.EntityType = acText Then
    
    If InStr(1, ent.TextString, strCheck) > 0 Then
    'List1.AddItem ent.TextString
    cntx = cntx + 1
    List1.AddItem ent.TextString & " " & cntx
    'ent.Color = acWhite
    End If
    
    End If
    
    
        
    
   
   
    Next ent
    List1.AddItem cntPVx
    SSET2.Delete
End Sub


Private Sub Command68_Click()
List1.Clear
End Sub

Private Sub Command69_Click()

Dim AutoCAD As acadapplication
Dim Thisdrawing As AcadDocument  'as Object
Dim activedocument As Object
Dim modelspace As AcadModelSpace 'not used
Dim paperspace As AcadPaperSpace 'not used
Dim SSET2 As Object
Dim insertionPoint(0 To 2) As Variant 'not used
Dim ent As AcadEntity 'as Object
Dim Excel As Object
Dim excelSheet As Object
Dim application As Object
Dim C As Integer, R As Integer
Dim entLayer As String
Dim entTextString As String
Dim entColor As Variant
Dim entName As String
Dim entName2
Dim entHandle As String

    On Error Resume Next
    
    Set Excel = GetObject(, "Excel.Application")
    If Err <> 0 Then
      Err.Clear
        Set Excel = CreateObject("Excel.Application")
        If Err <> 0 Then
            MsgBox "Could not load Excel.", vbExclamation
            End
        End If
    End If
    'On Error 'GoTo 0
    
    Excel.Visible = True
    Excel.Sheets("Sheet1").Select
    Set excelSheet = Excel.ActiveWorkbook.Sheets("Sheet1")
    

Set acadapp = GetObject(, "autocad.application")
Set Thisdrawing = acadapp.activedocument
Set acadapp = GetObject(, "AutoCAD.Application")
Set Thisdrawing = acadapp.activedocument

Apptivate

On Error Resume Next
    
   Set SSET2 = Thisdrawing.SelectionSets.Add("strSet1")

R = 1
SSET2.SelectOnScreen

    AppActivate ("Microsoft Excel")
    
    For Each ent In SSET2
    
    
        
        
        entSeq = R
        entArea = ent.Area
        entClosed = ent.Closed
        entName = ent.Layer
        'entName2 = ent.Name
        entAcre = (ent.Area / 43560)
            
        'write values to Excel cells
        excelSheet.Cells(R, 1).Value = entSeq
        excelSheet.Cells(R, 2).Value = Round(entArea, 2)
        excelSheet.Cells(R, 3).Value = Round(entAcre, 2)
        excelSheet.Cells(R, 4).Value = entClosed
        'excelSheet.Cells(R, 5).Value = entName2
        excelSheet.Cells(R, 5).Value = entName
    
    If ent.Closed = True Then ent.Color = R
    
    R = R + 1
    entAcre = 0
    entArea = 0
    Next ent
    
    Text1.text = SSET2.Count - 1
    
    SSET2.Delete
    
    
    excelSheet.Range("A1").Select
    
    
    Excel.Selection.EntireRow.Insert
    
    excelSheet.Range("A1:F1").Select
    
    Excel.Selection.AutoFilter
    
    excelSheet.Range("A1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Parcel"
   
    excelSheet.Range("B1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Square Feet"
   
    excelSheet.Range("C1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Acres"
    
    excelSheet.Range("D1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Closed"
    
    excelSheet.Range("E1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Layer"
    
    excelSheet.Range("F1").Select
    
    Excel.ActiveCell.FormulaR1C1 = " <<< Select Filter"
    excelSheet.Range("H1").Select
    
    excelSheet.Range("A1:G1").Select
   
    Excel.Selection.Font.Bold = True
    
    
    
    C = 1
    excelSheet.Columns(C).AutoFit
    C = 2
    excelSheet.Columns(C).AutoFit
    C = 3
    excelSheet.Columns(C).AutoFit
    C = 4
    excelSheet.Columns(C).AutoFit
    C = 5
    excelSheet.Columns(C).AutoFit
    C = 6
    excelSheet.Columns(C).AutoFit
    C = 7
    excelSheet.Columns(C).AutoFit
     
    
    excelSheet.Rows("1:1").Select
    excelSheet.Selection.RowHeight = 49.5
    excelSheet.Range("A1:G1").Select
    With excelSheet.Selection.Interior
        .ColorIndex = 37
        .Pattern = xlSolid
    End With
    
   


    excelSheet.Cells.Select
    With excelSheet.Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .MergeCells = False
    End With
    excelSheet.Range("A1").Select

End Sub

Private Sub Command7_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))
'On Error Resume Next
Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
 Dim lwpl As AutoCAD.AcadLWPolyline
    For Each ent In SSET2
    
      If (TypeOf ent Is AutoCAD.AcadLWPolyline) Then
        On Error Resume Next
            Set lwpl = ent
     lwpl.Width = 20
    
   End If
   ' End If
    Next ent
End Sub

Private Sub Command70_Click()

Dim AutoCAD As acadapplication
Dim Thisdrawing As AcadDocument  'as Object
Dim activedocument As Object
Dim modelspace As AcadModelSpace 'not used
Dim paperspace As AcadPaperSpace 'not used
Dim SSET2 As Object
Dim insertionPoint(0 To 2) As Variant 'not used
Dim ent As AcadEntity 'as Object
Dim Excel As Object
Dim excelSheet As Object
Dim application As Object
Dim C As Integer, R As Integer
Dim entLayer As String
Dim entTextString As String
Dim entColor As Variant
Dim entName As String
Dim entName2
Dim entHandle As String

    On Error Resume Next
    
    Set Excel = GetObject(, "Excel.Application")
    If Err <> 0 Then
      Err.Clear
        Set Excel = CreateObject("Excel.Application")
        If Err <> 0 Then
            MsgBox "Could not load Excel.", vbExclamation
            End
        End If
    End If
    On Error 'GoTo 0
    
    Excel.Visible = True
    Excel.Sheets("Sheet1").Select
    Set excelSheet = Excel.ActiveWorkbook.Sheets("Sheet1")
    

Set acadapp = GetObject(, "autocad.application")
Set Thisdrawing = acadapp.activedocument
Set acadapp = GetObject(, "AutoCAD.Application")
Set Thisdrawing = acadapp.activedocument

Apptivate

On Error Resume Next
    
   Set SSET2 = Thisdrawing.SelectionSets.Add("strSet1")

R = 1
SSET2.SelectOnScreen

    AppActivate ("Microsoft Excel")
    
    For Each ent In SSET2
    
    If ent.EntityType = acMtext Or ent.EntityType = acText Then
        
        
        entSeq = R
        entLabel = ent.TextString
        strRep = Replace(ent.TextString, ",", "")
        strlen = Len(strRep)
        strRep = Mid(strRep, 1, strlen - 3)
        entAcre = Val((strRep) / 43560)
        'entClosed = ent.Closed
        'entName = ent.entityname
        'entName2 = ent.Name
        'entHandle = (ent.Area / 43560)
            
        'write values to Excel cells
        excelSheet.Cells(R, 1).Value = entSeq
        excelSheet.Cells(R, 2).Value = entLabel
        excelSheet.Cells(R, 3).Value = entAcre
        'excelSheet.Cells(R, 4).Value = entName
        'excelSheet.Cells(R, 5).Value = entName2
        'excelSheet.Cells(R, 5).Value = entHandle
    
    
    
    R = R + 1
    entAcre = 0
    entSeq = 0
    entLabel = ""
    End If
    Next ent
    
    Text1.text = SSET2.Count - 1
    
    SSET2.Delete
    
    
    excelSheet.Range("A1").Select
    
    
    Excel.Selection.EntireRow.Insert
    
    excelSheet.Range("A1:F1").Select
    
    Excel.Selection.AutoFilter
    
    excelSheet.Range("A1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Parcel"
   
    excelSheet.Range("B1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Area"
   
    excelSheet.Range("C1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Acres"
    
    excelSheet.Range("D1").Select
    
    Excel.ActiveCell.FormulaR1C1 = " <<< Select Filter"
    excelSheet.Range("H1").Select
    
    excelSheet.Range("A1:G1").Select
   
    Excel.Selection.Font.Bold = True
    
    
    
    C = 1
    excelSheet.Columns(C).AutoFit
    C = 2
    excelSheet.Columns(C).AutoFit
    C = 3
    excelSheet.Columns(C).AutoFit
    C = 4
    excelSheet.Columns(C).AutoFit
    
    excelSheet.Rows("1:1").Select
    excelSheet.Selection.RowHeight = 42.5
    excelSheet.Range("A1:G1").Select
    With excelSheet.Selection.Interior
        .ColorIndex = 37
        .Pattern = xlSolid
    End With
    
   
    excelSheet.Cells.Select
    With excelSheet.Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .MergeCells = False
    End With
    excelSheet.Range("A1").Select
End Sub

Private Sub Command71_Click()

Dim AutoCAD As acadapplication
Dim Thisdrawing As AcadDocument  'as Object
Dim activedocument As Object
Dim modelspace As AcadModelSpace 'not used
Dim paperspace As AcadPaperSpace 'not used
Dim SSET2 As Object
Dim insertionPoint(0 To 2) As Variant 'not used
Dim ent As AcadEntity 'as Object
Dim Excel As Object
Dim excelSheet As Object
Dim application As Object
Dim C As Integer, R As Integer
Dim entLayer As String
Dim entTextString As String
Dim entColor As Variant
Dim entName As String
Dim entName2
Dim entHandle As String

    On Error Resume Next
    
    Set Excel = GetObject(, "Excel.Application")
    If Err <> 0 Then
      Err.Clear
        Set Excel = CreateObject("Excel.Application")
        If Err <> 0 Then
            MsgBox "Could not load Excel.", vbExclamation
            End
        End If
    End If
    On Error 'GoTo 0
    
    Excel.Visible = True
    Excel.Sheets("Sheet1").Select
    Set excelSheet = Excel.ActiveWorkbook.Sheets("Sheet1")
    

Set acadapp = GetObject(, "autocad.application")
Set Thisdrawing = acadapp.activedocument
Set acadapp = GetObject(, "AutoCAD.Application")
Set Thisdrawing = acadapp.activedocument
Dim Count
Apptivate

On Error Resume Next
    
   Set SSET2 = Thisdrawing.SelectionSets.Add("strSet3")

R = 1
SSET2.SelectOnScreen

    AppActivate ("Microsoft Excel")
    
    For Each ent In SSET2
    
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
    If ent.entityname = "AcDbBlockReference" Then
        
            array1 = ent.GetAttributes
                End If
        For Count = LBound(array1) To UBound(array1)

          'array1(Count).TextString
          'array1(Count).TagString
               
        
       
    'Next
    
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\
        
        
        ent1 = array1(0).TextString
        ent2 = array1(1).TextString
        ent3 = array1(2).TextString
        ent4 = array1(3).TextString
        'entName2 = ent.Name
        'ent5 = array1(4).TextString
            
        'write values to Excel cells
        excelSheet.Cells(R, 1).Value = ent1
        excelSheet.Cells(R, 2).Value = ent2
        excelSheet.Cells(R, 3).Value = ent3
        excelSheet.Cells(R, 4).Value = ent4
        'excelSheet.Cells(R, 5).Value = entName2
        'excelSheet.Cells(R, 5).Value = ent5
    
    Next Count
    
    R = R + 1
    
    entArea = 0
    Next ent
    
    Text1.text = SSET2.Count - 1
    
    SSET2.Delete
    
    
    excelSheet.Range("A1").Select
    
    
    Excel.Selection.EntireRow.Insert
    
    excelSheet.Range("A1:F1").Select
    
    Excel.Selection.AutoFilter
    
    excelSheet.Range("A1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Dat"
   
    excelSheet.Range("B1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Acres"
   
    excelSheet.Range("C1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "S.F."
    
    excelSheet.Range("D1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Seq."
    
    'excelSheet.Range("E1").Select
    
    'Excel.ActiveCell.FormulaR1C1 = "Acres"
    
    excelSheet.Range("E1").Select
    
    Excel.ActiveCell.FormulaR1C1 = " <<< Select Filter"
    excelSheet.Range("H1").Select
    
    excelSheet.Range("A1:G1").Select
   
    Excel.Selection.Font.Bold = True
    
    
    
    C = 1
    excelSheet.Columns(C).AutoFit
    C = 2
    excelSheet.Columns(C).AutoFit
    C = 3
    excelSheet.Columns(C).AutoFit
    C = 4
    excelSheet.Columns(C).AutoFit
    C = 5
    excelSheet.Columns(C).AutoFit
    C = 6
    excelSheet.Columns(C).AutoFit
    C = 7
    excelSheet.Columns(C).AutoFit
     
    
    excelSheet.Rows("1:1").Select
    'excelSheet.Selection.RowHeight = 22.5
    excelSheet.Range("A1:G1").Select
    With excelSheet.Selection.Interior
        .ColorIndex = 37
        .Pattern = xlSolid
    End With
    
   


    excelSheet.Cells.Select
    With excelSheet.Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .MergeCells = False
    End With
    excelSheet.Range("A1").Select



End Sub

Private Sub Command72_Click()
strTest1 = Text1.text
intLen = Len(strTest1)

List1.Clear
'For x = 1 To intLen
'ipos = InStr(x, strTest1, "-")
'List1.AddItem ipos
'Next
Y = Split(strTest1, vbCrLf)
For C = LBound(Y) To UBound(Y)
x = Split(Y(C), " ")
For z = LBound(x) To UBound(x)
If x(z) <> "" Then
List1.AddItem x(z)
End If
Next
Next
End Sub

Private Sub Command73_Click()

Dim AutoCAD As acadapplication
Dim Thisdrawing As AcadDocument  'as Object
Dim activedocument As Object
Dim modelspace As AcadModelSpace 'not used
Dim paperspace As AcadPaperSpace 'not used
Dim SSET2 As Object
Dim insertionPoint(0 To 2) As Variant 'not used
Dim ent As AcadEntity 'as Object
Dim Excel As Object
Dim excelSheet As Object
Dim application As Object
Dim C As Integer, R As Integer
Dim entLayer As String
Dim entTextString As String
Dim entColor As Variant
Dim entName As String
Dim entName2
Dim entHandle As String

    On Error Resume Next
    
    Set Excel = GetObject(, "Excel.Application")
    If Err <> 0 Then
      Err.Clear
        Set Excel = CreateObject("Excel.Application")
        If Err <> 0 Then
            MsgBox "Could not load Excel.", vbExclamation
            End
        End If
    End If
    On Error GoTo 0
    
    Excel.Visible = True
    Excel.Sheets("Sheet1").Select
    Set excelSheet = Excel.ActiveWorkbook.Sheets("Sheet1")
    

'Set acadapp = GetObject(, "autocad.application")
'Set Thisdrawing = acadapp.activedocument
'Set acadapp = GetObject(, "AutoCAD.Application")
'Set Thisdrawing = acadapp.activedocument

Apptivate

On Error Resume Next
    
   'Set SSET2 = Thisdrawing.SelectionSets.Add("strSet1")

R = 1
'SSET2.SelectOnScreen

    AppActivate ("Microsoft Excel")
    
    For xx = 0 To List1.ListCount - 1
    
    
        
        
        entLayer = List1.List(xx)
        
           If ent.Layer <> "" Then
        'write values to Excel cells
        excelSheet.Cells(R, 1).Value = entLayer
        End If
       
    
    
    R = R + 1
    
    entArea = 0
    Next
    
    Text1.text = SSET2.Count - 1
    
    SSET2.Delete
    
    
    excelSheet.Range("A1").Select
    
    
    Excel.Selection.EntireRow.Insert
    
    excelSheet.Range("A1:F1").Select
    
    Excel.Selection.AutoFilter
    
    excelSheet.Range("A1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Layer"
   
    excelSheet.Range("B1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Area"
   
    excelSheet.Range("C1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Closed"
    
    excelSheet.Range("D1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Entity"
    
    excelSheet.Range("E1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Acres"
    
    excelSheet.Range("F1").Select
    
    Excel.ActiveCell.FormulaR1C1 = " <<< Select Filter"
    excelSheet.Range("H1").Select
    
    excelSheet.Range("A1:G1").Select
   
    Excel.Selection.Font.Bold = True
    
    
    
    ''C = 1
    ''excelSheet.Columns(C).AutoFit
    ''C = 2
    ''excelSheet.Columns(C).AutoFit
    ''C = 3
    ''excelSheet.Columns(C).AutoFit
    ''C = 4
    ''excelSheet.Columns(C).AutoFit
    ''C = 5
    ''excelSheet.Columns(C).AutoFit
    ''C = 6
    ''excelSheet.Columns(C).AutoFit
    ''C = 7
    ''excelSheet.Columns(C).AutoFit
     
    
    excelSheet.Rows("1:1").Select
    excelSheet.Selection.RowHeight = 22.5
    excelSheet.Range("A1:G1").Select
    With excelSheet.Selection.Interior
        .ColorIndex = 37
        .Pattern = xlSolid
    End With
    
   


    excelSheet.Cells.Select
    With excelSheet.Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .MergeCells = False
    End With
    excelSheet.Range("A1").Select


End Sub

Private Sub Command74_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))

'Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

Apptivate
'On Error Resume Next
'SSET2.SelectOnScreen
    'autocad.ActiveDocument.
    List1.AddItem Thisdrawing.WindowTitle
    
   
      
End Sub

Private Sub Command75_Click()
Dim AutoCAD As acadapplication
Dim Thisdrawing As AcadDocument  'as Object
Dim movex(0 To 2) As Double
Dim insertionPoint2(0 To 2) As Double
 Set acadapp = GetObject(, "autocad.application")
    
Set Thisdrawing = acadapp.activedocument
Apptivate

For Y = 0 To ListInsert(0).ListCount - 1
insertionPoint2(0) = ListInsert(0).List(Y)
insertionPoint2(1) = ListInsert(1).List(Y)
x = insertionPoint2
NEWNAME = TextView.text
    
    
    Set blockrefobj = Thisdrawing.modelspace.InsertBlock(x, NEWNAME, 1#, 1#, 1#, 0)
    Next
End Sub

Private Sub Command76_Click()
Dim AutoCAD As acadapplication
Dim Thisdrawing As AcadDocument  'as Object
Dim movex(0 To 2) As Double
 Set acadapp = GetObject(, "autocad.application")
    
Set Thisdrawing = acadapp.activedocument
Apptivate

For x = 0 To 49
With Thisdrawing.Utility
    .InitializeUserInput 1
    
    insertionPoint2 = Thisdrawing.Utility.GetPoint(, vbCr & "insertion" & x & ": ")
   End With
   ListInsert(0).AddItem insertionPoint2(0)
   ListInsert(1).AddItem insertionPoint2(1)
Next


    
    
    
    
    

    
End Sub

Private Sub Command77_Click()

Dim AutoCAD As acadapplication
Dim Thisdrawing As AcadDocument  'as Object
Dim activedocument As Object
Dim modelspace As AcadModelSpace 'not used
Dim paperspace As AcadPaperSpace 'not used
Dim SSET2 As Object
Dim insertionPoint(0 To 2) As Variant 'not used
Dim ent As AcadEntity 'as Object
Dim Excel As Object
Dim excelSheet As Object
Dim application As Object
Dim C As Integer, R As Integer
Dim entLayer As String
Dim entTextString As String
Dim entColor As Variant
Dim entName As String
Dim entName2
Dim entHandle As String

    On Error Resume Next
    
    Set Excel = GetObject(, "Excel.Application")
    If Err <> 0 Then
      Err.Clear
        Set Excel = CreateObject("Excel.Application")
        If Err <> 0 Then
            MsgBox "Could not load Excel.", vbExclamation
            End
        End If
    End If
    On Error GoTo 0
    
    Excel.Visible = True
    Excel.Sheets("Sheet1").Select
    Set excelSheet = Excel.ActiveWorkbook.Sheets("Sheet1")
    

Set acadapp = GetObject(, "autocad.application")
Set Thisdrawing = acadapp.activedocument
Set acadapp = GetObject(, "AutoCAD.Application")
Set Thisdrawing = acadapp.activedocument

Apptivate

On Error Resume Next
    
   'Set SSET2 = Thisdrawing.SelectionSets.Add("strSet1")

R = 1
'SSET2.SelectOnScreen

    AppActivate ("Microsoft Excel")
    
    
    Set SSET2 = Thisdrawing.SelectionSets.Add("TEST_SSET")
        SSET2.Select acSelectionSetAll
        'ssetObj.SelectOnScreen
    
      For Each ent In SSET2
        
        entLayer = ent.Layer
        entArea = ent.TextString
        entClosed = ent.ObjectID
        entName = ent.entityname
        'entName2 = ent.Name
        entHandle = ent.Handle
            
        'write values to Excel cells
        excelSheet.Cells(R, 1).Value = entLayer
        excelSheet.Cells(R, 2).Value = entArea
        excelSheet.Cells(R, 3).Value = entClosed
        excelSheet.Cells(R, 4).Value = entName
        'excelSheet.Cells(R, 5).Value = entName2
        excelSheet.Cells(R, 5).Value = entHandle
    
    
    
    R = R + 1
    
    entArea = 0
    Next ent
    
    Text1.text = SSET2.Count - 1
    
    SSET2.Delete
    
    
    excelSheet.Range("A1").Select
    
    
    Excel.Selection.EntireRow.Insert
    
    excelSheet.Range("A1:F1").Select
    
    Excel.Selection.AutoFilter
    
    excelSheet.Range("A1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Layer"
   
    excelSheet.Range("B1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Area"
   
    excelSheet.Range("C1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Closed"
    
    excelSheet.Range("D1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Entity"
    
    excelSheet.Range("E1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Acres"
    
    excelSheet.Range("F1").Select
    
    Excel.ActiveCell.FormulaR1C1 = " <<< Select Filter"
    excelSheet.Range("H1").Select
    
    excelSheet.Range("A1:G1").Select
   
    Excel.Selection.Font.Bold = True
    
    
    
    C = 1
    excelSheet.Columns(C).AutoFit
    C = 2
    excelSheet.Columns(C).AutoFit
    C = 3
    excelSheet.Columns(C).AutoFit
    C = 4
    excelSheet.Columns(C).AutoFit
    C = 5
    excelSheet.Columns(C).AutoFit
    C = 6
    excelSheet.Columns(C).AutoFit
    C = 7
    excelSheet.Columns(C).AutoFit
     
    
    excelSheet.Rows("1:1").Select
    excelSheet.Selection.RowHeight = 22.5
    excelSheet.Range("A1:G1").Select
    With excelSheet.Selection.Interior
        .ColorIndex = 37
        .Pattern = xlSolid
    End With
    
   


    excelSheet.Cells.Select
    With excelSheet.Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .MergeCells = False
    End With
    excelSheet.Range("A1").Select
End Sub

Private Sub Command78_Click()

Dim AutoCAD As acadapplication
Dim Thisdrawing As AcadDocument  'as Object
Dim activedocument As Object
Dim modelspace As AcadModelSpace 'not used
Dim paperspace As AcadPaperSpace 'not used
Dim SSET2 As Object
Dim insertionPoint(0 To 2) As Variant 'not used
Dim ent As AcadEntity 'as Object
Dim Excel As Object
Dim excelSheet As Object
Dim application As Object
Dim C As Integer, R As Integer
Dim entLayer As String
Dim entTextString As String
Dim entColor As Variant
Dim entName As String
Dim entName2
Dim entHandle As String

    On Error Resume Next
    
    Set Excel = GetObject(, "Excel.Application")
    If Err <> 0 Then
      Err.Clear
        Set Excel = CreateObject("Excel.Application")
        If Err <> 0 Then
            MsgBox "Could not load Excel.", vbExclamation
            End
        End If
    End If
    On Error GoTo 0
    
    Excel.Visible = True
    Excel.Sheets("Sheet1").Select
    Set excelSheet = Excel.ActiveWorkbook.Sheets("Sheet1")
    

Set acadapp = GetObject(, "autocad.application")
Set Thisdrawing = acadapp.activedocument
Set acadapp = GetObject(, "AutoCAD.Application")
Set Thisdrawing = acadapp.activedocument

Apptivate

On Error Resume Next
    
   'Set SSET2 = Thisdrawing.SelectionSets.Add("strSet1")

R = 1
'SSET2.SelectOnScreen

    AppActivate ("Microsoft Excel")
    
    
    Set SSET2 = Thisdrawing.SelectionSets.Add("TEST_SSET")
        SSET2.Select acSelectionSetAll
        'ssetObj.SelectOnScreen
    
      For Each ent In SSET2
        
        entLayer = ent.Layer
        entArea = ent.TextString
        entClosed = ent.ObjectID
        entName = ent.entityname
        'entName2 = ent.Name
        entHandle = ent.Handle
            
        'write values to Excel cells
        For R = 1 To 720
        If ent.ObjectID = Trim(excelSheet.Cells(R, 3).Value) Then
        ent.TextString = Trim(excelSheet.Cells(R, 2).Value)
        excelSheet.Cells(R, 2).Value
        End If
        R = R + 1
        Next R
    
    
    
    
    entArea = 0
    Next ent
    
    Text1.text = SSET2.Count - 1
    
    SSET2.Delete
    
    
    excelSheet.Range("A1").Select
    
    
    Excel.Selection.EntireRow.Insert
    
    excelSheet.Range("A1:F1").Select
    
    Excel.Selection.AutoFilter
    
    excelSheet.Range("A1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Layer"
   
    excelSheet.Range("B1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Area"
   
    excelSheet.Range("C1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Closed"
    
    excelSheet.Range("D1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Entity"
    
    excelSheet.Range("E1").Select
    
    Excel.ActiveCell.FormulaR1C1 = "Acres"
    
    excelSheet.Range("F1").Select
    
    Excel.ActiveCell.FormulaR1C1 = " <<< Select Filter"
    excelSheet.Range("H1").Select
    
    excelSheet.Range("A1:G1").Select
   
   
   
    
End Sub

Private Sub Command8_Click()
For v = 0 To List1.ListCount - 1
Text1.text = Text1.text & List1.List(v) & vbCrLf
Next
End Sub

Private Sub Command9_Click()
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
 
Dim ent As AcadEntity 'as Object
Dim entityname As String
Dim hatchob As AcadHatch
ripTime = Mid(Str(Timer), 1, 6)
   strSet = Trim("SSET" & Trim(ripTime))



Apptivate
On Error Resume Next

   Set SSET2 = Thisdrawing.SelectionSets.Add(CStr(strSet))

'Apptivate
'On Error Resume Next
SSET2.SelectOnScreen
    'For Each ent In Thisdrawing.paperspace
    For Each ent In SSET2
    
    If ent.EntityType = acText Then
    'set ent = autocad
    ent.Erase
    End If
    
    
        
    
   
   
    Next ent
    SSET2.Delete
End Sub

Private Sub CmdGetSetVar_Click()
Dim AutoCAD As acadapplication
Dim activedocument As Object
Dim acadapplication As Object
Dim acadapp As acadapplication  '
Dim Thisdrawing  As AcadDocument
Dim sysVarName As String
Dim varData As Variant
Dim intData As Integer
Dim sysVarData

On Error Resume Next
Set acadapp = GetObject(, "autocad.application")
      acadapp.Visible = True
      'Apptivate
    Set Thisdrawing = acadapp.activedocument
    
sysVarName = "FILEDIA"
varData = Thisdrawing.GetVariable(sysVarName)
varBefore = varData

'sysVarName2 = "CMDDIA"
'varData2 = Thisdrawing.GetVariable(sysVarName2)
'varBefore2 = varData2

sysVarName = "FILEDIA"
intData = 0
sysVarData = intData
Thisdrawing.SetVariable sysVarName, sysVarData

'sysVarName2 = "CMDDIA"
'intData2 = 0
'sysVarData2 = intData2
'Thisdrawing.SetVariable sysVarName2, sysVarData2

End Sub

Private Sub CmdResetVar_Click()
Dim AutoCAD As acadapplication
Dim activedocument As Object
Dim acadapplication As Object
Dim acadapp As acadapplication  '
Dim Thisdrawing  As AcadDocument
Dim sysVarName As String
Dim varData As Variant
Dim intData As Integer
Dim sysVarData

On Error Resume Next
Set acadapp = GetObject(, "autocad.application")
      acadapp.Visible = True
      
    Set Thisdrawing = acadapp.activedocument

        sysVarName = "FILEDIA"
        intData = varBefore
        sysVarData = intData
        Thisdrawing.SetVariable sysVarName, sysVarData

End Sub

Sub Example_SetView()
    ' This example creates a new view.
    ' It then changes the active viewport to
    ' the newly created view.
    
    ' First, open a sample drawing.
    Thisdrawing.application.Documents.Open "C:\AutoCAD\Sample\campus.dwg"
    
    ' Create a new view
    Dim viewObj As AcadView
    Set viewObj = Thisdrawing.Views.Add("TESTVIEW")
    
    ' Set the view characteristics
    viewObj.Center(0) = 374: viewObj.Center(1) = 313
    viewObj.Width = 450
    viewObj.Height = 354
    
    ' Get the current active viewport
    Dim viewportObj As AcadViewport
    Set viewportObj = Thisdrawing.ActiveViewport
    MsgBox "Change to the saved view.", , "SetView Example"
       
    ' Set the view in the viewport
    viewportObj.SetView viewObj
    Thisdrawing.ActiveViewport = viewportObj
        
    Thisdrawing.Regen True
    
End Sub

Private Sub List1_Click()
TextView.text = List1.List(List1.ListIndex)
End Sub
