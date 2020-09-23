VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FmTVremote 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Location for WBlock"
   ClientHeight    =   12855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "MSE_Form2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12855
   ScaleWidth      =   5535
   Begin VB.TextBox TextDirectory 
      Height          =   375
      Left            =   120
      TabIndex        =   100
      Text            =   "Text10"
      Top             =   11070
      Width           =   4965
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00000000&
      Caption         =   "Open Drawing"
      ForeColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   90
      TabIndex        =   99
      Top             =   9750
      Width           =   1575
   End
   Begin VB.CommandButton Commanddir 
      Height          =   315
      Left            =   2520
      TabIndex        =   98
      Top             =   10320
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton Command18 
      Height          =   120
      Index           =   2
      Left            =   2730
      Picture         =   "MSE_Form2.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   97
      ToolTipText     =   "This Form Not On Top"
      Top             =   615
      Width           =   255
   End
   Begin VB.CommandButton Command18 
      Height          =   135
      Index           =   0
      Left            =   2730
      Picture         =   "MSE_Form2.frx":101D
      Style           =   1  'Graphical
      TabIndex        =   96
      ToolTipText     =   "This Form Always On Top"
      Top             =   450
      Width           =   255
   End
   Begin VB.CommandButton Command22 
      Caption         =   "<<<"
      Height          =   375
      Left            =   4710
      TabIndex        =   95
      Top             =   330
      Width           =   675
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Zoom "
      Height          =   525
      Left            =   90
      TabIndex        =   94
      ToolTipText     =   "Start an open AutoCAD session with a new document in it before zooming drawings."""
      Top             =   10080
      Width           =   2205
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1245
      Left            =   -60
      TabIndex        =   92
      Top             =   11640
      Width           =   6075
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "September, 2003"
         Height          =   255
         Left            =   2250
         TabIndex        =   93
         Top             =   90
         Width           =   1275
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   540
         Picture         =   "MSE_Form2.frx":1370
         Top             =   210
         Width           =   4500
      End
   End
   Begin VB.DirListBox Dir5 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   3015
      Left            =   60
      TabIndex        =   87
      Top             =   780
      Width           =   5355
   End
   Begin VB.DriveListBox Drive5 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   60
      TabIndex        =   86
      Top             =   450
      Width           =   1635
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Send to Excel"
      Height          =   525
      Left            =   3180
      Style           =   1  'Graphical
      TabIndex        =   85
      ToolTipText     =   "Send Image Names to Excel."
      Top             =   10080
      Width           =   2205
   End
   Begin VB.ComboBox Combo5 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0FF&
      Height          =   315
      Left            =   1740
      TabIndex        =   84
      Text            =   ".dvb"
      Top             =   450
      Width           =   945
   End
   Begin VB.FileListBox File5 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   5940
      Left            =   60
      MultiSelect     =   2  'Extended
      TabIndex        =   83
      Top             =   3780
      Width           =   5370
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00400040&
      Height          =   345
      Left            =   10980
      TabIndex        =   0
      Top             =   8070
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CommandButton cmdlstvw 
         Height          =   525
         Left            =   390
         Picture         =   "MSE_Form2.frx":6868
         Style           =   1  'Graphical
         TabIndex        =   89
         ToolTipText     =   "Return to Visual Palette, keeping connections to paths, but minimizing this form."
         Top             =   390
         Width           =   285
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   345
         Left            =   7200
         TabIndex        =   40
         Top             =   7380
         Visible         =   0   'False
         Width           =   1395
         Begin VB.Frame Frame3 
            Caption         =   "Frame3"
            Height          =   810
            Left            =   390
            TabIndex        =   65
            Top             =   270
            Visible         =   0   'False
            Width           =   4905
            Begin VB.TextBox Text2 
               BackColor       =   &H00000000&
               ForeColor       =   &H0000FFFF&
               Height          =   315
               Index           =   0
               Left            =   360
               TabIndex        =   82
               Top             =   720
               Width           =   135
            End
            Begin VB.TextBox Text3 
               BackColor       =   &H00000000&
               ForeColor       =   &H0000FFFF&
               Height          =   315
               Index           =   0
               Left            =   360
               TabIndex        =   81
               Top             =   1080
               Width           =   90
            End
            Begin VB.TextBox Text1 
               BackColor       =   &H00000000&
               ForeColor       =   &H0000FFFF&
               Height          =   315
               Index           =   0
               Left            =   360
               TabIndex        =   80
               Top             =   360
               Width           =   135
            End
            Begin VB.TextBox Text4 
               BackColor       =   &H00000000&
               ForeColor       =   &H0000FFFF&
               Height          =   315
               Index           =   0
               Left            =   360
               TabIndex        =   79
               Top             =   1440
               Width           =   105
            End
            Begin VB.TextBox Text5 
               BackColor       =   &H00000000&
               ForeColor       =   &H0000FFFF&
               Height          =   315
               Index           =   0
               Left            =   360
               TabIndex        =   78
               Top             =   1800
               Width           =   105
            End
            Begin VB.TextBox Text6 
               BackColor       =   &H00000000&
               ForeColor       =   &H0000FFFF&
               Height          =   315
               Index           =   0
               Left            =   360
               TabIndex        =   77
               Top             =   2160
               Width           =   135
            End
            Begin VB.TextBox Text7 
               BackColor       =   &H00000000&
               ForeColor       =   &H0000FFFF&
               Height          =   315
               Index           =   0
               Left            =   360
               TabIndex        =   76
               Top             =   2520
               Width           =   135
            End
            Begin VB.TextBox Text8 
               BackColor       =   &H00000000&
               ForeColor       =   &H0000FFFF&
               Height          =   315
               Index           =   0
               Left            =   360
               TabIndex        =   75
               Top             =   2880
               Width           =   105
            End
            Begin VB.ListBox lstblocks3 
               Height          =   3960
               Left            =   4470
               TabIndex        =   74
               Top             =   150
               Width           =   2655
            End
            Begin VB.ListBox List2 
               Height          =   645
               Left            =   2415
               TabIndex        =   73
               Top             =   300
               Width           =   2535
            End
            Begin VB.CommandButton Command5 
               Caption         =   "paste list"
               Height          =   255
               Left            =   3570
               TabIndex        =   72
               Top             =   120
               Width           =   855
            End
            Begin VB.TextBox text911 
               Height          =   285
               Left            =   6150
               TabIndex        =   71
               Text            =   "0"
               Top             =   315
               Width           =   615
            End
            Begin VB.TextBox Text22 
               Height          =   315
               Left            =   2490
               TabIndex        =   70
               Text            =   "Text22"
               Top             =   345
               Width           =   1155
            End
            Begin VB.TextBox Text21 
               Height          =   285
               Left            =   1200
               TabIndex        =   69
               Text            =   "Text21"
               Top             =   345
               Width           =   1215
            End
            Begin VB.ListBox List1 
               Height          =   840
               Left            =   2640
               TabIndex        =   68
               Top             =   285
               Width           =   2535
            End
            Begin VB.TextBox Textpaste 
               Height          =   3915
               IMEMode         =   3  'DISABLE
               Left            =   2010
               MultiLine       =   -1  'True
               ScrollBars      =   3  'Both
               TabIndex        =   67
               Top             =   255
               Width           =   2685
            End
            Begin VB.ListBox lstPaths 
               Height          =   255
               Index           =   0
               Left            =   2190
               TabIndex        =   66
               Top             =   405
               Width           =   4275
            End
         End
         Begin VB.ListBox lstPaths 
            Height          =   2595
            Index           =   2
            Left            =   0
            TabIndex        =   64
            Top             =   660
            Width           =   7485
         End
         Begin VB.ListBox lstPaths 
            Height          =   2400
            Index           =   1
            Left            =   4170
            TabIndex        =   63
            Top             =   0
            Width           =   7485
         End
         Begin VB.ListBox List3 
            Height          =   2010
            Left            =   0
            TabIndex        =   62
            Top             =   390
            Width           =   12375
         End
         Begin VB.Frame framealert 
            BackColor       =   &H00008080&
            BorderStyle     =   0  'None
            Caption         =   "Frame7"
            Height          =   255
            Left            =   390
            TabIndex        =   61
            Top             =   1350
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton Command19 
            Caption         =   "Clipboard"
            Height          =   465
            Left            =   1830
            TabIndex        =   60
            Top             =   2340
            Width           =   1305
         End
         Begin VB.CommandButton Command18 
            Caption         =   "Create Folders"
            Height          =   465
            Index           =   1
            Left            =   1830
            TabIndex        =   59
            Top             =   1800
            Width           =   1305
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Add to List"
            Height          =   465
            Left            =   -60
            TabIndex        =   58
            Top             =   2340
            Width           =   1305
         End
         Begin VB.CommandButton Command111 
            Caption         =   "Strip String"
            Height          =   465
            Left            =   -60
            TabIndex        =   57
            Top             =   1800
            Width           =   1305
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Filter"
            Height          =   330
            Left            =   1440
            TabIndex        =   56
            Top             =   3720
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Rename"
            Height          =   465
            Left            =   -60
            TabIndex        =   55
            Top             =   3420
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.CommandButton Command20 
            Caption         =   "Code Library"
            Height          =   585
            Left            =   90
            TabIndex        =   54
            Top             =   1170
            Width           =   1425
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Paths to Text"
            Height          =   585
            Left            =   1560
            TabIndex        =   53
            Top             =   1170
            Width           =   1425
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Clear"
            Height          =   465
            Left            =   5430
            TabIndex        =   52
            Top             =   5430
            Width           =   1305
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Load Text Box"
            Height          =   465
            Left            =   5550
            TabIndex        =   51
            Top             =   4860
            Width           =   1305
         End
         Begin VB.ListBox List4 
            Height          =   2985
            Left            =   60
            TabIndex        =   50
            Top             =   480
            Width           =   5265
         End
         Begin VB.TextBox Text99 
            BackColor       =   &H00000000&
            ForeColor       =   &H0000FF00&
            Height          =   5760
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   49
            Top             =   3600
            Width           =   11640
         End
         Begin VB.TextBox Textx 
            BackColor       =   &H00000000&
            ForeColor       =   &H0000FF00&
            Height          =   3495
            Left            =   2370
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   48
            ToolTipText     =   "Text Window"
            Top             =   990
            Width           =   8895
         End
         Begin VB.TextBox Text9 
            Height          =   450
            Left            =   300
            TabIndex        =   47
            Text            =   "ptfImage"
            Top             =   1230
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.CheckBox CKaddtext 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Add Path to Textbox"
            Height          =   285
            Left            =   90
            TabIndex        =   46
            Top             =   1320
            Width           =   1785
         End
         Begin VB.CommandButton cmdClip 
            Enabled         =   0   'False
            Height          =   405
            Left            =   180
            Picture         =   "MSE_Form2.frx":6C41
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Copy current selection to clipboard."
            Top             =   720
            Width           =   435
         End
         Begin VB.CommandButton Command15 
            Enabled         =   0   'False
            Height          =   345
            Left            =   210
            Picture         =   "MSE_Form2.frx":6FE4
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Save list..."
            Top             =   1050
            Width           =   375
         End
         Begin VB.CommandButton Command16 
            Caption         =   "Add FileNames"
            Height          =   465
            Left            =   210
            TabIndex        =   43
            Top             =   900
            Width           =   1305
         End
         Begin VB.ListBox LstHousE 
            BackColor       =   &H00000000&
            ForeColor       =   &H0000FF00&
            Height          =   3180
            Left            =   660
            TabIndex        =   42
            Top             =   600
            Width           =   14880
         End
         Begin VB.ListBox List5 
            Height          =   4935
            Left            =   240
            Sorted          =   -1  'True
            TabIndex        =   41
            Top             =   750
            Width           =   4905
         End
         Begin VB.Timer Timer1 
            Interval        =   3000
            Left            =   270
            Top             =   480
         End
      End
      Begin VB.Frame frFreeze 
         Caption         =   "Frame7"
         Height          =   195
         Left            =   2130
         TabIndex        =   22
         Top             =   13530
         Visible         =   0   'False
         Width           =   195
         Begin VB.CommandButton Command1 
            Height          =   345
            Left            =   210
            Picture         =   "MSE_Form2.frx":73AA
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   375
            Width           =   345
         End
         Begin VB.CommandButton save1 
            Height          =   345
            Index           =   0
            Left            =   600
            Picture         =   "MSE_Form2.frx":7B58
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   375
            Width           =   600
         End
         Begin VB.CommandButton Command10 
            BackColor       =   &H00000000&
            Height          =   345
            Left            =   2565
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Minimize"
            Top             =   345
            Width           =   345
         End
         Begin VB.CommandButton Command13 
            Height          =   345
            Left            =   1320
            Picture         =   "MSE_Form2.frx":8093
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Open default profile"
            Top             =   375
            Width           =   345
         End
         Begin VB.CommandButton Command14 
            Height          =   345
            Left            =   1695
            Picture         =   "MSE_Form2.frx":81ED
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   375
            Width           =   345
         End
         Begin VB.Frame Frame2 
            Caption         =   "Clear Text"
            Height          =   975
            Left            =   1110
            TabIndex        =   33
            Top             =   630
            Width           =   1815
         End
         Begin VB.CommandButton Command8 
            Caption         =   "X"
            Height          =   375
            Left            =   60
            TabIndex        =   32
            Top             =   420
            Width           =   495
         End
         Begin VB.CommandButton stealth1 
            Caption         =   "&Load Block Paths/Clipboard"
            Height          =   705
            Left            =   600
            TabIndex        =   31
            Top             =   120
            Width           =   1275
         End
         Begin VB.ComboBox Combo55 
            Height          =   315
            Left            =   11
            TabIndex        =   30
            Text            =   "Combo1"
            Top             =   390
            Width           =   1515
         End
         Begin VB.Frame Frame5 
            Caption         =   "Frame5"
            Height          =   435
            Left            =   0
            TabIndex        =   23
            Top             =   570
            Visible         =   0   'False
            Width           =   2265
            Begin VB.CommandButton Command9 
               Height          =   525
               Left            =   360
               TabIndex        =   28
               Top             =   285
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.CheckBox Check1 
               Height          =   225
               Left            =   435
               Picture         =   "MSE_Form2.frx":85B3
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   255
               Width           =   1335
            End
            Begin VB.CommandButton save1 
               Height          =   435
               Index           =   1
               Left            =   1710
               Picture         =   "MSE_Form2.frx":8AF6
               Style           =   1  'Graphical
               TabIndex        =   26
               Top             =   1020
               Width           =   465
            End
            Begin VB.TextBox TPATH 
               Height          =   1410
               Left            =   150
               MultiLine       =   -1  'True
               ScrollBars      =   3  'Both
               TabIndex        =   25
               Top             =   1485
               Width           =   1800
            End
            Begin VB.ListBox lstPathCollector 
               Height          =   3375
               Left            =   150
               TabIndex        =   24
               Top             =   7600
               Width           =   2040
            End
            Begin VB.Label Label1 
               Caption         =   "Click to set  ^^^ Addmode to ON"
               Height          =   405
               Left            =   0
               TabIndex        =   29
               Top             =   255
               Width           =   1365
            End
            Begin VB.Shape Shape1s 
               BorderColor     =   &H00C0FFFF&
               BorderWidth     =   3
               Height          =   330
               Index           =   0
               Left            =   0
               Top             =   240
               Visible         =   0   'False
               Width           =   555
            End
         End
         Begin VB.Shape shSave 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H000000FF&
            BorderStyle     =   2  'Dash
            BorderWidth     =   7
            FillColor       =   &H00C0FFFF&
            Height          =   405
            Left            =   165
            Top             =   345
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label Label25 
            BackColor       =   &H8000000A&
            Caption         =   "Open/Save Profiles Files"
            Height          =   240
            Left            =   210
            TabIndex        =   39
            Top             =   150
            Width           =   1455
         End
         Begin VB.Line Line2 
            X1              =   0
            X2              =   2880
            Y1              =   210
            Y2              =   210
         End
      End
      Begin VB.CommandButton Command3 
         Height          =   345
         Left            =   2010
         Picture         =   "MSE_Form2.frx":92A4
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   13050
         Visible         =   0   'False
         Width           =   345
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   900
         Top             =   1710
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   615
         Left            =   645
         TabIndex        =   90
         Top             =   780
         Width           =   1545
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Height          =   555
         Left            =   2910
         TabIndex        =   19
         Top             =   885
         Width           =   525
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Height          =   135
         Left            =   1950
         TabIndex        =   18
         ToolTipText     =   "Button 8 Row 2"
         Top             =   2070
         Width           =   240
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Height          =   135
         Left            =   1725
         TabIndex        =   17
         ToolTipText     =   "Button 7 Row 2"
         Top             =   2070
         Width           =   240
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Height          =   135
         Left            =   1470
         TabIndex        =   16
         ToolTipText     =   "Button 6 Row 2"
         Top             =   2070
         Width           =   240
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Height          =   135
         Left            =   1230
         TabIndex        =   15
         ToolTipText     =   "Button 5 Row 2"
         Top             =   2070
         Width           =   240
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Height          =   135
         Left            =   960
         TabIndex        =   14
         ToolTipText     =   "Button 4 Row 2"
         Top             =   2070
         Width           =   240
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Height          =   135
         Left            =   705
         TabIndex        =   13
         ToolTipText     =   "Button 3 Row 2"
         Top             =   2085
         Width           =   240
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Height          =   135
         Left            =   465
         TabIndex        =   12
         ToolTipText     =   "Button 2 Row 2"
         Top             =   2070
         Width           =   240
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Height          =   135
         Left            =   225
         TabIndex        =   11
         ToolTipText     =   "Button 1 Row 2"
         Top             =   2055
         Width           =   240
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1935
         TabIndex        =   10
         ToolTipText     =   "Button 8 Row 1"
         Top             =   1830
         Width           =   240
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Height          =   165
         Left            =   1725
         TabIndex        =   9
         ToolTipText     =   "Button 7 Row 1"
         Top             =   1860
         Width           =   240
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1470
         TabIndex        =   8
         ToolTipText     =   "Button 6 Row 1"
         Top             =   1830
         Width           =   240
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Height          =   210
         Left            =   1230
         TabIndex        =   7
         ToolTipText     =   "Button 5 Row 1"
         Top             =   1815
         Width           =   240
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Height          =   210
         Left            =   960
         TabIndex        =   6
         ToolTipText     =   "Button 4 Row 1"
         Top             =   1815
         Width           =   240
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   2220
         TabIndex        =   5
         ToolTipText     =   "Button 3 Row 1"
         Top             =   1230
         Width           =   240
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   465
         TabIndex        =   4
         ToolTipText     =   "Button 2 Row 1"
         Top             =   1815
         Width           =   240
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H0080FF80&
         Height          =   135
         Left            =   225
         TabIndex        =   3
         ToolTipText     =   "Button 1 Row 1"
         Top             =   1845
         Width           =   240
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   9885
      TabIndex        =   21
      Top             =   3480
      Width           =   1740
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5280
      Left            =   -165
      TabIndex        =   2
      Top             =   3690
      Width           =   195
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Zoom a group of drawings to extents."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   285
      Left            =   120
      TabIndex        =   91
      Top             =   60
      Width           =   5355
   End
   Begin VB.Label fileTot 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FF8080&
      Height          =   285
      Left            =   30
      TabIndex        =   88
      Top             =   9720
      Width           =   5265
   End
   Begin VB.Label Labelx 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   8820
      TabIndex        =   20
      Top             =   3330
      Width           =   3015
   End
End
Attribute VB_Name = "FmTVremote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public x As Integer
Dim text As String

Private Sub cmdClip_Click()
On Error GoTo hulu:
Clipboard.SetText LstHousE.List(LstHousE.ListIndex)
framealert.Visible = True
Exit Sub
hulu:
MsgBox Err.Description
End Sub

Private Sub Cmdlstvw_Click()
'On Error Resume Next 'remX
'Me.WindowState = 1
'acadXshow
End Sub
Private Sub listPnosee()
On Error Resume Next 'remX
Do Until inter = lstPaths.Count - 1
inter = inter + 1
lstPaths(inter).Visible = False
Loop
Command9_Click
End Sub
Private Sub listPsee()
Do Until inter = lstPaths.Count - 1
inter = inter + 1
lstPaths(inter).Visible = True
Loop
Command9_Click
End Sub
Private Sub Command1_Click()
Dim strItem() As String
Dim stritem2 As String
Dim intCounter As Integer
On Error Resume Next 'remX
appz = App.Path
'Open "C:\testfile.txt" For Input As #1
CommonDialog1.Filter = "Text Files|*.txt|All Files|*.*"
CommonDialog1.DialogTitle = "Plantacious Remote Palette Profiles"
CommonDialog1.CancelError = False
CommonDialog1.InitDir = appz
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
Open CommonDialog1.FileName For Input As #1
'filesize = LOF(1)
'Text = Input$(filesize, #1)

Do While Not EOF(1)
ReDim Preserve strItem(intCounter)
Input #1, strItem(intCounter)
Text1(1).text = strItem(intCounter)
intCounter = intCounter + 1

Do While Not EOF(1)
ReDim Preserve strItem(intCounter)
Input #1, strItem(intCounter)
Text2(1).text = strItem(intCounter)
intCounter = intCounter + 1

Do While Not EOF(1)
ReDim Preserve strItem(intCounter)
Input #1, strItem(intCounter)
Text3(1).text = strItem(intCounter)
intCounter = intCounter + 1

Do While Not EOF(1)
ReDim Preserve strItem(intCounter)
Input #1, strItem(intCounter)
Text4(1).text = strItem(intCounter)
intCounter = intCounter + 1

Do While Not EOF(1)
ReDim Preserve strItem(intCounter)
Input #1, strItem(intCounter)
Text5(1).text = strItem(intCounter)
intCounter = intCounter + 1

Do While Not EOF(1)
ReDim Preserve strItem(intCounter)
Input #1, strItem(intCounter)
Text6(1).text = strItem(intCounter)
intCounter = intCounter + 1
Do While Not EOF(1)
ReDim Preserve strItem(intCounter)
Input #1, strItem(intCounter)
Text7(1).text = strItem(intCounter)
intCounter = intCounter + 1

Do While Not EOF(1)
ReDim Preserve strItem(intCounter)
Input #1, strItem(intCounter)
Text8(1).text = strItem(intCounter)
intCounter = intCounter + 1

Do While Not EOF(1)
ReDim Preserve strItem(intCounter)
Input #1, strItem(intCounter)
Text1(2).text = strItem(intCounter)
intCounter = intCounter + 1

Do While Not EOF(1)
ReDim Preserve strItem(intCounter)
Input #1, strItem(intCounter)
Text2(2).text = strItem(intCounter)
intCounter = intCounter + 1

Do While Not EOF(1)
ReDim Preserve strItem(intCounter)
Input #1, strItem(intCounter)
Text3(2).text = strItem(intCounter)
intCounter = intCounter + 1

Do While Not EOF(1)
ReDim Preserve strItem(intCounter)
Input #1, strItem(intCounter)
Text4(2).text = strItem(intCounter)
intCounter = intCounter + 1

Do While Not EOF(1)
ReDim Preserve strItem(intCounter)
Input #1, strItem(intCounter)
Text5(2).text = strItem(intCounter)
intCounter = intCounter + 1

Do While Not EOF(1)
ReDim Preserve strItem(intCounter)
Input #1, strItem(intCounter)
Text6(2).text = strItem(intCounter)
intCounter = intCounter + 1

Do While Not EOF(1)
ReDim Preserve strItem(intCounter)
Input #1, strItem(intCounter)
Text7(2).text = strItem(intCounter)
intCounter = intCounter + 1

Do While Not EOF(1)
ReDim Preserve strItem(intCounter)
Input #1, strItem(intCounter)
Text8(2).text = strItem(intCounter)
intCounter = intCounter + 1





';;;;;;;;;;;;
';;;;;;;;;;;;
Do While Not EOF(1)
ReDim Preserve strItem(intCounter)
Input #1, strItem(intCounter)
Text21.text = strItem(intCounter)
intCounter = intCounter + 1

Do While Not EOF(1)
ReDim Preserve strItem(intCounter)
Input #1, strItem(intCounter)
Text22.text = strItem(intCounter)
intCounter = intCounter + 1
'888888888888888888
Loop
Loop
Loop
Loop
Loop
Loop
Loop
Loop
Loop
Loop
Loop
Loop
Loop
Loop
Loop
Loop
Loop
Loop



Close #1

End If
End Sub

Private Sub Command10_Click()
On Error Resume Next
'Me.WindowState = 1
'acadXshow
End Sub






Private Sub Command11_Click()
Dim x As Long, i As Long
For x = 0 To Dir5.ListCount - 1
'File5.Path = Dir5.List(x)
'For i = 0 To Dir5.ListCount - 1
LenStr = Len(Dir5.List(x))
ipos = InStrRev(Dir5.List(x), "\")
i = LenStr - ipos
strAdd = Right(Dir5.List(x), i)
For Y = 0 To List5.ListCount - 1
If Trim(strAdd) = List5.List(Y) Then
DontAdd = True
End If
Next
If DontAdd <> True Then
List5.AddItem strAdd
End If
Next
  





End Sub

Private Sub Command12_Click()
Textx.text = ""
For t = 0 To List5.ListCount - 1
Textx.text = Textx.text & List5.List(t) & vbCrLf
Next
End Sub

Private Sub Command13_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
filerX = App.Path & "\DefProfile.txt"
Command13.ToolTipText = filerX
End Sub

Private Sub Command14_Click()
On Error Resume Next
'
'verv = AppPathX(VPATH)
 Dim lFile As Integer
 On Error GoTo snot:
  lFile = FreeFile
    filex = Trim(App.Path) & "\" & "DefProfile.txt"
    Open filex For Output As 1
  Print #lFile, Text1(1).text
  Print #lFile, Text2(1).text
  Print #lFile, Text3(1).text
  Print #lFile, Text4(1).text
  Print #lFile, Text5(1).text
  Print #lFile, Text6(1).text
  Print #lFile, Text7(1).text
  Print #lFile, Text8(1).text
  
  Print #lFile, Text1(2).text
  Print #lFile, Text2(2).text
  Print #lFile, Text3(2).text
  Print #lFile, Text4(2).text
  Print #lFile, Text5(2).text
  Print #lFile, Text6(2).text
  Print #lFile, Text7(2).text
  Print #lFile, Text8(2).text
  Close #lFile
 Exit Sub
snot:
 MsgBox "......" & Err.Description
 Exit Sub
 
End Sub

Private Sub Command14_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
'On Error Resume Next
filex = Trim(App.Path) & "\" & "DefProfile.txt"
Command14.ToolTipText = filex
End Sub

Private Sub Command16_Click()
Dim x As Long, i As Long
For x = 0 To File5.ListCount - 1
'File5.Path = Dir5.List(x)
'For i = 0 To Dir5.ListCount - 1

strAdd = File5.List(x)
For Y = 0 To List5.ListCount - 1
If Trim(strAdd) = List5.List(Y) Then
DontAdd = True
End If
Next
If DontAdd <> True Then
List5.AddItem strAdd
End If
Next
End Sub

Private Sub Command17_Click()
List5.Clear
End Sub

Private Sub eeCommand18_Click()

For Y = 0 To List5.ListCount - 1
dirx = List5.List(Y)
MkDir ("c:\11111\" & dirx)
MkDir ("c:\11111\" & dirx & "_Port")
Next
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
Dim strItem() As String
Dim stritem2 As String
Dim intCounter As Integer
On Error Resume Next 'remX
appz = App.Path
Open "C:\dvb.txt" For Input As #1



Do While Not EOF(1)
ReDim Preserve strItem(intCounter)
Input #1, strItem(intCounter)
List5.AddItem strItem(intCounter)
intCounter = intCounter + 1
Loop
End Sub

Private Sub Command2_Click()
For x = 0 To File5.ListCount - 1
ipos = InStr(1, File5.List(x), Text9.text)
If ipos > 0 Then
'If x < File5.ListCount - 1 Then
Textx.text = Textx.text & File5.List(x) & vbCrLf
'End If
'If x = File5.ListCount - 1 Then
'Textx.text = Textx.text & File5.List(x)
'End If
Text9.text = Trim(Text9.text)
End If
Next
End Sub

Private Sub Command15_Click()


Dim x As Integer
    x = 0
        Text51 = ""
  '{FileOpen "path\origfile.wb2"}
  '{FileSaveAs "path\fileXL.xls";Confirm;;Excel v5/v7}
  msave = "{FileSaveAs "
  m1 = "{FileOpen "
  mquote = Chr(34)
  
  Dim strMain As String
  Dim ipos As Integer
  Dim ipos2 As Integer
  Dim strlen
  Dim strFilex As String
  
  Do Until x = LstHousE.ListCount
  strMain = Trim(LstHousE.List(x))
  MainLen = Len(strMain)
  ipos = InStrRev(strMain, "\")
  
  
  strLenx = (MainLen - ipos + 1)
  strFilex = Mid(strMain, ipos + 1, strLenx)
  ipos2 = InStrRev(strFilex, ".")
  strLenx = Len(strFilex)
  strFilex2 = Mid(strFilex, 1, ipos2 - 1) & ".xls"
  strPathx = Mid(strMain, 1, ipos)
  
  newpath = Trim(strPathx & strFilex2)
  
  'MsgBox strFilex & vbCrLf & _
         'strFilex2 & vbCrLf & _
         'strPathx & vbCrLf & _
         'newpath
         strEnd = Trim("{FileClose 1}")
         
         macFileOrig = m1 & mquote & LstHousE.List(x) & mquote & "}"
         newFilex = msave & mquote & newpath & mquote & "}"
         macEnd = strEnd
  
 ' Print #lFile, macFileOrig
  
  DoEvents
  Textx = Textx & macFileOrig & vbCrLf & _
    newFilex & vbCrLf & _
    macEnd & vbCrLf
  Text99.text = Text99.text & newpath
  
  x = x + 1
  Loop
 
  
 
End Sub

Private Sub Command20_Click()
FmTVremote2.Show
End Sub

Private Sub Command21_Click()
Dim acadapp As acadapplication  'Object  'acadapplication
Dim acadapplication2 As acadapplication
Dim Thisdrawing As Object
Dim application As AcadDocument
Dim modelspace As Object
Dim strCaption As String
Dim pathTB As String
Dim insertionPoint  '(0 To 2) As Double
Dim acadapp2 As Object
Dim Thisdrawing2 As Object 'AcadDocument  ' Object
Dim sysVarName As String
Dim varData As Variant
Dim intData As Integer
Dim sysVarData
  'On Error Resume Next
On Error GoTo mozo2:
  Set acadapp2 = GetObject(, "AutoCAD.Application")
    Set Thisdrawing2 = acadapp2.activedocument

Set acadapp2 = GetObject(, "AutoCAD.Application")
    Set Thisdrawing2 = acadapp2.activedocument
    
  Set acadapp2 = GetObject(, "AutoCAD.Application")
    Set Thisdrawing2 = acadapp2.activedocument

sysVarName = "SDI"
varData = Thisdrawing2.GetVariable(sysVarName)



'sdi=====================================================
If varData > 0 Then
For x = 0 To LstHousE.ListCount - 1

theDwg = Trim(LstHousE.List(x))
  Set acadapp2 = GetObject(, "AutoCAD.Application")
  Set Thisdrawing2 = acadapp2.activedocument
acadapp2.Visible = True
Apptivate
Thisdrawing2.Open (Trim(theDwg))
acadapp2.ZoomExtents
Thisdrawing2.Save

Next x

End If
'sdi=====================================================

'MDI======================================================
If varData = 0 Then
Set acadapp2 = GetObject(, "AutoCAD.Application")
    Set Thisdrawing2 = acadapp2.activedocument
    Dim NewDrawing1 As AcadDocument
    'Dim Newdrawing2 As AcadDocument
    'Set NewDrawing1 = Thisdrawing2.application.Documents.Add("Holder")

For x = 0 To LstHousE.ListCount - 1
  theDwg = Trim(LstHousE.List(x))
    Set acadapp2 = GetObject(, "AutoCAD.Application")
    Set Thisdrawing2 = acadapp2.activedocument
    acadapp2.Visible = True
    
Thisdrawing2.application.Documents.Open (Trim(theDwg))
Set Thisdrawing2 = acadapp2.activedocument
Thisdrawing2.Activate

acadapp2.ZoomExtents
Thisdrawing2.Save
Thisdrawing2.Close
Next x
Exit Sub
End If
'MDI======================================================
Exit Sub
mozo2:
If Err.Number = 429 Then
MsgBox Err.Description & ".  " & "Start an AutoCAD 2000x session with a new drawing."

End If
If Err.Number = -2145320900 Then
MsgBox Err.Description & ". " & "Open a blank drawing in current session."
End If
End Sub

Private Sub Command22_Click()
On Error Resume Next
Form1.Show
Unload Me
End Sub

Private Sub Command23_Click()

End Sub

Private Sub Command3_Click()

Unload Me
'FmBlk14.Show
FmTVremote.Show
End Sub



Private Sub Command4_Click()
Dim intLen As Integer
Dim ipos As Integer
Dim ipos2 As Integer
Dim strFilePre As String
Dim strFileEnd As String
On Error Resume Next
For x = 0 To File5.ListCount - 1
ipos = InStr(1, File5.FileName, "-")

intLen = Len(File5.List(x))

ipos2 = InStr(1, File5.List(x), ".dwg")
strLeft = intLen - ipos
strFilePre = Mid(File5.List(x), ipos + 1, strLeft - 4)
strFileEnd = Mid(File5.List(x), 1, ipos - 1)
strPath = Mid(LstHousE.List(x), 1, ipos2)
strMondo = strPath & strFilePre & "-" & strFileEnd '& ".dwg"

intFileLen = Len(File5.FileName)
strFileName = Mid(File5.FileName, 1, intFileLen - 4)

'Name LstHousE.List(x) As strMondo

List3.AddItem strFilePre
List3.AddItem strFileEnd
'List3.AddItem strPath
'List3.AddItem ipos2
List3.AddItem Dir5.Path
'List3.AddItem strFileName


'List3.AddItem strMondo
Next
End Sub

Private Sub Command5_Click()
Dim intCounter As Integer
Dim strItem As String
Dim gettext As String

 On Error Resume Next
   'Form1.List1.AddItem Clipboard.gettext
lstListbox1.AddItem strItem
End Sub
Private Sub Command6_Click()
If File5.ListCount = 0 Then
MsgBox "No images are found."
Exit Sub
End If

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
Dim strPath2 As String
Dim strFullpath2 As String
Dim lFile As Integer
lFile = FreeFile

'Get Files Location.
    On Error GoTo mop:
    'get excel if app already open
    
    
    
    Set Excel = GetObject(, "Excel.Application")
    If Err <> 0 Then
      Err.Clear
        'create/open excel if not
        Set Excel = CreateObject("Excel.Application")
        'handle excel error if app cannot be opened.
        If Err <> 0 Then
            MsgBox "Could not load Excel.", vbExclamation
            End
        End If
    End If
    On Error GoTo 0
    'Excel.Visible = True
    'Excel.Sheets("Sheet1").Select
    On Error Resume Next
    Set excelSheet = Excel.ActiveWorkbook.Sheets("Sheet1")
    Dim MyStamp



       Rcnt = File5.ListCount
       For R = 1 To Rcnt
       MyStamp = FileDateTime(File5.List(R - 1))
       
        excelSheet.Cells(R, 1).Value = "Drawing Name"
        excelSheet.Cells(R, 2).Value = Trim(File5.List(R - 1))
        excelSheet.Cells(R, 3).Value = MyStamp
        Excel.ActiveWorkbook.Sheets("Sheet1").Columns(1).AutoFit
        Excel.ActiveWorkbook.Sheets("Sheet1").Columns(2).AutoFit
        Excel.ActiveWorkbook.Sheets("Sheet1").Columns(3).AutoFit
       Next
       
       
        Exit Sub
mop:
        MsgBox Err.Description & ". Start an Excel session."
         
End Sub



Private Sub Command7_Click()
For x = 0 To LstHousE.ListCount - 1
Text99 = Text99 & LstHousE.List(x) & vbCrLf
Next
End Sub

Private Sub Command8_Click()
End
End Sub
Private Sub Command9_Click()
On Error Resume Next
lstPaths(0).Clear
lstPaths(1).Clear
Do Until intz > 1
intz = intz + 1
lstPaths(intz).AddItem Text1(1).text
lstPaths(intz).AddItem Text2(1).text
lstPaths(intz).AddItem Text3(1).text
lstPaths(intz).AddItem Text4(1).text
lstPaths(intz).AddItem Text5(1).text
lstPaths(intz).AddItem Text6(1).text
lstPaths(intz).AddItem Text7(1).text
lstPaths(intz).AddItem Text8(1).text
DoEvents
Loop

End Sub

Private Sub Commanddir_Click()
Dir5_Change
Dir5_Click
End Sub

Private Sub Dir5_Change()
On Error GoTo hutu:
txtCheck = Dir5.Path
lstPathCollector.AddItem Dir5.Path
'x = x + 1
'list4.AddItem dir5.

If Check1.Value = 1 Then
If x = 2 Then
TPATH = TPATH & Dir5.Path & Chr(13) & Chr(10)

End If
End If

    File5.Path = Dir5.Path 'sets file path.
    
    cmdOpenAll5_Click
    If LstHousE.ListCount = 0 Then
    cmdClip.Enabled = False
    End If
    
If x > 2 Then
x = 1
End If
Label2.Caption = x
Command16_Click
If Len(Dir5.Path) <> 3 Then
Form1.TextDirectory.text = Trim(Dir5.Path & "\") & File5.FileName
Else
Form1.TextDirectory.text = Dir5.Path & File5.FileName
End If
Exit Sub


hutu:
    MsgBox Err.Description    '& " " & Err.LastDllError
End Sub

Private Sub Command111_Click()
Dim x As Long, i As Long
For x = 0 To Dir5.ListCount - 1
On Error Resume Next
LenStr = Len(Dir5.List(x))
ipos = InStrRev(Dir5.List(x), "\")
i = LenStr - ipos
strAdd = Right(Dir5.List(x), i)
    List4.AddItem Dir5.List(x)
    Textx.text = Textx.text & strAdd & vbCrLf
    List5.AddItem strAdd
    
    
Next



End Sub





Private Sub File5_Click()
If Len(Dir5.Path) <> 3 Then
TextDirectory.text = Trim(Dir5.Path & "\") & File5.FileName
Else
TextDirectory.text = Dir5.Path & File5.FileName
End If
End Sub

Private Sub File5_DblClick()

If Check2.Value = 0 Then
Dim delFile As String
'If Button = 1 Then
On Error Resume Next
If File5.List(File5.ListIndex) <> "" Then
Dim response As Integer
   
response = MsgBox("This will permanently delete: " & File5.List(File5.ListIndex) & _
vbCrLf & _
vbCrLf & "       Do you wish to delete this file?", 4, "SpeedWBlock2003")

    Select Case response
        
        
        Case vbYes:
       '''''Function kill file
prePath = Dir5.Path & "\"
delFile = Trim(prePath & File5.List(File5.ListIndex))
Kill (delFile)


File5.Refresh
vv = Val(txtlstindex.text)
listdwgelems.RemoveItem (vv)
LstHousE.RemoveItem (vv)
'''''Function kill file

        

        Case vbNo:
'MsgBox "Exited function"
        Exit Sub


    End Select
   End If

    
  Exit Sub
  Else
  
Dim acadapp As acadapplication  'Object  'acadapplication
Dim acadapplication2 As acadapplication
Dim Thisdrawing As Object
Dim modelspace As Object
Dim strCaption As String
Dim pathTB As String
'Dim paperspace As AcadPaperSpace
Dim insertionPoint  '(0 To 2) As Double
'On Error Resume Next 'GoTo zoop:

Dim acadapp2 As Object
Dim Thisdrawing2 As Object 'AcadDocument  ' Object
'Set acadapp = GetObject(, "AutoCAD.Application")
        ''If Err Then
           '' Err.Clear
           
'On Error Resume Next
        

'Set acadapp = GetObject(, "autocad.application")

theDwg = Trim(LstHousE.List(File5.ListIndex))
'on Error Resume Next
    Set acadapp2 = CreateObject("AutoCAD.Application")
    Set Thisdrawing2 = acadapp2.activedocument
  
acadapp2.Visible = True

Thisdrawing2.Open (Trim(theDwg))

Apptivate

''acadapp2.ZoomExtents
''Thisdrawing2.Save
''Thisdrawing2.Close

Set acadapp2 = Nothing
Set Thisdrawing2 = Nothing

End If



Exit Sub
huzu:
MsgBox Err.Description
End Sub






Private Sub Form_Activate()
On Error Resume Next

On Error GoTo errhandler
   Combo5.AddItem "*.dvb"
   
    
    
    Combo5.ListIndex = 0
    File5.Pattern = Combo5.text
    
'zxx = AppPathX(VPATH)
'Me.Caption = zxx
    Dir5.Path = Trim(Form1.TextDirectory.text)
    
    Exit Sub
errhandler:
Exit Sub
End Sub

'++++========dir4
'=====================
'===================== 5
'++++========dir and in form load22
' and in list_click
'dirlist
Private Sub cmdOpenAll5_Click()
Dim PathAndName As String
Dim Path As String
'On Error Resume Next 'remX
File5.Pattern = (Combo5.text)
LstHousE.Clear
If Right(File5.Path, 1) <> "\" Then
Path = File5.Path + "\"
Else
Path = File5.Path
End If

PathAndName = Path + File5.FileName
    Dim i As Integer
    For i = 0 To File5.ListCount - 1
    
LstHousE.AddItem Path & File5.List(i)
fileTot.Caption = Str(LstHousE.ListCount) & " File(s) Found."



Next i

If CKaddtext.Value = 1 Then
Textx.text = Textx.text & Path & vbCrLf
End If


Labelx.Caption = " " & Str(LstHousE.ListCount) & " " & "Drawings."
End Sub
Private Sub Dir5_Click()
On Error Resume Next
File5.Path = Dir5.Path 'sets file path.
    LstHousE.Clear
    fileTot.Caption = "0"
    cmdOpenAll5_Click
    
   x = x + 1
    If x > 2 Then
    x = 1
    End If
    
End Sub

Private Sub Combo5_Click()
On Error Resume Next 'remX
    File5.Pattern = Trim(Combo5.text) 'sets combo choice To file list box
End Sub
Private Sub drive5_Change()
    On Error GoTo skyz:
    Dir5.Path = Drive5.Drive 'sets directory path.
    
Exit Sub
skyz:
MsgBox "Drive unavailable"
Exit Sub
End Sub
Private Sub Dir5b_Change()
On Error Resume Next
File5.Path = Dir5.Path 'sets file path.
    LstHousE.Clear
    cmdOpenAll5_Click
End Sub
'++++========dir5
'=====================

Private Sub AXEL()
'Load Me
''acadXshow
End Sub
Private Sub savex_Click()
 Dim lFile As Integer
On Error Resume Next 'remX
  lFile = FreeFile
    CommonDialog1.DialogTitle = "Save Layer Settings For LandscapeXGenerator"
    CommonDialog1.Filter = "Super LXg Layer settings|*.LXg|Text Files|*.txt"
    CommonDialog1.Flags = 2
    CommonDialog1.CancelError = False
    
    CommonDialog1.FileName = ""
    CommonDialog1.ShowSave
  If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Output As 1
  Print #lFile, TPATH
  Close #lFile
 End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Move (Form1.Left + Form1.Width)
Me.Left = Form1.Left + Form1.Width
Me.Top = Form1.Top
'Command13_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Form1.Show
Unload Me
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next 'remX
'Dir5.BackColor = vbWhite
'Text1(1).BackColor = vbWhite
shSave.Visible = False
Command10.BackColor = vbBlack
End Sub






Private Sub Label21_Click()
Me.Caption = App.Path
End Sub

Private Sub Label22_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dir5.BackColor = vbRed
Text1(1).BackColor = vbWhite
End Sub

Private Sub Label23_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Text1(1).BackColor = vbRed
Dir5.BackColor = vbWhite
'Text1(1).BackColor = vbWhite
End Sub

Private Sub Label24_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
shSave.Visible = True
Text1(1).BackColor = vbWhite
Command10.BackColor = vbBlack
End Sub

Private Sub Label26_Click()
MsgBox "Row 1"

End Sub

Private Sub Label26_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label26.ToolTipText = "Row 1"
Shape1.Visible = True
End Sub

Private Sub Label27_Click()
MsgBox "Row 2"

End Sub

Private Sub Label27_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Shape2.Visible = True
Label27.ToolTipText = "Row 1"
End Sub


Private Sub LstHousE_Click()
On Error Resume Next
LstHousE.ToolTipText = LstHousE.text
cmdClip.Enabled = True

'cmdClip.ToolTipText
End Sub

Private Sub LstHousE_DblClick()
LstHousE.RemoveItem LstHousE.ListIndex
fileTot.Caption = LstHousE.ListCount
End Sub

Private Sub lstPathCollector_Click()
lstPathCollector.ToolTipText = lstPathCollector.text
End Sub

Private Sub Text1_Change(Index As Integer)
  On Error Resume Next 'remX
    If Text1(1) <> "" Then
Text1(1).BackColor = vbWhite
End If
If Text1(2) <> "" Then
Text1(2).BackColor = vbWhite
End If
'savesetting(
End Sub


Private Sub FILE5nosee()
On Error Resume Next
File5.Visible = False
End Sub
Private Sub FILE5see()
On Error Resume Next
File5.Visible = True
End Sub



Private Sub Timer1_Timer()
On Error Resume Next
framealert.Visible = False
End Sub
