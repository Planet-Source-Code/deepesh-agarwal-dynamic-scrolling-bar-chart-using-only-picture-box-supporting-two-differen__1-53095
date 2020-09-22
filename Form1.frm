VERSION 5.00
Begin VB.Form frmDemo 
   BackColor       =   &H00800000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Dynamic Scrolling Bar Graph Demo Using PictureBox"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picbarbox2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      DrawWidth       =   2
      Height          =   2055
      Left            =   4800
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1995
      ScaleWidth      =   3555
      TabIndex        =   6
      Top             =   885
      Width           =   3615
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3960
      Top             =   0
   End
   Begin VB.PictureBox picBarBox 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      DrawWidth       =   2
      Height          =   2055
      Left            =   360
      Picture         =   "Form1.frx":051C
      ScaleHeight     =   1995
      ScaleWidth      =   3555
      TabIndex        =   0
      Top             =   885
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":0A38
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   720
      Index           =   4
      Left            =   360
      TabIndex        =   12
      Top             =   3720
      Width           =   7965
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Scrolling Style"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   240
      Index           =   2
      Left            =   1080
      TabIndex        =   11
      Top             =   600
      Width           =   1995
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Non-Scrolling style"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   240
      Index           =   3
      Left            =   5595
      TabIndex        =   10
      Top             =   600
      Width           =   1995
   End
   Begin VB.Label l2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seed Value : 0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6030
      TabIndex        =   9
      Top             =   3075
      Width           =   1050
   End
   Begin VB.Label lblLegend_Y 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100 ----  ----  ----  ----  ----  ----  ----  ----  ----  000"
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   3
      Left            =   4440
      TabIndex        =   8
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblLegend_Y 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100 ----  ----  ----  ----  ----  ----  ----  ----  ----  000"
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   2
      Left            =   8400
      TabIndex        =   7
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.deepeshagarwal.tk/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   240
      Index           =   1
      Left            =   2760
      TabIndex        =   5
      Top             =   3360
      Width           =   3195
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dynamic Scrolling Bar Graph Demo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   240
      Index           =   0
      Left            =   2640
      TabIndex        =   4
      Top             =   240
      Width           =   3705
   End
   Begin VB.Label lblLegend_Y 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100 ----  ----  ----  ----  ----  ----  ----  ----  ----  000"
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   1
      Left            =   3960
      TabIndex        =   3
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblLegend_Y 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100 ----  ----  ----  ----  ----  ----  ----  ----  ----  000"
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   375
   End
   Begin VB.Label l 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seed Value : 0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1590
      TabIndex        =   1
      Top             =   3120
      Width           =   1050
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================================
' Module    : frmDemo
' DateTime  : 10/04/2004 18:03
' Author    : Deepesh Agarwal
' Purpose   : Display Dynamic bar Chart using only Picture Box
' Website   : http://www.deepeshagarwal.tk/
' E-mail    : mail2me_here@softhome.net
' use the three different types of back bars images supplied
' to make picture box of your size,use any image editing software to make new size from given.
' If you think this code is nice, please do vote for me at PSC,
' this will encourage me to release more code.
' find the link to this code in the @PSC_ReadMe_XXXX.txt file
' in the Zip\folder of this project.
'=========================================================================================

'=========================================================================================
' Dear, VB User, Get my free (reallly free No Adware,spyware or nasty thing like that) NSIS based installer for your VB apps
' (1). Its Easy (Wizard Based)
' (2). Based on rock Solid NSIS Super-PIMP Tecnology
' (3). Will compress your project size to almost half, YES your installer size
'      will be less then your total file size
' (4). Support for adding Splash Skin with fading effect and Sound.
' and much more features, visit product page at http://www.deepeshagarwal.tk
'  Also, Visit my site for Free-Software's like:
'  1). The-AdPolice - Blocks 21000+ adservers to save bandwidth
'  2). Dr. System 2.0 -  Schedule Computer Maintainence - A must for every computer user
'  3). Service Controller XP (A Must For XP User) - Start,Stop,Pause and change startup type of 2000/XP services with recommended settings for different system config.
'  4). Easy Uninstaller 1.0 - Fast and advanced alternative to windows Add\Remove Applet
'   And Many More........

'=========================================================================================


Option Explicit
'For Launching website address
Private Const SW_SHOWNORMAL       As Integer = 1
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Form Level variables
Private BAR_HEIGHT                As Integer
Private Horizontal_Position      As Integer
Private GraphPoints(0 To 99)      As Long

Private Sub Form_Load()
    'Start Our Random Seed Generator for testing purpose
    Randomize
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Display Hyperlink effect
    Label1(1).FontUnderline = False
    Label1(4).FontUnderline = False
End Sub

Private Sub Label1_Click(Index As Integer)
    'Launch Default browser with the web-address
    Select Case Index
        Case 1
            'Launch My Website
            ShellExecute Me.hwnd, vbNullString, "http://www.deepeshagarwal.tk", vbNullString, App.Path, SW_SHOWNORMAL
        Case 4
            'Launch PSC website
            ShellExecute Me.hwnd, vbNullString, "http://www.planet-source-code.com/", vbNullString, App.Path, SW_SHOWNORMAL
    End Select
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Display Hyperlink effect
    Select Case Index
        Case 1
            Label1(1).FontUnderline = True
        Case 4
            Label1(4).FontUnderline = True

    End Select
End Sub

Private Sub Timer1_Timer()
    'This timer is just for Testing, in your project call Updatebar directly
    Dim Random_Seed As Long
    On Error Resume Next
    'Generate Random Number for TESTING
    Random_Seed = Int((100 * Rnd) + 1)

    'Display Scrolling style
    Call UpdateGraph(picBarBox, Random_Seed, True)
    l.Caption = "Seed Value :" & Random_Seed


    'Display Non-Scrolling style
    Call UpdateGraph(picbarbox2, Random_Seed, False)
    l2.Caption = "Seed Value :" & Random_Seed
End Sub

Public Sub UpdateGraph(Pbox As PictureBox, BAR_HEIGHT As Long, isScrollStyle As Boolean)
    On Error Resume Next

    'Prepare Our Picture Box for BAR GRAPH STYLE
    Pbox.ScaleLeft = 0
    Pbox.ScaleTop = 100
    Pbox.ScaleWidth = 100
    Pbox.ScaleHeight = -100

    'Call according to selected style
    If isScrollStyle = True Then
        'Display Using Scrolling Style
        Call ScrollStyle(BAR_HEIGHT, Pbox)
    Else
        'Display Non-Scrolling variety
        Call NonScrollStyle(BAR_HEIGHT, Pbox)
    End If
End Sub

Public Sub NonScrollStyle(BAR_HEIGHT As Long, Pbox As PictureBox)
    'Non-Scrolling variety
    If Horizontal_Position >= 100 Then
        Horizontal_Position = 0
        Pbox.Cls
    Else
        Horizontal_Position = Horizontal_Position + 1
    End If
    'Show Donwload Speed
    Pbox.Line (Horizontal_Position, 0)-(Horizontal_Position, BAR_HEIGHT), vbGreen, BF
End Sub

Private Sub ShiftPoints()
    'Shift BAR GRAPH Points to left to display the scrolling effect
    Dim Cnt As Long
    On Error Resume Next
    'Shift all the points from the graph one place to the left
    For Cnt = LBound(GraphPoints) To UBound(GraphPoints) - 1
        GraphPoints(Cnt) = GraphPoints(Cnt + 1)
    Next Cnt
End Sub

Public Sub ScrollStyle(BAR_HEIGHT As Long, Pbox As PictureBox)
    'The Scrolling Style
    Dim Cnt As Long
    On Error Resume Next
    'Shift points to the left
    Call ShiftPoints
    GraphPoints(UBound(GraphPoints)) = BAR_HEIGHT
    Pbox.Cls
    'Replot, graph using new values
    For Cnt = LBound(GraphPoints) To UBound(GraphPoints) - 1
        Pbox.Line (Cnt, 0)-(Cnt, GraphPoints(Cnt + 1)), vbGreen, BF
    Next Cnt

End Sub
