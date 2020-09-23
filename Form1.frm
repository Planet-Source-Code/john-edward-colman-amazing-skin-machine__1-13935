VERSION 5.00
Object = "{3489755E-DC13-11D4-9242-000102711081}#7.0#0"; "MetalCBProj.ocx"
Begin VB.Form frmExample 
   AutoRedraw      =   -1  'True
   Caption         =   "METAL SKIN"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6705
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   260
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   447
   StartUpPosition =   2  'CenterScreen
   Begin MetalCBProj.MetalCB MetalCB1 
      Height          =   390
      Index           =   3
      Left            =   2250
      TabIndex        =   5
      Top             =   825
      Width           =   1965
      _extentx        =   3466
      _extenty        =   688
      fontsize        =   12
      fontcharset     =   0
      fontbold        =   -1  'True
      fontitalic      =   0   'False
      fontname        =   "Copperplate Gothic Bold"
      fontstrike      =   0   'False
      fontunder       =   0   'False
      fontweight      =   700
      caption         =   "Lead"
      forecolor       =   -2147483634
      shademax        =   100
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2640
      Left            =   300
      ScaleHeight     =   172
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   117
      TabIndex        =   3
      Top             =   825
      Width           =   1815
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Click a button"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   75
         TabIndex        =   4
         Top             =   1200
         Width           =   1650
      End
   End
   Begin MetalCBProj.MetalCB MetalCB1 
      Height          =   765
      Index           =   0
      Left            =   2250
      TabIndex        =   0
      Top             =   1200
      Width           =   1965
      _extentx        =   3466
      _extenty        =   1349
      fontsize        =   12
      fontcharset     =   0
      fontbold        =   -1  'True
      fontitalic      =   -1  'True
      fontname        =   "MS Sans Serif"
      fontstrike      =   0   'False
      fontunder       =   0   'False
      fontweight      =   700
      caption         =   "Blue Steel"
      forecolor       =   16777215
      red             =   0
      blue            =   99
      multiplier      =   -1.3
   End
   Begin MetalCBProj.MetalCB MetalCB1 
      Height          =   765
      Index           =   1
      Left            =   2250
      TabIndex        =   1
      Top             =   1950
      Width           =   1965
      _extentx        =   3466
      _extenty        =   1349
      fontsize        =   18
      fontcharset     =   0
      fontbold        =   0   'False
      fontitalic      =   0   'False
      fontname        =   "Univers"
      fontstrike      =   0   'False
      fontunder       =   0   'False
      fontweight      =   400
      caption         =   "Gold"
      forecolor       =   16777215
      red             =   70
      green           =   30
      blue            =   0
      multiplier      =   -1.3
   End
   Begin MetalCBProj.MetalCB MetalCB1 
      Height          =   765
      Index           =   2
      Left            =   2250
      TabIndex        =   2
      Top             =   2700
      Width           =   1965
      _extentx        =   3466
      _extenty        =   1349
      fontsize        =   18
      fontcharset     =   0
      fontbold        =   -1  'True
      fontitalic      =   0   'False
      fontname        =   "Univers"
      fontstrike      =   0   'False
      fontunder       =   0   'False
      fontweight      =   700
      caption         =   "Silver"
      forecolor       =   16777215
      multiplier      =   -1.3
   End
   Begin MetalCBProj.MetalCB MetalCB1 
      Height          =   390
      Index           =   4
      Left            =   4350
      TabIndex        =   6
      Top             =   825
      Width           =   1965
      _extentx        =   3466
      _extenty        =   688
      fontsize        =   12
      fontcharset     =   0
      fontbold        =   -1  'True
      fontitalic      =   0   'False
      fontname        =   "Copperplate Gothic Bold"
      fontstrike      =   0   'False
      fontunder       =   0   'False
      fontweight      =   700
      caption         =   "Rubber"
      forecolor       =   -2147483634
      red             =   100
      shademax        =   100
   End
   Begin MetalCBProj.MetalCB MetalCB1 
      Height          =   765
      Index           =   5
      Left            =   4350
      TabIndex        =   7
      Top             =   1200
      Width           =   1965
      _extentx        =   3466
      _extenty        =   1349
      fontsize        =   12
      fontcharset     =   0
      fontbold        =   -1  'True
      fontitalic      =   -1  'True
      fontname        =   "MS Sans Serif"
      fontstrike      =   0   'False
      fontunder       =   0   'False
      fontweight      =   700
      caption         =   "Diamond"
      red             =   100
      green           =   100
      blue            =   255
      multiplier      =   -1.3
   End
   Begin MetalCBProj.MetalCB MetalCB1 
      Height          =   765
      Index           =   6
      Left            =   4350
      TabIndex        =   8
      Top             =   2700
      Width           =   1965
      _extentx        =   3466
      _extenty        =   1349
      fontsize        =   18
      fontcharset     =   0
      fontbold        =   0   'False
      fontitalic      =   0   'False
      fontname        =   "Univers"
      fontstrike      =   0   'False
      fontunder       =   0   'False
      fontweight      =   400
      caption         =   "Ice"
      red             =   0
      green           =   40
      blue            =   60
      multiplier      =   -3
      shademin        =   100
      borderstyle     =   1
   End
   Begin MetalCBProj.MetalCB MetalCB1 
      Height          =   765
      Index           =   7
      Left            =   4350
      TabIndex        =   9
      Top             =   1950
      Width           =   1965
      _extentx        =   3466
      _extenty        =   1349
      fontsize        =   18
      fontcharset     =   0
      fontbold        =   -1  'True
      fontitalic      =   0   'False
      fontname        =   "Univers"
      fontstrike      =   0   'False
      fontunder       =   0   'False
      fontweight      =   700
      caption         =   "Slate"
      forecolor       =   16777215
      blue            =   100
      multiplier      =   -1.3
      shademax        =   100
   End
   Begin MetalCBProj.MetalCB MetalCB2 
      Height          =   390
      Left            =   5850
      TabIndex        =   10
      Top             =   150
      Width           =   390
      _extentx        =   688
      _extenty        =   688
      fontsize        =   18
      fontcharset     =   0
      fontbold        =   -1  'True
      fontitalic      =   0   'False
      fontname        =   "Univers"
      fontstrike      =   0   'False
      fontunder       =   0   'False
      fontweight      =   700
      caption         =   "X"
      forecolor       =   4210752
      red             =   0
      green           =   0
      blue            =   88
      multiplier      =   -1.3
      shademax        =   200
      shademin        =   100
      clickdepth      =   1
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------------
' Hey you !!!!
' This code was originally written by John Colman
' If you use this code, give me a credit.
' (john_colman@hotmail.com)
'---------------------------------------------------------------------

'These API calls change are used to change the visible regieon of a window or control
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

'This is used to fool Windows into thinking we clicked where we didn't
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

'We use this to make a sizable borderless window
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const WS_DLGFRAME = &H400000



Private Sub GradientFill(obj As Object, Min As Integer, Max As Integer, r As Integer, g As Integer, b As Integer, Optional Multiplier As Double = 1)
    On Error Resume Next
    Dim i As Long
    Dim c As Integer
    Dim st As Double
    
    Max = Max - Min
    st = obj.ScaleHeight / 3.142 * Abs(Multiplier)
    If Multiplier > 0 Then
        For i = 0 To obj.ScaleHeight
            c = Abs(Max * Sin(i / st)) + Min
            obj.Line (0, i)-(obj.ScaleWidth, i), RGB(c + r, c + g, c + b)
        Next
    Else
        For i = 0 To obj.ScaleHeight
            c = Abs(Max * Cos(i / st)) + Min
            obj.Line (0, i)-(obj.ScaleWidth, i), RGB(c + r, c + g, c + b)
        Next
    End If
End Sub

Private Sub Form_Load()
    'Fill in the picture box
    GradientFill Picture1, 0, 255, 55, 55, 55, 1.3
    
    'Make special frame
    SetWindowLong Me.hwnd, GWL_STYLE, GetWindowLong(Me.hwnd, GWL_STYLE) + WS_DLGFRAME
    
    'Shave the corners off the gold control!
    MakeRoundRect MetalCB1(1), 40, 40
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub

Private Sub Form_Resize()
    GradientFill Me, 0, 255, 33, 33, 33, 1.3
    MetalCB2.Left = Me.ScaleWidth - 1.5 * MetalCB2.Width
End Sub

Private Sub MetalCB1_Click(Index As Integer)
    Label1.Caption = MetalCB1(Index).Caption
    With MetalCB1(Index)
        GradientFill Picture1, .MinShade, .maxshade, .red, .green, .blue, -.Multiplier
    End With
End Sub

Private Sub MakeRoundRect(ctl As Control, rx&, ry&)
    On Error Resume Next
    Dim r1 As Long
    With ctl
        r1 = CreateRoundRectRgn(0, 0, .Width, .Height, rx, ry)
        SetWindowRgn .hwnd, r1, True
    End With
    DeleteObject r1
End Sub

Private Sub MetalCB2_Click()
    End
End Sub
