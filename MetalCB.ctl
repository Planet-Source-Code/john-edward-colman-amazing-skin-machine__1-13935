VERSION 5.00
Begin VB.UserControl MetalCB 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1950
   ClipBehavior    =   0  'None
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   50
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   130
   ToolboxBitmap   =   "MetalCB.ctx":0000
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MetalCB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   675
      TabIndex        =   0
      Top             =   225
      Width           =   600
   End
End
Attribute VB_Name = "MetalCB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'---------------------------------------------------------------------
' Hey you !!!!
' This code was originally written by John Colman
' If you use this code, give me a credit.
' (john_colman@hotmail.com)
'---------------------------------------------------------------------

Dim pMin As Integer
Dim pMax As Integer
Dim pRed As Integer
Dim pGreen As Integer
Dim pBlue As Integer
Dim pMultiplier As Double

Dim LabelTop As Integer

Dim pClickDepth As Integer
Public Event Click()

Private Sub Redraw(Optional DrawClicked As Boolean = False)
    On Error Resume Next
    Dim i As Long
    Dim Color As Integer
    Dim iStep As Double
    Dim MaxVal As Integer
    
    MaxVal = pMax - pMin
    iStep = UserControl.ScaleHeight / 3.142 * Abs(pMultiplier)
    
    'This code may seem redundant but its faster this way
    If Multiplier > 0 Then
        If DrawClicked Then
            For i = 0 To UserControl.ScaleHeight
                Color = Abs(MaxVal * Sin(i / iStep)) + pMin
                UserControl.Line (0, i + pClickDepth)-(UserControl.ScaleWidth, i + pClickDepth), RGB(Color + pRed, Color + pGreen, Color + pBlue)
            Next
        Else
            For i = 0 To UserControl.ScaleHeight
                Color = Abs(MaxVal * Sin(i / iStep)) + pMin
                UserControl.Line (0, i)-(UserControl.ScaleWidth, i), RGB(Color + pRed, Color + pGreen, Color + pBlue)
            Next
        End If
    Else
        If DrawClicked Then
            For i = 0 To UserControl.ScaleHeight
                Color = Abs(MaxVal * Cos(i / iStep)) + pMin
                UserControl.Line (0, i + pClickDepth)-(UserControl.ScaleWidth, i + pClickDepth), RGB(Color + pRed, Color + pGreen, Color + pBlue)
            Next
        Else
            For i = 0 To UserControl.ScaleHeight
                Color = Abs(MaxVal * Cos(i / iStep)) + pMin
                UserControl.Line (0, i)-(UserControl.ScaleWidth, i), RGB(Color + pRed, Color + pGreen, Color + pBlue)
            Next
        End If
    End If
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        PressButton True
    End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    PressButton False
    RaiseEvent Click
End Sub


Private Sub UserControl_InitProperties()
    Label1.Caption = "MetalCB"
    pRed = 33
    pGreen = 33
    pBlue = 33
    pMultiplier = -1.5
    pMax = 255
    pMin = 0
    pClickDepth = 3
    Redraw
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        PressButton True
    End If
End Sub

Private Sub PressButton(Pressed As Boolean)
    Redraw Pressed
    Label1.Top = LabelTop + IIf(Pressed, pClickDepth, 0)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    PressButton False
    RaiseEvent Click
End Sub

Private Sub UserControl_Resize()
    Redraw
    Label1.Left = (UserControl.ScaleWidth - Label1.Width) / 2
    LabelTop = (UserControl.ScaleHeight - Label1.Height) / 2
    PressButton False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    PropBag.WriteProperty "FontSize", Label1.Font.Size
    PropBag.WriteProperty "FontCharset", Label1.Font.Charset
    PropBag.WriteProperty "FontBold", Label1.Font.Bold
    PropBag.WriteProperty "FontItalic", Label1.Font.Italic
    PropBag.WriteProperty "FontName", Label1.Font.Name
    PropBag.WriteProperty "FontStrike", Label1.Font.Strikethrough
    PropBag.WriteProperty "FontUnder", Label1.Font.Underline
    PropBag.WriteProperty "FontWeight", Label1.Font.Weight
    
    PropBag.WriteProperty "Caption", Label1.Caption, "MetalCB"
    PropBag.WriteProperty "ForeColor", Label1.ForeColor, RGB(0, 0, 0)
    PropBag.WriteProperty "Font", Label1.Font
    PropBag.WriteProperty "Red", pRed, 33
    PropBag.WriteProperty "Green", pGreen, 33
    PropBag.WriteProperty "Blue", pBlue, 33
    PropBag.WriteProperty "Multiplier", pMultiplier, -1.5
    PropBag.WriteProperty "ShadeMax", pMax, 255
    PropBag.WriteProperty "ShadeMin", pMin, 0
    PropBag.WriteProperty "ClickDepth", pClickDepth, 3
    PropBag.WriteProperty "BorderStyle", UserControl.BorderStyle, 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Label1.Font.Size = PropBag.ReadProperty("FontSize", Label1.Font.Size)
    Label1.Font.Charset = PropBag.ReadProperty("FontCharset", Label1.Font.Charset)
    Label1.Font.Bold = PropBag.ReadProperty("FontBold", Label1.Font.Bold)
    Label1.Font.Italic = PropBag.ReadProperty("FontItalic", Label1.Font.Italic)
    Label1.Font.Name = PropBag.ReadProperty("FontName", Label1.Font.Name)
    Label1.Font.Strikethrough = PropBag.ReadProperty("FontStrike", Label1.Font.Strikethrough)
    Label1.Font.Underline = PropBag.ReadProperty("FontUnder", Label1.Font.Underline)
    Label1.Font.Weight = PropBag.ReadProperty("FontWeight", Label1.Font.Weight)
    
    Label1.Caption = PropBag.ReadProperty("Caption", "MetalCB")
    Label1.ForeColor = PropBag.ReadProperty("ForeColor", RGB(0, 0, 0))
    pRed = PropBag.ReadProperty("Red", 33)
    pGreen = PropBag.ReadProperty("Green", 33)
    pBlue = PropBag.ReadProperty("Blue", 33)
    pMultiplier = PropBag.ReadProperty("Multiplier", -1.5)
    pMax = PropBag.ReadProperty("ShadeMax", 255)
    pMin = PropBag.ReadProperty("ShadeMin", 0)
    pClickDepth = PropBag.ReadProperty("ClickDepth", 3)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    Redraw
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption.VB_UserMemId = -518
Attribute Caption.VB_MemberFlags = "200"
    Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal vNewValue As String)
    PropertyChanged "Caption"
    Label1.Caption = vNewValue
    UserControl_Resize
End Property

Public Property Get Blue() As Integer
Attribute Blue.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Blue = pBlue
End Property

Public Property Let Blue(ByVal vNewColor As Integer)
    If vNewColor > 255 Or vNewColor < 0 Then
        pBlue = 0
    Else
        pBlue = vNewColor
    End If
    PropertyChanged "Blue"
    Redraw
End Property

Public Property Get Red() As Integer
Attribute Red.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Red = pRed
End Property

Public Property Let Red(ByVal vNewColor As Integer)
    If vNewColor > 255 Or vNewColor < 0 Then
        pRed = 0
    Else
        pRed = vNewColor
    End If
    PropertyChanged "Red"
    Redraw
End Property

Public Property Get Green() As Integer
Attribute Green.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Green = pGreen
End Property

Public Property Let Green(ByVal vNewColor As Integer)
    If vNewColor > 255 Or vNewColor < 0 Then
        pGreen = 0
    Else
        pGreen = vNewColor
    End If
    PropertyChanged "Green"
    Redraw
End Property

Public Property Get Multiplier() As Double
Attribute Multiplier.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Multiplier = pMultiplier
End Property

Public Property Let Multiplier(ByVal vNewValue As Double)
    pMultiplier = vNewValue
    PropertyChanged "Multiplier"
    Redraw
End Property

Public Property Get MaxShade() As Integer
Attribute MaxShade.VB_ProcData.VB_Invoke_Property = ";Appearance"
    MaxShade = pMax
End Property

Public Property Let MaxShade(ByVal vNewMax As Integer)
    If vNewMax > 255 Or vNewMax < 0 Then
        pMax = 255
    Else
        pMax = vNewMax
    End If
    PropertyChanged "MaxShade"
    Redraw
End Property

Public Property Get MinShade() As Integer
Attribute MinShade.VB_ProcData.VB_Invoke_Property = ";Appearance"
    MinShade = pMin
End Property

Public Property Let MinShade(ByVal vNewMin As Integer)
    If vNewMin > 255 Or vNewMin < 0 Then
        pMin = 0
    Else
        pMin = vNewMin
    End If
    PropertyChanged "MinShade"
    Redraw
End Property

Public Property Get ClickDepth() As Integer
Attribute ClickDepth.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ClickDepth = pClickDepth
End Property

Public Property Let ClickDepth(ByVal vNewDepth As Integer)
    pClickDepth = vNewDepth
    PropertyChanged "ClickDepth"
End Property

Public Property Get Font() As stdole.StdFont
Attribute Font.VB_ProcData.VB_Invoke_Property = "StandardFont;Font"
Attribute Font.VB_UserMemId = -512
    Set Font = Label1.Font
End Property

Public Property Let Font(ByVal vNewFont As stdole.StdFont)
    Set Label1.Font = vNewFont
    UserControl_Resize
    PropertyChanged "FontSize"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = Label1.ForeColor
End Property

Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
     Label1.ForeColor = vNewValue
     PropertyChanged "ForeColor"
End Property

Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BorderStyle.VB_UserMemId = -504
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal vNewValue As Integer)
    UserControl.BorderStyle = vNewValue
    PropertyChanged "BorderStyle"
End Property

Public Property Get hWnd() As Variant
Attribute hWnd.VB_UserMemId = -515
Attribute hWnd.VB_MemberFlags = "400"
    hWnd = UserControl.hWnd
End Property
