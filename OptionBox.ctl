VERSION 5.00
Begin VB.UserControl OptionBox 
   AutoRedraw      =   -1  'True
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1800
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   20
   ScaleMode       =   0  'User
   ScaleWidth      =   120
End
Attribute VB_Name = "OptionBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function StretchDIBits& Lib "gdi32" (ByVal hdc&, ByVal X&, ByVal Y&, ByVal dx&, ByVal dy&, ByVal SrcX&, ByVal SrcY&, ByVal Srcdx&, ByVal Srcdy&, Bits As Any, BInf As Any, ByVal Usage&, ByVal Rop&)
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Private HiLite&
Private HiLite2&
Private LoLite&
Private Greyed&
Private Shadow&
Private mEnabled As Boolean
Private cw&, ch&, X2&
Private mCurrentState As Boolean
Private mShadowLine As Boolean
Private mBackStyle As Integer
Private mCaption As String

Private xONOFF&
Private wONOFF&

Public Event Click()
Public Event DblClick()
Public Property Get Caption() As String
    Caption = mCaption
End Property

Public Property Let Caption(ByVal vNewValue As String)
    mCaption = vNewValue
    PropertyChanged Caption
    DrawControl
End Property

Public Property Let Value(bVal As Boolean)
mCurrentState = bVal
PropertyChanged Value
DrawControl
End Property
Public Property Get Value() As Boolean
Value = mCurrentState
End Property

Public Property Let BackStyle(bVal As Integer)
If bVal < 1 Or bVal > 2 Or bVal = mBackStyle Then Exit Property
mBackStyle = bVal
PropertyChanged BackStyle
DrawControl
End Property
Public Property Get BackStyle() As Integer
BackStyle = mBackStyle
End Property

Public Property Let ShadowLine(bVal As Boolean)
mShadowLine = bVal
PropertyChanged ShadowLine
DrawControl
End Property
Public Property Get ShadowLine() As Boolean
ShadowLine = mShadowLine
End Property
Private Sub SplitRGB(ByVal clr&, r&, G&, B&)
    r = clr And &HFF: G = (clr \ &H100&) And &HFF: B = (clr \ &H10000) And &HFF
End Sub
Private Sub Gradient(dc&, X&, Y&, dx&, dy&, ByVal c1&, ByVal c2&, v As Boolean)
Dim r1&, G1&, B1&, r2&, G2&, B2&, B() As Byte
Dim i&, lR!, lG!, lB!, dR!, dG!, dB!, BI&(9), xx&, yy&, dd&, hRPen&
    If dx = 0 Or dy = 0 Then Exit Sub
    If v Then xx = 1: yy = dy: dd = dy Else xx = dx: yy = 1: dd = dx
    SplitRGB c1, r1, G1, B1: SplitRGB c2, r2, G2, B2: ReDim B(dd * 4 - 1)
    dR = (r2 - r1) / (dd - 1): lR = r1: dG = (G2 - G1) / (dd - 1): lG = G1: dB = (B2 - B1) / (dd - 1): lB = B1
    For i = 0 To (dd - 1) * 4 Step 4: B(i + 2) = lR: lR = lR + dR: B(i + 1) = lG: lG = lG + dG: B(i) = lB: lB = lB + dB: Next
    BI(0) = 40: BI(1) = xx: BI(2) = -yy: BI(3) = 2097153: StretchDIBits dc, X, Y, dx, dy, 0, 0, xx, yy, B(0), BI(0), 0, vbSrcCopy
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_InitProperties()
    Value = False
End Sub
Private Sub UserControl_Paint()
DrawControl
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
         mCurrentState = .ReadProperty("Value", False)
         mShadowLine = .ReadProperty("ShadowLine", False)
         mBackStyle = .ReadProperty("BackStyle", False)
         mCaption = .ReadProperty("Caption", "ToggleBox")
    End With
End Sub

Private Sub UserControl_Show()
DrawControl
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Value", mCurrentState
        .WriteProperty "ShadowLine", mShadowLine
        .WriteProperty "BackStyle", mBackStyle
        .WriteProperty "Caption", mCaption
    End With
End Sub
Sub DrawLine(ByRef dc&, X1&, Y1&, X2&, Y2&, c&)
Dim p&, Pt As POINTAPI
    p = CreatePen(0, 1, c): DeleteObject SelectObject(dc, p)
    Pt.X = X1: Pt.Y = Y1
    MoveToEx dc, X1, Y1, Pt: LineTo dc, X2, Y2
    DeleteDC p
End Sub

Private Sub UserControl_DblClick()
mCurrentState = Not mCurrentState
DrawControl
RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
mCurrentState = Not mCurrentState
DrawControl
End Sub

Private Sub UserControl_Resize()
    
    'UserControl.Width = 700
    UserControl.Height = 270
    
    cw = UserControl.Width \ Screen.TwipsPerPixelX
    ch = UserControl.Height \ Screen.TwipsPerPixelY

    wONOFF = (700 \ Screen.TwipsPerPixelX)
    xONOFF = (UserControl.Width \ Screen.TwipsPerPixelX) - wONOFF

    X2 = cw \ 2
    
    DrawControl

End Sub
Private Sub UserControl_Initialize()
    
    UserControl.FontName = "Segoe UI"
    UserControl.FontSize = 8
    UserControl.FontBold = False
    
    mCurrentState = False
    mShadowLine = False
    mBackStyle = 1
    
    HiLite = RGB(215, 215, 215)
    HiLite2 = RGB(255, 255, 255)
    LoLite = RGB(165, 165, 165)
    Shadow = RGB(150, 150, 150)
    Greyed = RGB(190, 190, 190)
    
    
End Sub
Sub DrawControlON()

    If mBackStyle = 1 Then
    Gradient UserControl.hdc, 0, 0, cw, ch / 2, HiLite, HiLite2, True
    Gradient UserControl.hdc, 0, ch / 2, cw, ch, HiLite, LoLite, True
    Else
    Gradient UserControl.hdc, 0, 0, cw, ch, vbWhite, HiLite, True
    End If
    Gradient UserControl.hdc, xONOFF + 2, 2, (wONOFF \ 2 - 2), ch - 4, RGB(59, 109, 219), RGB(108, 168, 250), True
    
    
    If mShadowLine Then
        DrawLine UserControl.hdc, 0, ch - 1, cw - 1, ch - 1, Shadow
        DrawLine UserControl.hdc, cw - 1, 0, cw - 1, ch - 1, Shadow
    End If
    
    
    DrawLine UserControl.hdc, xONOFF + 2, 2, xONOFF + ((wONOFF \ 2)), 2, RGB(50, 100, 200)
    DrawLine UserControl.hdc, xONOFF + ((wONOFF \ 2)) - 1, 2, xONOFF + ((wONOFF \ 2)) - 1, ch - 2, RGB(117, 173, 255)

    SetTextColor UserControl.hdc, vbWhite
    TextOut UserControl.hdc, xONOFF + 3, 2, "ON", 2
    
    OutputCaption
    
    UserControl.Refresh

End Sub
Sub OutputCaption()
    SetTextColor UserControl.hdc, vbWhite
    TextOut UserControl.hdc, 3, 3, mCaption, Len(mCaption)
    SetTextColor UserControl.hdc, RGB(50, 50, 50)
    TextOut UserControl.hdc, 3, 2, mCaption, Len(mCaption)
End Sub
Sub DrawControl()
    If mCurrentState Then
        DrawControlON
    Else
        DrawControlOFF
    End If
End Sub
Sub DrawControlOFF()
    If mBackStyle = 1 Then
    Gradient UserControl.hdc, 0, 0, cw, ch / 2, HiLite, HiLite2, True
    Gradient UserControl.hdc, 0, ch / 2, cw, ch, HiLite, LoLite, True
    Else
    Gradient UserControl.hdc, 0, 0, cw, ch, vbWhite, HiLite, True
    End If
    
    If mShadowLine Then
        DrawLine UserControl.hdc, 0, ch - 1, cw - 1, ch - 1, Shadow
        DrawLine UserControl.hdc, cw - 1, 0, cw - 1, ch - 1, Shadow
    End If


    SetTextColor UserControl.hdc, vbWhite
    TextOut UserControl.hdc, cw - 26, 3, "OFF", 3
    
    SetTextColor UserControl.hdc, RGB(50, 50, 50)
    TextOut UserControl.hdc, cw - 26, 2, "OFF", 3
    
    OutputCaption

    UserControl.Refresh

End Sub

