Attribute VB_Name = "modTransparent"
Option Explicit

Private Const RGN_AND = 1
Private Const RGN_OR = 2
Private Const RGN_XOR = 3
Private Const RGN_DIFF = 4
Private Const RGN_COPY = 5
Private Declare Function CreateRectRgn _
                Lib "gdi32" (ByVal X1 As Long, _
                             ByVal Y1 As Long, _
                             ByVal X2 As Long, _
                             ByVal Y2 As Long) As Long
Private Declare Function CombineRgn _
                Lib "gdi32" (ByVal hDestRgn As Long, _
                             ByVal hSrcRgn1 As Long, _
                             ByVal hSrcRgn2 As Long, _
                             ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn _
                Lib "user32" (ByVal hWnd As Long, _
                              ByVal hRgn As Long, _
                              ByVal bRedraw As Boolean) As Long
Private Declare Function SelectObject _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal hObject As Long) As Long
Private Declare Function DeleteObject _
                Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC _
                Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleDC _
                Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetPixel _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal x As Long, _
                             ByVal y As Long) As Long

Public Function MakeRegion(ByRef frm As Form, _
                           ByVal TrnsColor As Long) As Long
    frm.BorderStyle = 0
    Dim ScaleSize As Long
    Dim Width, Height As Long
    Dim rgnMain As Long
    Dim x, y As Long
    Dim rgnPixel As Long
    Dim RGBColor As Long
    Dim dcMain As Long
    Dim bmpMain As Long
    ScaleSize = frm.ScaleMode
    frm.ScaleMode = 3
    Width = frm.ScaleX(frm.Picture.Width, vbHimetric, vbPixels)
    Height = frm.ScaleY(frm.Picture.Height, vbHimetric, vbPixels)
    frm.Width = Width * Screen.TwipsPerPixelX
    frm.Height = Height * Screen.TwipsPerPixelY
    rgnMain = CreateRectRgn(0, 0, Width, Height)
    dcMain = CreateCompatibleDC(frm.hDC)
    bmpMain = SelectObject(dcMain, frm.Picture.Handle)

    For y = 0 To Height
        For x = 0 To Width
            RGBColor = GetPixel(dcMain, x, y)

            If RGBColor = TrnsColor Then
                rgnPixel = CreateRectRgn(x, y, x + 1, y + 1)
                CombineRgn rgnMain, rgnMain, rgnPixel, RGN_XOR
                DeleteObject rgnPixel
            End If

        Next x
    Next y

    SelectObject dcMain, bmpMain
    DeleteDC dcMain
    DeleteObject bmpMain

    If rgnMain <> 0 Then
        SetWindowRgn frm.hWnd, rgnMain, True
        MakeRegion = rgnMain
    End If

    frm.ScaleMode = ScaleSize
End Function

