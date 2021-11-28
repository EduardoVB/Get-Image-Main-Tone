Attribute VB_Name = "mColors"
Option Explicit

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type RGBQUAD
    rgbRed As Byte
    rgbGreen As Byte
    rgbBlue As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type PicBmp
    Size As Long
    Type As PictureTypeConstants
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type


Private Declare Sub ColorRGBToHLS Lib "shlwapi" (ByVal clrRGB As Long, ByRef pwHue As Integer, ByRef pwLuminance As Integer, ByRef pwSaturation As Integer)
Private Declare Function ColorHLSToRGB Lib "shlwapi" (ByVal wHue As Integer, ByVal wLuminance As Integer, ByVal wSaturation As Integer) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFOHEADER, ByVal wUsage As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC&) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC&) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle&, iPic As IPicture) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC&, pBitmapInfo As BITMAPINFO, ByVal un&, lplpVoid&, ByVal Handle&, ByVal dw&) As Long

Private Const DIB_RGB_COLORS As Long = 0

Public Enum ToneBrightness
    tbLight
    tbLighter
    tbDark
    tbDarker
    tbAutoLightDark
    tbAutoLighterDarker
    tbNormal
End Enum

Public Function PicMainTone(nPicture As StdPicture, Optional nBrightness As ToneBrightness = tbAutoLighterDarker) As OLE_COLOR
    Dim iBmp As BITMAP
    Dim iBMPiH As BITMAPINFOHEADER
    Dim iBits() As Byte
    Dim IID_IDispatch As GUID
    Dim Picb As PicBmp
    Dim BMI As BITMAPINFO
    
    Dim iTmpDC As Long
    Dim X As Long
    Dim iMax As Long
    Dim iColor As Long
    Dim iPicWidth As Long
    Dim iPicHeight As Long
    Dim iBitMap As Long
    Dim iOldObj As Long
    
    Dim iHues(240) As Single
    Dim iHueZones(6) As Single
    Dim iLum As Long
    Dim iSat(6) As Long
    Dim iSat_n(6) As Long
    
    Dim h As Integer
    Dim l As Integer
    Dim s As Integer
    Dim z As Long
    Dim iMax2 As Single
    Dim iMainHue As Byte
    Dim v As Single
    Dim f As Integer
    Dim t As Integer
    Dim iPic As StdPicture
    Dim iPicObj As IPicture
    Dim iScale As Single
    Dim hBmp As Long
    Dim hBmpPrev As Long
    Dim iWidth As Long
    Dim iHeight As Long
    Const cImgMaxSizeHimetric As Long = 5000
    Dim iResample As Boolean
    
    If nPicture.Type <> vbPicTypeBitmap Then
        iResample = True
    Else
        GetObject nPicture.Handle, Len(iBmp), iBmp
        If (nPicture.Height > cImgMaxSizeHimetric) Or (nPicture.Width > cImgMaxSizeHimetric) Or (iBmp.bmBitsPixel <> 32) Then
            iResample = True
        End If
    End If
    If iResample Then
        If (nPicture.Height > cImgMaxSizeHimetric) Or (nPicture.Width > cImgMaxSizeHimetric) Then
            If nPicture.Height > nPicture.Width Then
                iScale = cImgMaxSizeHimetric / nPicture.Height
            Else
                iScale = cImgMaxSizeHimetric / nPicture.Width
            End If
        Else
            iScale = 1
        End If
        iTmpDC = CreateCompatibleDC(0)
        Set iPicObj = nPicture
        iWidth = nPicture.Width * iScale * 1440 / 2540 / Screen.TwipsPerPixelX
        iHeight = nPicture.Height * iScale * 1440 / 2540 / Screen.TwipsPerPixelY
        
        With BMI.bmiHeader
            .biSize = Len(BMI.bmiHeader)
            .biWidth = iWidth
            .biHeight = iHeight
            .biPlanes = 1
            .biBitCount = 32
        End With
        Picb.hBmp = CreateDIBSection(0, BMI, 0, 0, 0, 0)
        hBmp = Picb.hBmp
        
        hBmpPrev = SelectObject(iTmpDC, hBmp)
        iPicObj.Render iTmpDC, 0, 0, iWidth, iHeight, 0, 0, nPicture.Width, nPicture.Height, 0&
        hBmp = SelectObject(iTmpDC, hBmpPrev)
        DeleteDC iTmpDC
        
        With IID_IDispatch
           .Data1 = &H20400
           .Data4(0) = &HC0
           .Data4(7) = &H46
         End With
    
        'fill Pic with necessary parts
        With Picb
           .Size = Len(Picb)         'Length of structure
           .Type = vbPicTypeBitmap  'Type of Picture (bitmap)
           .hBmp = hBmp             'Handle to bitmap
           .hPal = 0&               'Handle to palette (may be null)
         End With
    
        Call OleCreatePictureIndirect(Picb, IID_IDispatch, 1, iPic)
        GetObject iPic.Handle, Len(iBmp), iBmp
    Else
        Set iPic = nPicture
    End If
    
    With iBMPiH
        .biSize = Len(iBMPiH) '40
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = iBmp.bmWidth
        .biHeight = iBmp.bmHeight
        .biSizeImage = ((.biWidth * 4 + 3) And &HFFFFFFFC) * .biHeight
        iPicWidth = .biWidth
        iPicHeight = .biHeight
    End With
    
    ReDim iBits(Len(iBMPiH) + iBMPiH.biSizeImage)
    
    iTmpDC = CreateCompatibleDC(0)
    
    GetDIBits iTmpDC, iPic.Handle, 0, iBmp.bmHeight, iBits(0), iBMPiH, DIB_RGB_COLORS
    DeleteDC iTmpDC
        
    iMax = iBMPiH.biSizeImage - 1
    For X = 0 To iMax - 4 Step 4
        ColorRGBToHLS RGB(iBits(X + 2), iBits(X + 1), iBits(X)), h, l, s
        h = Int(h)
'        If h = 148 Then Stop
        'Debug.Print RGB(iBits(X + 2), iBits(X + 1), iBits(X))
        v = s / 240 * (1 - Abs(l - 120) / 120)
        iHues(h) = iHues(h) + v
        z = Int(h / 40)
        iHueZones(z) = iHueZones(z) + iHues(h)
        iLum = iLum + l
        iSat(z) = iSat(z) + s
        iSat_n(z) = iSat_n(z) + 1
    Next
    iLum = iLum / (iMax / 4)
    
    iMax2 = 0
    z = 0
    For X = 0 To UBound(iHueZones)
        If iHueZones(X) > iMax2 Then
            iMax2 = iHueZones(X)
            z = X
        End If
    Next
    
    iMax2 = 0
    iMainHue = 0
    f = (z + 0.5) * 40 - 20
    If f < 0 Then f = 0
    t = f + 40
    If t > 240 Then t = 240
    For X = f To t
        If iHues(X) > iMax2 Then
            iMax2 = iHues(X)
            iMainHue = X
        End If
    Next
    If iSat_n(z) = 0 Then iSat_n(z) = 1
    If iSat(z) < 60 Then iSat(z) = 60
    If iLum > 237 Then
        iSat(z) = 1
        iSat_n(z) = 1
    End If
    If nBrightness = tbNormal Then
        PicMainTone = ColorHLSToRGB(iMainHue, iLum, iSat(z) / iSat_n(z))
    ElseIf nBrightness = tbLight Then
        PicMainTone = ColorHLSToRGB(iMainHue, 200, iSat(z) / iSat_n(z))
    ElseIf nBrightness = tbDark Then
        PicMainTone = ColorHLSToRGB(iMainHue, 40, iSat(z) / iSat_n(z))
    ElseIf nBrightness = tbLighter Then
        PicMainTone = ColorHLSToRGB(iMainHue, 220, iSat(z) / iSat_n(z))
    ElseIf nBrightness = tbDarker Then
        PicMainTone = ColorHLSToRGB(iMainHue, 20, iSat(z) / iSat_n(z))
    ElseIf nBrightness = tbAutoLighterDarker Then
        If iLum > 100 Then
            PicMainTone = ColorHLSToRGB(iMainHue, 220, iSat(z) / iSat_n(z))
        Else
            PicMainTone = ColorHLSToRGB(iMainHue, 30, iSat(z) / iSat_n(z))
        End If
    Else
        If iLum > 100 Then
            PicMainTone = ColorHLSToRGB(iMainHue, 200, iSat(z) / iSat_n(z))
        Else
            PicMainTone = ColorHLSToRGB(iMainHue, 40, iSat(z) / iSat_n(z))
        End If
    End If
End Function

