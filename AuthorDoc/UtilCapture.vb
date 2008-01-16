Option Strict Off
Option Explicit On
Module UtilCapture
	'Copyright 2000 by AQUA TERRA Consultants
	'
	' Visual Basic 4.0 16/32 Capture Routines
	'
	' This module contains several routines for capturing windows into a
	' picture.  All the routines work on both 16 and 32 bit Windows
	' platforms.
	' The routines also have palette support.
	'
	' CreateBitmapPicture - Creates a picture object from a bitmap and
	' palette
	' CaptureWindow - Captures any window given a window handle
	' CaptureActiveWindow - Captures the active window on the desktop
	' CaptureForm - Captures the entire form
	' CaptureClient - Captures the client area of a form
	' CaptureScreen - Captures the entire screen
	' PrintPictureToFitPage - prints any picture as big as possible on
	' the page
	'
	' NOTES
	'    - No error trapping is included in these routines
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Private Structure PALETTEENTRY
		Dim peRed As Byte
		Dim peGreen As Byte
		Dim peBlue As Byte
		Dim peFlags As Byte
	End Structure
	Private Structure LOGPALETTE
		Dim palVersion As Short
		Dim palNumEntries As Short
		<VBFixedArray(255)> Dim palPalEntry() As PALETTEENTRY ' Enough for 256 colors
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			ReDim palPalEntry(255)
		End Sub
	End Structure
	Private Structure GUID
		Dim Data1 As Integer
		Dim Data2 As Short
		Dim Data3 As Short
		<VBFixedArray(7)> Dim Data4() As Byte
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			ReDim Data4(7)
		End Sub
	End Structure
	
	Private Const RASTERCAPS As Integer = 38
	Private Const RC_PALETTE As Integer = &H100s
	Private Const SIZEPALETTE As Integer = 104
	Private Structure RECT
		'UPGRADE_NOTE: Left was upgraded to Left_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Left_Renamed As Integer
		Dim Top As Integer
		'UPGRADE_NOTE: Right was upgraded to Right_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Right_Renamed As Integer
		Dim Bottom As Integer
	End Structure
	
	Private Declare Function GetWindowTextLength Lib "user32"  Alias "GetWindowTextLengthA"(ByVal hwnd As Integer) As Integer
	Private Declare Function GetWindowText Lib "user32"  Alias "GetWindowTextA"(ByVal hwnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
	
	Private Declare Function GetForegroundWindow Lib "user32" () As Integer
	Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Integer) As Integer
	Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer) As Integer
	Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC As Integer, ByVal iCapabilitiy As Integer) As Integer
	'UPGRADE_WARNING: Structure PALETTEENTRY may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetSystemPaletteEntries Lib "GDI32" (ByVal hDC As Integer, ByVal wStartIndex As Integer, ByVal wNumEntries As Integer, ByRef lpPaletteEntries As PALETTEENTRY) As Integer
	'UPGRADE_WARNING: Structure LOGPALETTE may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function CreatePalette Lib "GDI32" (ByRef lpLogPalette As LOGPALETTE) As Integer
	Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Integer, ByVal hObject As Integer) As Integer
	Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Integer, ByVal XDest As Integer, ByVal YDest As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hDCSrc As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal dwRop As Integer) As Integer
	Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Integer) As Integer
	Private Declare Function SelectPalette Lib "GDI32" (ByVal hDC As Integer, ByVal hPalette As Integer, ByVal bForceBackground As Integer) As Integer
	Private Declare Function RealizePalette Lib "GDI32" (ByVal hDC As Integer) As Integer
	Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Integer) As Integer
	Private Declare Function GetDC Lib "user32" (ByVal hwnd As Integer) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Integer, ByRef lpRect As RECT) As Integer
	Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Integer, ByVal hDC As Integer) As Integer
	Private Declare Function GetDesktopWindow Lib "user32" () As Integer
	Private Structure PicBmp
		Dim Size As Integer
		Dim Type As Integer
		Dim hBmp As Integer
		Dim hPal As Integer
		Dim Reserved As Integer
	End Structure
	'UPGRADE_WARNING: Structure IPicture may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	'UPGRADE_WARNING: Structure GUID may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	'UPGRADE_WARNING: Structure PicBmp may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (ByRef PicDesc As PicBmp, ByRef RefIID As GUID, ByVal fPictureOwnsHandle As Integer, ByRef IPic As System.Drawing.Image) As Integer
	
	Public Function GetActiveWindowRect() As RECT
		Dim hWndActive As Integer
		Dim RectActive As RECT
		' Get a handle to the active/foreground window
		hWndActive = GetForegroundWindow()
		' Get the dimensions of the window
		GetWindowRect(hWndActive, RectActive)
		'UPGRADE_WARNING: Couldn't resolve default property of object GetActiveWindowRect. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetActiveWindowRect = RectActive
	End Function
	
	Public Function GetActiveWindowText() As String
		Dim WinTextLength As Integer
		Dim WinText As String
		Dim hWndActive As Integer
		Dim i As Integer
		
		hWndActive = GetForegroundWindow()
		
		WinTextLength = GetWindowTextLength(hWndActive) + 1
		WinText = New String(Chr(0), WinTextLength + 1)
		GetWindowText(hWndActive, WinText, WinTextLength)
		i = InStr(1, WinText, Chr(0))
		If i <> 0 Then WinText = Left(WinText, i - 1)
		GetActiveWindowText = WinText
	End Function
	
	'
	' CreateBitmapPicture
	'    - Creates a bitmap type Picture object from a bitmap and palette
	'
	' hBmp
	'    - Handle to a bitmap
	'
	' hPal
	'    - Handle to a Palette
	'    - Can be null if the bitmap doesn't use a palette
	'
	' Returns
	'    - Returns a Picture object containing the bitmap
	Public Function CreateBitmapPicture(ByVal hBmp As Integer, ByVal hPal As Integer) As System.Drawing.Image
		
		Dim r As Integer
		Dim Pic As PicBmp
		' IPicture requires a reference to "Standard OLE Types"
		Dim IPic As System.Drawing.Image
		'UPGRADE_WARNING: Arrays in structure IID_IDispatch may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim IID_IDispatch As GUID
		' Fill in with IDispatch Interface ID
		With IID_IDispatch
			.Data1 = &H20400
			.Data4(0) = &HC0s
			.Data4(7) = &H46s
		End With
		' Fill Pic with necessary parts
		With Pic
			.Size = Len(Pic) ' Length of structure
			'UPGRADE_ISSUE: Constant vbPicTypeBitmap was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
			.Type = vbPicTypeBitmap ' Type of Picture (bitmap)
			.hBmp = hBmp ' Handle to bitmap
			.hPal = hPal ' Handle to palette (may be null)
		End With
		' Create Picture object
		r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
		' Return the new Picture object
		CreateBitmapPicture = IPic
	End Function
	'
	' CaptureWindow
	'    - Captures any portion of a window
	'
	' hWndSrc
	'    - Handle to the window to be captured
	'
	' Client
	'    - If True CaptureWindow captures from the client area of the
	'      window
	'    - If False CaptureWindow captures from the entire window      '
	' LeftSrc, TopSrc, WidthSrc, HeightSrc
	'    - Specify the portion of the window to capture
	'    - Dimensions need to be specified in pixels
	'
	' Returns
	'    - Returns a Picture object containing a bitmap of the specified
	'      portion of the window that was captured
	Public Function CaptureWindow(ByVal hWndSrc As Integer, ByVal Client As Boolean, ByVal LeftSrc As Integer, ByVal TopSrc As Integer, ByVal WidthSrc As Integer, ByVal HeightSrc As Integer) As System.Drawing.Image
		
		
		Dim hDCMemory As Integer
		Dim hBmp As Integer
		Dim hBmpPrev As Integer
		Dim r As Integer
		Dim hDCSrc As Integer
		Dim hPal As Integer
		Dim hPalPrev As Integer
		Dim RasterCapsScrn As Integer
		Dim HasPaletteScrn As Integer
		Dim PaletteSizeScrn As Integer
		
		'UPGRADE_WARNING: Arrays in structure LogPal may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim LogPal As LOGPALETTE
		' Depending on the value of Client get the proper device context
		If Client Then
			hDCSrc = GetDC(hWndSrc) ' Get device context for client area
		Else
			hDCSrc = GetWindowDC(hWndSrc) ' Get device context for entire window
		End If
		' Create a memory device context for the copy process
		hDCMemory = CreateCompatibleDC(hDCSrc)
		' Create a bitmap and place it in the memory DC
		hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
		hBmpPrev = SelectObject(hDCMemory, hBmp)
		' Get screen properties
		RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster capabilities
		HasPaletteScrn = RasterCapsScrn And RC_PALETTE ' Palette support
		PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of palette
		
		' If the screen has a palette make a copy and realize it
		If HasPaletteScrn And (PaletteSizeScrn = 256) Then
			' Create a copy of the system palette
			LogPal.palVersion = &H300s
			LogPal.palNumEntries = 256
			r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
			hPal = CreatePalette(LogPal)
			' Select the new palette into the memory DC and realize it
			hPalPrev = SelectPalette(hDCMemory, hPal, 0)
			r = RealizePalette(hDCMemory)
		End If
		' Copy the on-screen image into the memory DC
		'UPGRADE_ISSUE: Constant vbSrcCopy was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)
		' Remove the new copy of the  on-screen image
		hBmp = SelectObject(hDCMemory, hBmpPrev)
		' If the screen has a palette get back the palette that was
		' selected in previously
		If HasPaletteScrn And (PaletteSizeScrn = 256) Then
			hPal = SelectPalette(hDCMemory, hPalPrev, 0)
		End If
		' Release the device context resources back to the system
		r = DeleteDC(hDCMemory)
		r = ReleaseDC(hWndSrc, hDCSrc)
		' Call CreateBitmapPicture to create a picture object from the
		' bitmap and palette handles.  Then return the resulting picture
		' object.
		CaptureWindow = CreateBitmapPicture(hBmp, hPal)
	End Function
	'
	' CaptureScreen
	'    - Captures the entire screen
	'
	' Returns
	'    - Returns a Picture object containing a bitmap of the screen
	Public Function CaptureScreen() As System.Drawing.Image
		Dim hWndScreen As Integer
		' Get a handle to the desktop window
		hWndScreen = GetDesktopWindow()
		' Call CaptureWindow to capture the entire desktop give the handle
		' and return the resulting Picture object
		CaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) \ VB6.TwipsPerPixelX, VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) \ VB6.TwipsPerPixelY)
	End Function
	'
	' CaptureForm
	'    - Captures an entire form including title bar and border
	'
	' frmSrc
	'    - The Form object to capture
	'
	' Returns
	'    - Returns a Picture object containing a bitmap of the entire
	'      form
	Public Function CaptureForm(ByRef frmSrc As System.Windows.Forms.Form) As System.Drawing.Image
		' Call CaptureWindow to capture the entire form given it's window
		' handle and then return the resulting Picture object
		'UPGRADE_ISSUE: Constant vbPixels was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: Constant vbTwips was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: Form method frmSrc.ScaleY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: Form method frmSrc.ScaleX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		CaptureForm = CaptureWindow(frmSrc.Handle.ToInt32, False, 0, 0, frmSrc.ScaleX(VB6.PixelsToTwipsX(frmSrc.Width), vbTwips, vbPixels), frmSrc.ScaleY(VB6.PixelsToTwipsY(frmSrc.Height), vbTwips, vbPixels))
	End Function
	'
	' CaptureClient
	'    - Captures the client area of a form
	'
	' frmSrc
	'    - The Form object to capture
	'
	' Returns
	'    - Returns a Picture object containing a bitmap of the form's
	' client area
	'
	Public Function CaptureClient(ByRef frmSrc As System.Windows.Forms.Form) As System.Drawing.Image
		' Call CaptureWindow to capture the client area of the form given
		' it's window handle and return the resulting Picture object
		'UPGRADE_ISSUE: Constant vbPixels was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: Form property frmSrc.ScaleMode is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8027179A-CB3B-45C0-9863-FAA1AF983B59"'
		'UPGRADE_ISSUE: Form method frmSrc.ScaleY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: Form method frmSrc.ScaleX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		CaptureClient = CaptureWindow(frmSrc.Handle.ToInt32, True, 0, 0, frmSrc.ScaleX(VB6.PixelsToTwipsX(frmSrc.ClientRectangle.Width), frmSrc.ScaleMode, vbPixels), frmSrc.ScaleY(VB6.PixelsToTwipsY(frmSrc.ClientRectangle.Height), frmSrc.ScaleMode, vbPixels))
	End Function
	'
	' CaptureActiveWindow
	'    - Captures the currently active window on the screen      '
	' Returns
	'    - Returns a Picture object containing a bitmap of the active
	'      window
	'
	Public Function CaptureActiveWindow() As System.Drawing.Image
		Dim hWndActive As Integer
		Dim r As Integer
		Dim RectActive As RECT
		' Get a handle to the active/foreground window
		hWndActive = GetForegroundWindow()
		' Get the dimensions of the window
		r = GetWindowRect(hWndActive, RectActive)
		' Call CaptureWindow to capture the active window given it's
		' handle and return the Resulting Picture object
		CaptureActiveWindow = CaptureWindow(hWndActive, False, 0, 0, RectActive.Right_Renamed - RectActive.Left_Renamed, RectActive.Bottom - RectActive.Top)
	End Function
	'
	' PrintPictureToFitPage
	'    - Prints a Picture object as big as possible
	'
	' Prn
	'    - Destination Printer object
	'
	' Pic
	'    - Source Picture object
	'UPGRADE_ISSUE: VB.Printer type was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
	Public Sub PrintPictureToFitPage(ByRef Prn As Object, ByRef Pic As System.Drawing.Image)
		Const vbHiMetric As Short = 8
		Dim PicRatio As Double
		Dim PrnWidth As Double
		Dim PrnHeight As Double
		Dim PrnRatio As Double
		Dim PrnPicWidth As Double
		Dim PrnPicHeight As Double
		' Determine if picture should be printed in landscape or portrait
		' and set the orientation
		'UPGRADE_ISSUE: Picture property Pic.Width was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: Picture property Pic.Height was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		If Pic.Height >= Pic.Width Then
			'UPGRADE_ISSUE: Constant vbPRORPortrait was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
			'UPGRADE_ISSUE: Printer property Prn.Orientation was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			Prn.Orientation = vbPRORPortrait ' Taller than wide
		Else
			'UPGRADE_ISSUE: Constant vbPRORLandscape was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
			'UPGRADE_ISSUE: Printer property Prn.Orientation was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			Prn.Orientation = vbPRORLandscape ' Wider than tall
		End If
		' Calculate device independent Width to Height ratio for picture
		'UPGRADE_ISSUE: Picture property Pic.Height was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: Picture property Pic.Width was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		PicRatio = Pic.Width / Pic.Height
		' Calculate the dimentions of the printable area in HiMetric
		'UPGRADE_ISSUE: Printer property Prn.ScaleMode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		'UPGRADE_ISSUE: Printer property Prn.ScaleWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		'UPGRADE_ISSUE: Printer method Prn.ScaleX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		PrnWidth = Prn.ScaleX(Prn.ScaleWidth, Prn.ScaleMode, vbHiMetric)
		'UPGRADE_ISSUE: Printer property Prn.ScaleMode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		'UPGRADE_ISSUE: Printer property Prn.ScaleHeight was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		'UPGRADE_ISSUE: Printer method Prn.ScaleY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		PrnHeight = Prn.ScaleY(Prn.ScaleHeight, Prn.ScaleMode, vbHiMetric)
		' Calculate device independent Width to Height ratio for printer
		PrnRatio = PrnWidth / PrnHeight
		' Scale the output to the printable area
		If PicRatio >= PrnRatio Then
			' Scale picture to fit full width of printable area
			'UPGRADE_ISSUE: Printer property Prn.ScaleMode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			'UPGRADE_ISSUE: Printer method Prn.ScaleX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			PrnPicWidth = Prn.ScaleX(PrnWidth, vbHiMetric, Prn.ScaleMode)
			'UPGRADE_ISSUE: Printer property Prn.ScaleMode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			'UPGRADE_ISSUE: Printer method Prn.ScaleY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			PrnPicHeight = Prn.ScaleY(PrnWidth / PicRatio, vbHiMetric, Prn.ScaleMode)
		Else
			' Scale picture to fit full height of printable area
			'UPGRADE_ISSUE: Printer property Prn.ScaleMode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			'UPGRADE_ISSUE: Printer method Prn.ScaleY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			PrnPicHeight = Prn.ScaleY(PrnHeight, vbHiMetric, Prn.ScaleMode)
			'UPGRADE_ISSUE: Printer property Prn.ScaleMode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			'UPGRADE_ISSUE: Printer method Prn.ScaleX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			PrnPicWidth = Prn.ScaleX(PrnHeight * PicRatio, vbHiMetric, Prn.ScaleMode)
		End If
		' Print the picture using the PaintPicture method
		'UPGRADE_ISSUE: Printer method Prn.PaintPicture was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		Prn.PaintPicture(Pic, 0, 0, PrnPicWidth, PrnPicHeight)
	End Sub
End Module