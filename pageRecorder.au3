#include <Misc.au3>
#include <WinAPI.au3>
#include <GDIPlus.au3>
#include <WindowsConstants.au3>
#include <WinAPIGdi.au3>
#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include <String.au3>
#include <File.au3>

Global $screenScale = 2 ; Adjust screen scale factor as needed

Opt("MouseCoordMode", 1)  ; Use absolute coordinates for mouse position

Global $mainGUI
Global $pageRect, $buttonRect, $pageCount, $autoCrop
Global $hOverlay, $hOverlayDC, $hPageRectOverlay, $hPageRectOverlayDC, $hButtonRectOverlay, $hButtonRectOverlayDC
Global $avgCroppedWidth, $avgCroppedHeight, $numCroppedImage

Func Main()
    ; Create main GUI dialog
    $mainGUI = GUICreate("Document Screen Capture Tool", 300, 200)
    GUISetFont(10)
    Local $btnSelectPage = GUICtrlCreateButton("Page Area", 50, 10, 200, 30)
    Local $btnSelectButton = GUICtrlCreateButton("Turn Page Button", 50, 50, 200, 30)
    Local $btnStartCapture = GUICtrlCreateButton("Start Capture", 50, 160, 200, 30)
	Local $lblPageCount = GUICtrlCreateLabel("Page Num", 50, 95, 80, 20)
    Local $txtPageCount = GUICtrlCreateInput("1", 120, 90, 40, 25)
	Local $checkboxAutoCrop = GUICtrlCreateCheckbox("AutoCrop", 170, 90, 100, 25)

	GUICtrlSetState ($checkboxAutoCrop, $GUI_CHECKED)
    GUISetState(@SW_SHOW, $mainGUI)

    ; Handle main GUI events
    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE
                ExitLoop
            Case $btnSelectPage
                $pageRect = SelectRect("PageArea")
            Case $btnSelectButton
                $buttonRect = SelectRect("ButtonArea")
            Case $btnStartCapture
                $pageCount = GUICtrlRead($txtPageCount)
				$autoCrop = IsChecked($checkboxAutoCrop)
                If Not StringIsInt($pageCount) Or $pageCount < 1 Then
                    MsgBox($MB_ICONERROR, "Invalid Input", "Please enter a valid number of pages.")
                Else
					; Clear the overlay and DCs
					_WinAPI_ReleaseDC($hOverlay, $hOverlayDC)
					GUIDelete($hOverlay)
					$hOverlay = 0

					_WinAPI_ReleaseDC($hPageRectOverlay, $hPageRectOverlayDC)
					GUIDelete($hPageRectOverlay)
					$hPageRectOverlay = 0

					_WinAPI_ReleaseDC($hButtonRectOverlay, $hButtonRectOverlayDC)
					GUIDelete($hButtonRectOverlay)
					$hButtonRectOverlay = 0

                    StartCapture()
                EndIf
        EndSwitch
    WEnd
EndFunc

Func IsChecked($idControlID)
        Return BitAND(GUICtrlRead($idControlID), $GUI_CHECKED) = $GUI_CHECKED
EndFunc   ;==>_IsChecked


Func SelectRect($selctAreaName)
	If $selctAreaName = "PageArea" And $hPageRectOverlay <> 0 Then
		_WinAPI_ReleaseDC($hPageRectOverlay, $hPageRectOverlayDC)
		GUIDelete($hPageRectOverlay)
	EndIf

	If $selctAreaName = "ButtonArea" And $hButtonRectOverlay <> 0 Then
		_WinAPI_ReleaseDC($hButtonRectOverlay, $hButtonRectOverlayDC)
		GUIDelete($hButtonRectOverlay)
	EndIf

    ; Hide the main GUI temporarily
    GUISetState(@SW_HIDE, $mainGUI)

	If $hOverlay = 0 Then
	    ; Create full-screen overlay for selection
		$hOverlay = GUICreate("Overlay", @DesktopWidth * $screenScale, @DesktopHeight * $screenScale, 0, 0, $WS_POPUP, $WS_EX_LAYERED + $WS_EX_TOPMOST + $WS_EX_TOOLWINDOW)
		GUISetBkColor(0x000000)
		WinSetTrans("Overlay", "", 100) ; Set overlay transparency
		GUISetState(@SW_SHOW)
		$hOverlayDC = _WinAPI_GetDC($hOverlay)
	Else
		WinActivate("Overlay")
	EndIf

    ; Initialize GDI+ for drawing
    _GDIPlus_Startup()


    ; Load cross cursor
    Local $hCursor = _WinAPI_LoadCursor(0, $OCR_CROSS)
    Local $isDragging = False
    Local $startX, $startY, $endX, $endY

    ; Continuously set cross cursor until selection is completed
    While Not $isDragging
        _WinAPI_SetCursor($hCursor)

        ; Wait for mouse click to start selection
        If _IsPressed("01") Then
            Local $startPos = MouseGetPos()
            $startX = $startPos[0]
            $startY = $startPos[1]
            $isDragging = True
        EndIf
        Sleep(10)
    WEnd

    ; Track and draw selection rectangle
    While $isDragging
        _WinAPI_SetCursor($hCursor) ; Reinforce cross cursor during selection

        Local $currentPos = MouseGetPos()
        Local $currentX = $currentPos[0]
        Local $currentY = $currentPos[1]

        ; Clear previous drawings
        _WinAPI_RedrawWindow($hOverlay, 0, 0, $RDW_INVALIDATE + $RDW_UPDATENOW + $RDW_ERASE)

        ; Draw the rubber-band rectangle
        If Abs($currentX - $startX) > 1 Or Abs($currentY - $startY) > 1 Then
            DrawRectangle($hOverlayDC, $startX, $startY, $currentX, $currentY)
        EndIf

        ; Finalize selection on mouse release
        If Not _IsPressed("01") Then
            $endX = $currentX
            $endY = $currentY
            $isDragging = False
        EndIf
        Sleep(10)
    WEnd

    ; Cleanup
    _WinAPI_SetCursor(0) ; Reset cursor to default
    _GDIPlus_Shutdown()

	; Hide the overlay window
	GUISetState(@SW_MINIMIZE, $hOverlay)

    ; Show the main GUI again
    GUISetState(@SW_SHOW, $mainGUI)


    ; Calculate and return the rectangle area
    Local $rectX = Min($startX, $endX)
    Local $rectY = Min($startY, $endY)
    Local $rectWidth = Abs($endX - $startX)
    Local $rectHeight = Abs($endY - $startY)


	; Create smaller overlay for the selected rectangle area
	Local $hSelectedOverlay = GUICreate($selctAreaName, $rectWidth, $rectHeight, $rectX, $rectY, $WS_POPUP, $WS_EX_LAYERED + $WS_EX_TOPMOST + $WS_EX_TOOLWINDOW)
	GUISetBkColor(0x000000)
	WinSetTrans($selctAreaName, "", 100)
	GUISetState(@SW_SHOW, $hSelectedOverlay)

	; Draw the selected rectangle on the smaller overlay
	Local $hSelectedOverlayDC = _WinAPI_GetDC($hSelectedOverlay)

	; Assign the UI handle and DC handle from local to global
	If $selctAreaName = "PageArea" Then
		$hPageRectOverlay = $hSelectedOverlay
		$hPageRectOverlayDC = $hSelectedOverlayDC
	ElseIf $selctAreaName = "ButtonArea" Then
		$hButtonRectOverlay = $hSelectedOverlay
		$hButtonRectOverlayDC = $hSelectedOverlayDC
	EndIf

	DrawRectangle($hSelectedOverlayDC, 0, 0, $rectWidth, $rectHeight) ; Adjust coordinates to overlay's size


    ; Define and return array explicitly
    Local $rectArray[4]
    $rectArray[0] = $rectX
    $rectArray[1] = $rectY
    $rectArray[2] = $rectWidth
    $rectArray[3] = $rectHeight
    Return $rectArray
EndFunc


Func StartCapture()
    _GDIPlus_Startup()

    Local $pageImages[$pageCount]
    $numCroppedImage = 0
    $avgCroppedWidth = 0
    $avgCroppedHeight = 0
    For $i = 1 To $pageCount
        Local $imageName = "Page_" & StringFormat("%03d", $i) & ".png"
        CaptureAndSaveImage($pageRect[0], $pageRect[1], $pageRect[2], $pageRect[3], $imageName)
		Sleep(50)

		If $autoCrop = 1 Then
			Local $croppedImageName = CropImage($imageName)
			$pageImages[$i - 1] = $croppedImageName
		Else
			$pageImages[$i - 1] = $imageName
		EndIf


        ; Click the turn-page button
        MouseClick("left", $buttonRect[0] + ($buttonRect[2] / 2), $buttonRect[1] + ($buttonRect[3] / 2), 1, 10)
        Sleep(1000)
    Next

    ; Convert cropped images to pdf
    Local $pdfFileName = @YEAR & "-" & @MON & "-" & @MDAY & "_" & @HOUR & "_" & @MIN & "_" & @SEC & "_output.pdf"
    Local $pdfCreated = CombineImagesToPDF($pageImages, $pdfFileName)

    ; Delete the crooped images file
	If $pdfCreated = True Then
		For $i = 0 To UBound($pageImages) - 1
			FileDelete($pageImages[$i])
		Next
	EndIf


    _GDIPlus_Shutdown()

EndFunc

Func CombineImagesToPDF($aImages, $sOutputPDF)
    ; Build the ImageMagick command string
    Local $sCommand = "magick "

    ; Append each image path to the command string
    For $i = 0 To UBound($aImages) - 1
        $sCommand &= '"'  & @ScriptDir & "\" & $aImages[$i] & '" '
    Next

    ; Append the output PDF path to the command
    $sCommand &= '"' & @ScriptDir & "\" & $sOutputPDF & '"'

	ConsoleWrite($sCommand)

    ; Run the command to generate the PDF
    RunWait(@ComSpec & " /c " & $sCommand, "", @SW_HIDE )

    ; Check if the PDF was created successfully
    If FileExists($sOutputPDF) Then
        MsgBox(0, "Success", "PDF created successfully: " & $sOutputPDF)
		Return True
    Else
        MsgBox(16, "Error", "Failed to create PDF.")
		Return False
    EndIf
EndFunc

Func DrawRectangle($hDC, $x1, $y1, $x2, $y2)
    Local $left = Min($x1, $x2)
    Local $top = Min($y1, $y2)
    Local $right = Max($x1, $x2)
    Local $bottom = Max($y1, $y2)

    Local $hGraphics = _GDIPlus_GraphicsCreateFromHDC($hDC)
    Local $hPen = _GDIPlus_PenCreate(0x7FFF0000, 2)
    _GDIPlus_GraphicsDrawRect($hGraphics, $left, $top, $right - $left, $bottom - $top, $hPen)

    _GDIPlus_PenDispose($hPen)
    _GDIPlus_GraphicsDispose($hGraphics)
EndFunc

Func CaptureAndSaveImage($x, $y, $width, $height, $filename)
    $x *= $screenScale
    $y *= $screenScale
    $width *= $screenScale
    $height *= $screenScale

    Local $hBitmap = _WinAPI_CreateCompatibleBitmap(_WinAPI_GetDC(0), $width, $height)
    Local $hDCMem = _WinAPI_CreateCompatibleDC(_WinAPI_GetDC(0))
    Local $hOld = _WinAPI_SelectObject($hDCMem, $hBitmap)

    _WinAPI_BitBlt($hDCMem, 0, 0, $width, $height, _WinAPI_GetDC(0), $x, $y, $SRCCOPY)

    Local $hBitmapGDI = _GDIPlus_BitmapCreateFromHBITMAP($hBitmap)
    _GDIPlus_ImageSaveToFile($hBitmapGDI, $filename)

    _WinAPI_SelectObject($hDCMem, $hOld)
    _WinAPI_DeleteObject($hBitmap)
    _WinAPI_DeleteDC($hDCMem)
    _GDIPlus_ImageDispose($hBitmapGDI)
EndFunc

Func Min($val1, $val2)
    Return ($val1 < $val2) ? $val1 : $val2
EndFunc

Func Max($val1, $val2)
    Return ($val1 > $val2) ? $val1 : $val2
EndFunc

Func ConvertColorToGrey($color)
	Local $iR = BitShift(BitAND($color, 0x00FF0000), 16) ;extract red color channel
	Local $iG = BitShift(BitAND($color, 0x0000FF00), 8) ;extract green color channel
	Local $iB = BitAND($color, 0x000000FF) ;;extract blue color channel
	Local $iGrey = ($iR + $iG + $iB) / 3 ;convert pixels to average greyscale color format
	return $iGrey;
EndFunc

Func CropImage($filePath)
    _GDIPlus_Startup()
    Local $hImage = _GDIPlus_ImageLoadFromFile($filePath)
    Local $width = _GDIPlus_ImageGetWidth($hImage)
    Local $height = _GDIPlus_ImageGetHeight($hImage)
    Local $hBitmap = _GDIPlus_BitmapCloneArea($hImage, 0, 0, $width, $height, $GDIP_PXF32ARGB)

    ; Calculate the average color with larger step size to speed up
    Local $totalR = 0, $totalG = 0, $totalB = 0, $pixelCount = 0, $totalGrey = 0

	For $y = 0 To $height - 1 Step 1
		Local $color = _GDIPlus_BitmapGetPixel($hBitmap, $width / 2, $y)
		$totalGrey += ConvertColorToGrey($color)
		$pixelCount += 1
	Next

    Local $avgGrey = $totalGrey / $pixelCount
    Local $threshold = $avgGrey / 3  ; Adjust sensitivity
	ConsoleWrite("$avgGrey: " & $avgGrey & @CRLF);
	ConsoleWrite("$threshold: " & $threshold & @CRLF);


    ; Initialize crop boundaries
    Local $left = 0, $top = 0, $right = $width - 1, $bottom = $height - 1
    Local $color, $delta
	Local $step = 2
	Local $margin = 10

    ; Find boundaries with larger steps
    For $x = 0 To $width - 1 Step $step
        For $y = 0 To $height - 1 Step $step
            $color = _GDIPlus_BitmapGetPixel($hBitmap, $x, $y)
			Local $thisGrey = ConvertColorToGrey($color);
            If $thisGrey < $threshold Then
                $left = $x - $margin
                ExitLoop 2
            EndIf
        Next
    Next

    For $x = $width - 1 To 0 Step - $step
        For $y = 0 To $height - 1 Step $step
            $color = _GDIPlus_BitmapGetPixel($hBitmap, $x, $y)
			Local $thisGrey = ConvertColorToGrey($color);
            If $thisGrey < $threshold Then
                $right = $x + $margin
                ExitLoop 2
            EndIf
        Next
    Next

    For $y = 0 To $height - 1 Step $step
        For $x = 0 To $width - 1 Step $step
            $color = _GDIPlus_BitmapGetPixel($hBitmap, $x, $y)
			Local $thisGrey = ConvertColorToGrey($color);
            If $thisGrey < $threshold Then
                $top = $y - $margin
                ExitLoop 2
            EndIf
        Next
    Next

    For $y = $height - 1 To 0 Step - $step
        For $x = 0 To $width - 1 Step $step
            $color = _GDIPlus_BitmapGetPixel($hBitmap, $x, $y)
			Local $thisGrey = ConvertColorToGrey($color);
            If $thisGrey < $threshold Then
                $bottom = $y + $margin
                ExitLoop 2
            EndIf
        Next
    Next

    ; Calculate new cropped dimensions
    Local $cropWidth = $right - $left + 1
    Local $cropHeight = $bottom - $top + 1

	; Adjust the crop width and height if they are abnormally small
	; Usually this happens at the last page
	Local $portion = 0.5
    If ($cropWidth < $avgCroppedWidth * $portion) Then
        $cropWidth = $avgCroppedWidth
    EndIf

    If ($cropHeight < $avgCroppedHeight * $portion) Then
        $cropHeight = $avgCroppedHeight
    EndIf

    $avgCroppedWidth = ($avgCroppedWidth * $numCroppedImage + $cropWidth) / ($numCroppedImage + 1)
    $avgCroppedHeight = ($avgCroppedHeight * $numCroppedImage + $cropHeight) / ($numCroppedImage + 1)
    $numCroppedImage += 1

    ; Create the cropped bitmap
    Local $hCroppedBitmap = _GDIPlus_BitmapCloneArea($hBitmap, $left, $top, $cropWidth, $cropHeight, $GDIP_PXF32ARGB)

    ; Save cropped image
    Local $newFilePath = StringReplace($filePath, ".png", "_cropped.png")
    _GDIPlus_ImageSaveToFile($hCroppedBitmap, $newFilePath)
    ConsoleWrite("Cropped image saved to: " & $newFilePath & @CRLF)

    ; Cleanup
    _GDIPlus_BitmapDispose($hCroppedBitmap)
    _GDIPlus_BitmapDispose($hBitmap)
    _GDIPlus_ImageDispose($hImage)
    _GDIPlus_Shutdown()

    ; Delete original image
    FileDelete($filePath)
    ConsoleWrite("Original image deleted: " & $filePath & @CRLF)

	Return $newFilePath
EndFunc


; Start the program
Main()
