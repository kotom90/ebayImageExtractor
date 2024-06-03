#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <FileConstants.au3>
#include <WindowsConstants.au3>
#include <GUIConstantsEx.au3>

$mainWindow = GUICreate("Ebay Auto Tools",400,200)
$btnGetAllImages = GUICtrlCreateButton("Get all images from csv", 10,10,150,25)
$btnTypeImgUrls = GUICtrlCreateButton("Type image URLs to XLS", 170,10,150,25)

GUISetState(@SW_SHOW,$mainWindow)

While 1
	Switch GUIGetMsg()
        Case $GUI_EVENT_CLOSE
            Exit
		Case $btnGetAllImages
			downloadAllImages()
		Case $btnTypeImgUrls
			sleep(5000)
			typeImgUrls()
	EndSwitch
WEnd

Func typeImgUrls()
	$itemIds = getItemIdsFromCsv()
	for $i = 1 To UBound($itemIds) - 1
		$imageUrls = getImageUrlsForItemId($itemIds[$i])
		for $j = 1 To UBound($imageUrls) - 1
			;ClipPut($imageUrls[$j])
			;Send("^v")
			Send($imageUrls[$j])
			sleep(100)
			if ($j < UBound($imageUrls) - 1) Then
				Send("|")
			EndIf
		next
		Send("{ENTER}")
	next
EndFunc

Func downloadAllImages()
	$itemIds = getItemIdsFromCsv()
	Local $folderPath = @ScriptDir & "\images"
	if not FileExists($folderPath) Then
		DirCreate($folderPath)
	EndIf
	for $i = 1 To UBound($itemIds) - 1
		if not FileExists($folderPath & "\" & $itemIds[$i]) Then
			DirCreate($folderPath & "\" & $itemIds[$i])
		EndIf
		$folderPathImages = $folderPath & "\" & $itemIds[$i] 
		$imageUrls = getImageUrlsForItemId($itemIds[$i])
		ConsoleWrite($i & @CRLF)
		for $j = 1 To UBound($imageUrls) - 1
			ConsoleWrite($j & " " & $imageUrls[$j] & @CRLF)
			$imageUrl = $imageUrls[$j]
			Local $pattern = "\.png"
			if StringRegExp($imageUrl, $pattern) Then
				$fileName = $folderPathImages & "\" & $j & ".png"
			Else
				$fileName = $folderPathImages & "\" & $j & ".jpg"
			EndIf
			InetGet($imageUrl, $fileName)
				;MsgBox(0,"","")
				;$file = FileOpen($folderPathImages & "\" & $j,2)
				;FileWrite($file, $image)
		next
	next
EndFunc


Func getItemIdsFromCsv()
	; Replace with the path to your CSV file
	Local $sCSVFilePath = @ScriptDir & "\ebay.com_unsold_10.03.2024.csv"

	; Specify the column number to extract (1-based index)
	Local $iColumnToExtract = 0

	; Read the content of the CSV file
	Local $sCSVContent = FileRead($sCSVFilePath)

	$string = BinaryToString($sCSVContent)
	$string = StringReplace($string,'"','')
	;ConsoleWrite($string)

	; Check if reading was successful
	If @error Then
		MsgBox($MB_OK + $MB_ICONERROR, "Error", "Failed to read CSV file. Error code: " & @error)
	Else
		;ConsoleWrite($sCSVContent)
		; Split the CSV content into rows
		Local $aCSVRows = StringSplit($string, @LF, BitOr($STR_NOCOUNT,$STR_CHRSPLIT))
		;ConsoleWrite(UBound($aCSVRows))

		; Initialize an array to store the values of the specified column
		Local $aColumnValues[UBound($aCSVRows)]

		;ConsoleWrite(UBound($aCSVRows))
		
		; Loop through each row and extract the specified column
		For $i = 1 To UBound($aCSVRows) - 1
			; Split the row into columns using comma as the delimiter
			Local $aColumns = StringSplit($aCSVRows[$i], ",",$STR_NOCOUNT)
			$aColumnValues[$i] = $aColumns[$iColumnToExtract]
			
			;ConsoleWrite($aColumnValues[$i] & @CRLF)
		Next
		;ConsoleWrite(UBound($aColumnValues))
		
	EndIf
	return $aColumnValues
EndFunc

Func getImageUrlsForItemId($itemId)
	sleep(500)

	Local $sURL = "https://www.ebay.com/itm/" & $itemId
	;MsgBox(0,"Item ID",$itemId)

	; Send HTTP request and retrieve HTML content
	Local $sHTMLContent = InetRead($sURL)
	;Convert binary to string to be able to filter with regexp
	$contentString = BinaryToString($sHTMLContent)

	; Check if the download was successful
	If @error Then
		MsgBox($MB_OK + $MB_ICONERROR, "Error", "Failed to retrieve HTML content. Error code: " & @error)
	Else
		Local $allMatches[] = []
		; Extract URLs starting with "https" and save them to links.txt
		Local $pattern = "https:\/\/[^\s]+l1600\.png"
		$pngMatches = StringRegExp($contentString, $pattern, 3)

		$pngMatches = _ArrayUnique($pngMatches,0,0,0,$ARRAYUNIQUE_NOCOUNT)
		ConsoleWrite("PNG MATCHES: " & UBound($pngMatches)& @CRLF) 
		
		Local $pattern = "https:\/\/[^\s]+l1600\.jpg"
		$jpgMatches = StringRegExp($contentString, $pattern, 3)
		
		$jpgMatches = _ArrayUnique($jpgMatches,0,0,0,$ARRAYUNIQUE_NOCOUNT)
		ConsoleWrite("JPG MATCHES: " & UBound($jpgMatches) & @CRLF)
		
		for $i = 0 To UBound($pngMatches) - 1
			_ArrayAdd($allMatches,$pngMatches[$i])
		next
		
		for $i = 0 To UBound($jpgMatches) - 1
			_ArrayAdd($allMatches,$jpgMatches[$i])
		next
		
		;allMatches has the first element empty
		ConsoleWrite("ALL MATCHES: " & UBound($allMatches) - 1 & @CRLF)
		return $allMatches
	EndIf
EndFunc