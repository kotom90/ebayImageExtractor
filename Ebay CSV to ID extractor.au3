#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <FileConstants.au3>

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
		
		ConsoleWrite($aColumnValues[$i] & @CRLF)
    Next
	;ConsoleWrite(UBound($aColumnValues))
	
EndIf
