' Author			Date
' danbros			
'
'https://github.com/danbros/Excel-worksheet
'
'AULA 18 Somar numeros da coluna positivos e negativos separadamente
'Sub em botão de comando (ActiveX)

Private Sub btPosNeg_Click()

	Dim WS           As Worksheet
	Dim resultPos   As Long
	Dim resultNeg   As Long

	Set WS = Sheets("Plan 1")
	resultPos = 0
	resultNeg = 0

	WS.Range("A1").Select

	Do While ActiveCell.Value <> ""
	    
	    If ActiveCell.Value > 0 Then
	        resultPos = resultPos + ActiveCell
	    Else
	        resultNeg = resultNeg + ActiveCell
	    End If
	    
	 ActiveCell.Offset(1, 0).Select
	 
	 Loop
	 
	 MsgBox "Total Pos: " & resultPos & Chr(13) & "Total Neg: " & resultNeg
	 
	 Range("B1").Value = "Total Pos: " & resultPos
	 Range("B2").Value = "TOtal Neg: " & resultNeg
 
End Sub
