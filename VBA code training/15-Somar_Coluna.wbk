' Author			Date
' danbros			
'
'https://github.com/danbros/Excel-worksheet
'
'AULA 15 - Soma valor total da coluna A de A1 até “”, mostra resultado em B1
'
'Sub em botão de comando (ActiveX)

Private Sub btSum_Click()

	Dim result  As Integer
	Dim WS      As Worksheet


	Set WS = Sheets("Plan 1")
	result = 0 'para quando rodar ter a certeza de não ter valor ou ter trazido lixo da memória

	W.Select

	W.Range("A1").Select

	Do While ActiveCell <> ""
	    
	    result = result + ActiveCell.Value
	    ActiveCell.Offset(1, 0).Select

	Loop

	W.Range("B1").Value = result

End Sub

