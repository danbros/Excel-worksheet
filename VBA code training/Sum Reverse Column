' Author			Date
' danbros			
'
'https://github.com/danbros/Excel-worksheet
'
'AULA 16 Soma desde a última linha da coluna A1 preenchida até a linha 1 e mostra o resultado em B1
'Sub em botão de comando (ActiveX)

Private Sub btRSum_Click()

	Dim result  As Integer
	Dim WS      As Worksheet
	'Dim ultCel	As Range	

	Set WS = Sheets("Plan 1")
	result = 0

	'set ultCel = WS.Range("A1048576").End(xlUp)
	WS.Range("A1048576").Select
	ActiveCell.End(xlUp).Select

			'ultCel.Row > 1
	Do While ActiveCell.Row > 1
	    
	    result = result + ActiveCell.Value 
	    ActiveCell.Offset(-1, 0).Select

	Loop

	WS.Range("B1").Value = result
	'Alternativa para mostrar no fim da coluna(se ultCel range não foi modificado:
	'ultCel.Offset(1, 0).Value = "O Resultado é " & Resultado

End Sub