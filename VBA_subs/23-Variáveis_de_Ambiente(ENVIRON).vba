' Author			Date
' danbros			
'
'https://github.com/danbros/Excel-worksheet
'
'AULA 23 Imprime todas as variáveis de ambiente
'Sub em bot?o de comando (ActiveX)

Private Sub btENV_Click ()
	
	Dim WS 		As Worksheet
	Dim Count	As Integer

	Set WS = Sheets("Plan 1")

	For Count = 1 To 100

		WS.Cells(Count, 1).Value = Environ$(Count)

	Next

End Sub


