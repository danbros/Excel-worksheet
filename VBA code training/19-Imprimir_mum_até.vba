' Author			Date
' danbros			
'
'https://github.com/danbros/Excel-worksheet
'
'AULA 19 Imprimir numeros até onde o usuário desejar (b2)
'Sub em botão de comando (ActiveX)

Private Sub btFor_Click()

	Dim WS      As Worksheet
	Dim count   As Integer

	Set WS = Sheets("Planilha1")

	Range("A:A").ClearContents

	For count = 1 To WS.Range("B2").Value
	    Range("A" & count).Value = count
	Next
	
End Sub
