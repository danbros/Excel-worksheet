' Author			Date
' danbros			
'
'https://github.com/danbros/Excel-worksheet
'
'AULA 21-22 Bloquear e desbloquear plan com e sem senha dinâmica
'Sub em botão de comando (ActiveX)

'(22) = senha dinâmica

'Botão bloquear
Private Sub btBlock_Click()

	Dim Count       	As Integer
	Dim CurrentPlan   	As Variant
	'Dim Psword      	As String (22)

	'imput usuario] (22)
	'Psword = inputBox("Digite a senha de bloqueio", "Atenção") [imput usuario]

	'Pausa atualização de tela
	Application.ScreenUpdating = False

	Count = Sheets.count 'quantidade de planilhas

	For Each CurrentPlan In Sheets
	    CurrentPlan.Protect Password:="123"    'ou Password:=Psword (22)
	Next

	'Normaliza
	Application.ScreenUpdating = True

	MsgBox "Block " & Count & " plan...", vbOKOnly, "Status"

End Sub

'Botão desbloquear'
Private Sub btUnlock_Click()

	Dim Count			As Integer
	Dim CurrentPlan   	As Variant
	'Dim Psword      	As String

	'[imput usuario] (22)
	'sPsword = inputBox("Digite a senha de desbloqueio", "Atenção")

	Application.ScreenUpdating = False

	Count = Sheets.count 'quantidade de planilhas

	For Each CurrentPlan In Sheets
	    CurrentPlan.Unprotect Password:="123"  'ou Password:=Psword
	Next

	Application.ScreenUpdating = True

	MsgBox "Unlock " & Count & " plan...", vbOKOnly, "Status"

End Sub