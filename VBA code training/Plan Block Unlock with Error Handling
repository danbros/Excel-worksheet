' Author			Date
' danbros			
'
'https://github.com/danbros/Excel-worksheet
'
'AULA 21-22 Bloquear e desbloquear plan com e sem senha dinâmica
'Sub em botão de comando (ActiveX)

Private Sub btBlock_Click()

	Dim Count       	As Integer
	Dim CurrentPlan   	As Variant
	Dim Psword      	As String

	Psword = inputBox("Digite a senha de bloqueio", "Atenção")

	Application.ScreenUpdating = False

	Count = Sheets.count

	For Each CurrentPlan In Sheets
	    CurrentPlan.Protect Password:=Psword
	Next

	Application.ScreenUpdating = True

	MsgBox "Block " & Count & " plan...", vbOKOnly, "Status"

End Sub

Private Sub btUnlock_Click()

	Dim Count			As Integer
	Dim CurrentPlan   	As Variant
	Dim Psword      	As String

	Psword = inputBox("Digite a senha de desbloqueio", "Atenção")

	Application.ScreenUpdating = False

	Count = Sheets.count

	For Each CurrentPlan In Sheets
	    CurrentPlan.Unprotect Password:=Psword
	Next

	Application.ScreenUpdating = True

	MsgBox "Unlock " & Count & " plan...", vbOKOnly, "Status"

End Sub