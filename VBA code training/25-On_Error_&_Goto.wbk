' Author            Date
' danbros
'
'https://github.com/danbros/Excel-worksheet
'
'AULA 25 Tratamento de erro com "On Error Resume Next", "On Error Goto", e ponto de sa?da.
'Subs em bot?o de comando (ActiveX)

Private Sub btBlock_Click()

    Dim Count           As Integer
    Dim CurrentPlan     As Variant
    Dim Psword          As String

    Psword = InputBox("Digite a senha de bloqueio", "Aten??o")

    Application.ScreenUpdating = False

    Count = Sheets.Count

    For Each CurrentPlan In Sheets
    
        CurrentPlan.Protect Password:=Psword
        
    Next

    Application.ScreenUpdating = True

    MsgBox "Block " & Count & " plan...", vbOKOnly, "Status"

End Sub

Private Sub btUnlock_Click()

    Dim Count           As Integer
    Dim CurrentPlan     As Variant
    Dim Psword          As String

    Psword = InputBox("Digite a senha de desbloqueio", "Aten??o")

    Application.ScreenUpdating = False

    Count = Sheets.Count

    'As pr?ximas 2 linhas s?o ignor?veis
    'Ao errar retoma pr?xima instru??es(Next)
    On Error Resume Next
    'deve-se lembrar de desativar depois com
    On Error GoTo 0
    '(que desativa o On Error Resume Next) ou continuar?o pulando erros
    
    For Each CurrentPlan In Sheets
    
        'Ao emitir erro, vai para "error_cod:", executa e retorna na prox instru??o (pulando a instru??o com erro)
        On Error GoTo error_cod
        
        CurrentPlan.Unprotect Password:=Psword
        
    Next
    
    Application.ScreenUpdating = True

    MsgBox "Unlock " & Count & " plan...", vbOKOnly, "Status"
    '"Exit sub" abaixo sai da sub no caso da senha ser a correta
    'Exit Sub     'Posse ser trocada por um ponto de sa?da, que trata todas as sa?das do algoritmo, executando logo abaixo
    
exit_point:
    'Para o caso de um erro acabar travando o ponto de sa?da
    On Error Resume Next
    
    'se houvesse algum arquivo aberto ou vari?vel, aqui poderia fechar ou eliminar ela da mem?ria.
    'ex: set WS = nothing
    
    'Ponto de sa?da da aplicação, quando houver erro (chamada em error_cod) e quando n?o houver.
    Exit Sub

error_cod:
    MsgBox "Plan not unlock. Invalid password", vbOKOnly, "Status"
    
    'Chama ponto de sa?da, para n?o executar prox instru??o, j? que a senha estar? errada
    Resume exit_point
    
End Sub
