' Author			Date
' danbros			
'
'https://github.com/danbros/Excel-worksheet
'
'AULA 24 Imprime nome de usuario e data de entrada e saida em 'Controle de Acesso"
'Sub na Entrada e saída do Workbook

Private Sub Workbook_Open()

    'Declaração de variável
    Dim WS          As Worksheet
    Dim LastCell    As Range
    
    Set WS = Sheets("Controle de Acesso")
    'var lastCell recebe célula que for ctrl+up  +  down a partir de A1048576
    Set LastCell = WS.Range("A1048576").End(xlUp).Offset(1, 0)
    
    'recebe nome do usuario
    LastCell.Value = Environ$("Username")
    'recebe data e hora na celula a direita dele
    LastCell.Offset(0, 1).Value = Date & " / " & Time
    
    Sheets("Plan 1").Select

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)

    Dim WS          As Worksheet
    Dim LastCell    As Range
    
    Set WS = Sheets("Controle de Acesso")
    'var lastCell recebe célula que for ctrl+up  +  down  + 2x.right
    Set LastCell = WS.Range("A1048576").End(xlUp).Offset(0, 2)
    
    'imprime data e hora da saída na celula lastCell
    LastCell.Value = Date & " / " & Time
    
    Sheets("Plan 1").Select
    
    'Desliga alertas (como o de "quer salvar?" ao sair)
    Application.DisplayAlerts = False
        'Salva a pasta de trabalho ativa
        ActiveWorkbook.Save
    'Liga alertas de display
    Application.DisplayAlerts = True

End Sub