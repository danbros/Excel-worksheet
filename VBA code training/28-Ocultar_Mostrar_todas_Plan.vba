' Author            Date
' danbros
'
'https://github.com/danbros/Excel-worksheet
'
'AULA 28 Apaga conteúdo de todas as planilhas na "Sheets" ativa
'Oculta/Mostra todas as planilhas


Option Explicit

Sub btApaga_Esconde_Reexibe()

'Para cada planilha (varPlan) em "Sheets", faça...
For Each varPlan In Sheets

    'Se nome da plan for diferente de "Nome Plan qualquer", faça
    If varPlan.Name <> "Nome Plan qualquer" Then
    
        'Apaga conteúdo das células usadas em varPlan
        varPlan.UsedRange.ClearContents
        
        'Exclui coluna inteira de células usadas
        varPlan.UsedRange.EntireColumn.Delete
        
        'Ocultas planilhas
        varPlan.Visible = False
        
        'Mostra planilhas
        varPlan.Visible = True
    
    End If
       
Next

MsgBox "OK", vbOKOnly

End Sub