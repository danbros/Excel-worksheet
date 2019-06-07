' Author            Date
' danbros
'
'https://github.com/danbros/Excel-worksheet
'
'AULA 30 Criando combinações de valores com base na coluna A
'(Base gerador de número de loteria)


Option Explicit

Sub Combinação_Vertical()

Dim WS          As Worksheet    'Var sheets 1
Dim Valor       As Range        'Var para valor a ser combinado
Dim UltLin      As Range        'Var última celula vazia na coluna C
Dim i           As Integer      'Var contadora

Set WS = Sheets("Planilha1")

WS.Select
WS.Range("A1").Select

Set Valor = ActiveCell

'Reseta a coluna C para imprimir novas combinações
WS.Range("C:C").EntireColumn.ClearContents
WS.Range("C1").Value = "Comb. Vertical"

'Loop 1 para percorrer valores da coluna A
Do While ActiveCell.Value <> ""
    
    'Loop 2 para percorer coluna A combinando "Valor" com ´ActiveCel
    '"Application.WorksheetFunction.CountA()" é igual função CONT.VALORES()
    For i = 1 To Application.WorksheetFunction.CountA(WS.Range("A:A"))
    
        'Seleciona célula ativa a ser combinada com "Valor"
        WS.Cells(i, 1).Select
       
        'Guarda a range da linha vazia da coluna C a ser preenchida com combinação
        'Set UltLin = WS.Range("C1048576").End(xlUp).Offset(1, 0)
        Set UltLin = WS.Range("C" & Rows.Count).End(xlUp).Offset(1, 0)
                    
        'Condição para não comparar valores iguais
        If ActiveCell.Value <> Valor.Value Then
            
            
'            -----------Combinação excluindo qualquer repetição (1 - 2, 2 - 1)------------
'
'            Dim j As Integer
'            Dim Str As String
'
'            Str = ActiveCell.Value & " - " & Valor.Value
'
'            For j = Application.WorksheetFunction.CountA(WS.Range("C:C")) To 1 Step -1
'
'                If Str = Cells(j, 3).Value Then
'
'                     j = -1
'                     Exit For
'
'                End If
'
'            Next
'
'            If j <> -1 Then
'
'                WS.Range("C" & UltLin.Row).Value = "'" & Valor.Value & " - " & ActiveCell.Value
'
'            End If
            
            
            '--------Combinação excluido apenas o mesmo número (1 - 1)
            
            'Concatena "Valor" com célula ativa e imprime na última coluna C vazia
            '*Apóstrofo para range receber valor em string e não identificar "/" como
            'símbolo matemático.
            WS.Range("C" & UltLin.Row).Value = "'" & Valor.Value & " - " & ActiveCell.Value
            
        End If
    
    Next
    
    Set Valor = Valor.Offset(1, 0)
    Valor.Select

Loop

End Sub