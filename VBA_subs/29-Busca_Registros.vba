' Author            Date
' danbros
'
'https://github.com/danbros/Excel-worksheet
'
'AULA 29 Captura registros da coluna "A" da Plan1 e busca nas culunas "A"
' de todas as Plan do WorkBook, depois imprime em qual Plan foi
' encontrado outro registro identico


Option Explicit

Sub Busca()

Dim WS               As Worksheet
Dim UltCel           As Range
Dim Nome             As String
Dim NomePlan         As String

Set WS = Sheets("Planilha1")

Application.ScreenUpdating = False

WS.Select
Range("B:B").ClearContents

Range("A1").Select

'Loop 1, percorre Plan1 para capturar nomes procurados
Do While ActiveCell <> ""
    
    'Guarda nome procurado
    Nome = ActiveCell.Value
    
    'Guarda o range da ultima célula da Plan1 usada na pesquisa
    Set UltCel = ActiveCell
    
    'Loop 2, percorre todas as Plan. Começa em "A" (2, segunda Plan do WB)
    'até última Plan (Sheets.Count)
    
    Dim i As Integer
    
    For i = 2 To Sheets.Count
    
        'Seleciona Plan de acordo com iterador "i"
        Sheets(i).Select
        
        Sheets(i).Range("A1").Select
        
        'Guarda o nome da última Plan percorrida neste loop
        NomePlan = Sheets(i).Name
    
        'Percorre coluna "A" da Plan(i)
        Do While ActiveCell.Value <> ""
            
            'Se célula ativa for igual nome da Plan 1, então
            If ActiveCell.Value = Nome Then
            
                WS.Select
                UltCel.Select
                
                'Se célula à frente da ativa estiver vazia, então
                If ActiveCell.Offset(0, 1).Value = "" Then
                
                    ActiveCell.Offset(0, 1).Value = NomePlan
                    Sheets(i).Select
                    Exit Do
                
                Else
                
                    ActiveCell.Offset(0, 1).Value = ActiveCell.Offset(0, 1).Value _
                    & " / " & NomePlan
                
                    Exit Do
                
                End If
        
            End If
        
            ActiveCell.Offset(1, 0).Select
            
        Loop
    
    Next
    
    WS.Select
    UltCel.Select

    ActiveCell.Offset(1, 0).Select

Loop

Application.ScreenUpdating = True

End Sub