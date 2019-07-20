'@danbros
'Resumo código VBA Marketplace

'======WORKBOOK OPEN()======

Private explicit

Private Sub Workbook_Open()

Application.ScreenUpdating = False

' Desativa todas as tabs das planilhas
ActiveWindow.DisplayWorkbookTabs = False

' Chama função para determinar zoom
Z = func_z()

For Each WS In Worksheets
    bloquear
    ActiveWindow.zoom = Z
Next

Application.ScreenUpdating = True

End Sub

'======================================

'===========Módulo1===========

Private explicit
Option Private Module

' Constante senha
Public Const KEY As String = "123"

Sub Inserir_Mercadoria()
    
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim rng_Table As Range
    
    Set WB = ActiveWorkbook
    Set WS = WB.Sheets("Compra de Mercadorias")
    
    'Desliga movimento de tela
    Application.ScreenUpdating = False
    'Desprotege plan
    WS.Unprotect Password:=KEY
    
    ''Elimina um erro não identificado de Tabela bloqueada
    WS.ListObjects(1).Range.Offset(1, 0).Value = _
    WS.ListObjects(1).Range.Offset(1, 0).Value
    
    WS.ListObjects(1).ListRows.Add (1)
    WS.ListObjects(1).Range.Activate
    
    'Aponta para primeira célula da Tabela
    Set rng_Table = ActiveCell.Offset(1, 0)
    
    'Recebe string com fórmula
    rng_Table.Value = "=TODAY()"
    rng_Table.Value = rng_Table.Value
    rng_Table.Offset(0, 1).Select
    
    'func
    bloquear
    Application.ScreenUpdating = True
    
End Sub

Sub Excluir()
    
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim rng_Table As Range
    Dim obj_Table As ListObject
    
    Set WB = ActiveWorkbook
    Set WS = WB.Sheets(ActiveSheet.Name)
    Set rng_Table = Selection
    'aponta para a primeira tabela da planilha
    Set obj_Table = WS.ListObjects(1)
    
    Application.ScreenUpdating = False
    WS.Unprotect Password:=KEY
       
    'Se rngTable estiver fora da tabela, invoca msg.
    If Intersect(rng_Table, WS.ListObjects(obj_Table.Name).DataBodyRange) Is Nothing Then
        MsgBox "Não foi selecionado uma célula dentro da tabela"
    ElseIf rng_Table.Cells.Count > 1 Then
        MsgBox "Selecione apenas uma célula de cada vez :("
    Else
        rng_Table.Rows.Delete
        MsgBox "Registro excluído com sucesso", vbOKOnly, "Status"
    End If
    
    'func
    bloquear
    Application.ScreenUpdating = True
    
End Sub

Sub Nova_Venda()
    
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim rng_Table As Range

    Set WB = ActiveWorkbook
    Set WS = WB.Sheets(ActiveSheet.Name)
    
    Application.ScreenUpdating = False
    WS.Unprotect Password:=KEY
    
    'Elimina um erro não identificado de Tabela bloqueada
    WS.ListObjects(1).Range.Offset(1, 0).Value = _
    WS.ListObjects(1).Range.Offset(1, 0).Value
    
    WS.ListObjects(1).ListRows.Add (1)
    WS.ListObjects(1).Range.Select

    Set rng_Table = ActiveCell.Offset(1, 0)
    
    rng_Table.Value = rng_Table.Offset(1, 0).Value + 1
    rng_Table.Offset(0, 1).Value = "=TODAY()"
    rng_Table.Offset(0, 1).Value = rng_Table.Offset(0, 1).Value
    rng_Table.Offset(0, 2).Select
    
    'func
    bloquear
    Application.ScreenUpdating = True
    
End Sub
Sub Inserir_Produto()
    
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim rng_Table As Range

    Set WB = ActiveWorkbook
    Set WS = WB.Sheets(ActiveSheet.Name)
    
    Application.ScreenUpdating = False
    WS.Unprotect Password:=KEY
    
    'Elimina um erro estranho de Tabela bloqueada
    WS.Range("B5").Value = WS.Range("B5").Value
    
    WS.ListObjects(1).ListRows.Add (1)
    WS.ListObjects(1).Range.Select

    Set rng_Table = ActiveCell.Offset(1, 0)
    
    rng_Table.Value = rng_Table.Offset(1, 0).Value
    rng_Table.Offset(0, 1).Value = "=TODAY()"
    rng_Table.Offset(0, 1).Value = rng_Table.Offset(0, 1).Value
    rng_Table.Offset(0, 2).Select
    
    'func
    bloquear
    Application.ScreenUpdating = True
    
End Sub

'=========================================

'======Módulo2======

Private explicit
Option Private Module

' Constante senha
Public Const KEY As String = "123"

Sub Inserir_Mercadoria()
    
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim rng_Table As Range
    
    Set WB = ActiveWorkbook
    Set WS = WB.Sheets("Compra de Mercadorias")
    
    'Desliga movimento de tela
    Application.ScreenUpdating = False
    'Desprotege plan
    WS.Unprotect Password:=KEY
    
    ''Elimina um erro não identificado de Tabela bloqueada
    WS.ListObjects(1).Range.Offset(1, 0).Value = _
    WS.ListObjects(1).Range.Offset(1, 0).Value
    
    WS.ListObjects(1).ListRows.Add (1)
    WS.ListObjects(1).Range.Activate
    
    'Aponta para primeira célula da Tabela
    Set rng_Table = ActiveCell.Offset(1, 0)
    
    'Recebe string com fórmula
    rng_Table.Value = "=TODAY()"
    rng_Table.Value = rng_Table.Value
    rng_Table.Offset(0, 1).Select
    
    'func
    bloquear
    Application.ScreenUpdating = True
    
End Sub

Sub Excluir()
    
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim rng_Table As Range
    Dim obj_Table As ListObject
    
    Set WB = ActiveWorkbook
    Set WS = WB.Sheets(ActiveSheet.Name)
    Set rng_Table = Selection
    'aponta para a primeira tabela da planilha
    Set obj_Table = WS.ListObjects(1)
    
    Application.ScreenUpdating = False
    WS.Unprotect Password:=KEY
       
    'Se rngTable estiver fora da tabela, invoca msg.
    If Intersect(rng_Table, WS.ListObjects(obj_Table.Name).DataBodyRange) Is Nothing Then
        MsgBox "Não foi selecionado uma célula dentro da tabela"
    ElseIf rng_Table.Cells.Count > 1 Then
        MsgBox "Selecione apenas uma célula de cada vez :("
    Else
        rng_Table.Rows.Delete
        MsgBox "Registro excluído com sucesso", vbOKOnly, "Status"
    End If
    
    'func
    bloquear
    Application.ScreenUpdating = True
    
End Sub

Sub Nova_Venda()
    
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim rng_Table As Range

    Set WB = ActiveWorkbook
    Set WS = WB.Sheets(ActiveSheet.Name)
    
    Application.ScreenUpdating = False
    WS.Unprotect Password:=KEY
    
    'Elimina um erro não identificado de Tabela bloqueada
    WS.ListObjects(1).Range.Offset(1, 0).Value = _
    WS.ListObjects(1).Range.Offset(1, 0).Value
    
    WS.ListObjects(1).ListRows.Add (1)
    WS.ListObjects(1).Range.Select

    Set rng_Table = ActiveCell.Offset(1, 0)
    
    rng_Table.Value = rng_Table.Offset(1, 0).Value + 1
    rng_Table.Offset(0, 1).Value = "=TODAY()"
    rng_Table.Offset(0, 1).Value = rng_Table.Offset(0, 1).Value
    rng_Table.Offset(0, 2).Select
    
    'func
    bloquear
    Application.ScreenUpdating = True
    
End Sub
Sub Inserir_Produto()
    
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim rng_Table As Range

    Set WB = ActiveWorkbook
    Set WS = WB.Sheets(ActiveSheet.Name)
    
    Application.ScreenUpdating = False
    WS.Unprotect Password:=KEY
    
    'Elimina um erro estranho de Tabela bloqueada
    WS.Range("B5").Value = WS.Range("B5").Value
    
    WS.ListObjects(1).ListRows.Add (1)
    WS.ListObjects(1).Range.Select

    Set rng_Table = ActiveCell.Offset(1, 0)
    
    rng_Table.Value = rng_Table.Offset(1, 0).Value
    rng_Table.Offset(0, 1).Value = "=TODAY()"
    rng_Table.Offset(0, 1).Value = rng_Table.Offset(0, 1).Value
    rng_Table.Offset(0, 2).Select
    
    'func
    bloquear
    Application.ScreenUpdating = True
    
End Sub

