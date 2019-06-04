' Author            Date
' danbros
'
'https://github.com/danbros/Excel-worksheet
'
'AULA 26 Cadastrar e Excluir registros usando ActiveX e validação de dados
'AULA 27 Alterar registros


'***WORKBOOK (VBA)***

Option Explicit

Private Sub Workbook_Open()
'
' /Workbook/Open do 26-Registros_Cad.xlsm
' Sub para atualizar "ComboBox1" e "CBoxBusca"
' da Plan "MODO 1" ao abrir WB
'

AttComboBox1

End Sub




'***PLAN 1 (VBA)***

Option Explicit

Private Sub btCadastro1_Click()
'
' /Plan1/btCadastro1/Click do 27-Registros_Cad.xlsm
' Sub para o botão "Cadastro" (em "Clientes")da Plan "MODO 1"
'

Dim WSCad     As Worksheet
Dim WS         As Worksheet
Set WSCad = Sheets("MODO 1")
Set WS = Sheets("Clientes")

Application.ScreenUpdating = False

WSCad.Select
    
Range("B5:G5").Select
Selection.Copy

WS.Activate
WS.Range("A1048576").Select
Selection.End(xlUp).Select
ActiveCell.Offset(1, 0).Select
'Range("A3").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False   'Colar somente valores
'ActiveSheet.Paste

WSCad.Select
Application.CutCopyMode = False
Selection.ClearContents

MsgBox "Cadastrado...", vbOKOnly, "Status"

Call AttComboBox1 '& updating = true

End Sub



Private Sub btBusca_Click()
'
' /Plan1/CBoxBusca/Click do 27-Registros_Cad.xlsm
' Sub para buscar cliente pelo "CBoxBusca" da Plan "MODO 1"
'

Dim WS As Worksheet
Dim WSCad As Worksheet
Dim Nome As String

Set WS = Sheets("Clientes")
Set WSCad = Sheets("MODO 1")
Nome = Sheets("MODO 1").CBoxBusca.Value

Application.ScreenUpdating = False

'Checa se CBoxBusca está vazio, se sim encerra
If Nome = "" Then
    
    MsgBox "Nenhum nome selecionado, busca abortada", vbOKOnly, "Status"
    Exit Sub

End If

WS.Select
WS.Range("A2").Select

'Percorre lista clientes enquanto <> ""
Do While ActiveCell.Value <> ""

    If Nome = ActiveCell.Value Then
       
       'WSCad.Range("B5").Value = ActiveCell.Value
       'WSCad.Range("C5").Value = ActiveCell.Offset(0, 1).Value
       'WSCad.Range("D5").Value = ActiveCell.Offset(0, 2).Value
       'WSCad.Range("E5").Value = ActiveCell.Offset(0, 3).Value
       'WSCad.Range("F5").Value = ActiveCell.Offset(0, 4).Value
       'WSCad.Range("G5").Value = ActiveCell.Offset(0, 5).Value
       
       WSCad.Range("B5:G5").Value = WS.Range(ActiveCell.Address, ActiveCell.Offset(0, 5).Address).Value
       
       Exit Do
    
    End If
    
    ActiveCell.Offset(1, 0).Select
    
Loop

WSCad.Select
WSCad.Range("B5").Select

Application.ScreenUpdating = True

End Sub



Private Sub btExcluir1_Click()
'
' /Plan1/btExcluir1/Click do 26-Registros_Cad.xlsm
' Sub para excluir cliente pelo "ComboBox1" da Plan "MODO 1"
'

Dim WS As Worksheet
Dim Nome As String

Set WS = Sheets("Clientes")
Nome = Sheets("MODO 1").ComboBox1.Value

Application.ScreenUpdating = False

'Checa se ComboBox1 está vazio, se sim encerra
If Nome = "" Then
    
    MsgBox "Nenhum nome selecionado, exclusão abortada", vbOKOnly, "Status"
    Exit Sub

End If

WS.Select
WS.Range("A2").Select

'Percorre lista clientes enquanto <> ""
Do While ActiveCell.Value <> ""

    If Nome = ActiveCell.Value Then
        
        'Deleta linha da Tabela
        ActiveCell.EntireRow.Delete
        MsgBox "Cadastro Apagado", vbOKOnly, "Status"
        Exit Do
    
    End If
    
    ActiveCell.Offset(1, 0).Select
    
Loop

'Invoca sub para atualizar cbbox
AttComboBox1

Sheets("MODO 1").Select

Application.ScreenUpdating = True

End Sub



Private Sub btAlterar_Click()

' Plan1/btAlterar/Click do 27-Registros_Cad.xlsm
' Sub para alterar os registros do cliente na Plan "MODO 1"
'
' Verifica (pelo nome) se o cliente existe
' Se existir, atualiza seus dados e finaliza

Application.ScreenUpdating = False

Dim WS As Worksheet
Dim WSCad As Worksheet
Dim Nome As String
    
Set WS = Sheets("Clientes")
Set WSCad = Sheets("MODO 1")
Nome = Sheets("MODO 1").CBoxBusca.Value
    
WS.Select
WS.Range("A2").Select

Do While ActiveCell.Value <> ""

    If ActiveCell = Nome Then
        
        'Modifica a linha do cliente de acordo com a plan "MODO 1"
        WS.Range(ActiveCell.Address, ActiveCell.Offset(0, 5).Address).Value = WSCad.Range("B5:G5").Value
        
        MsgBox "Registros atualizados com sucesso", vbOKOnly, "Status"
        
       'Atualiza os ComboBox's e ativa o ScreenUpdatind
        Call AttComboBox1
        
        'Sai da Sub
        Exit Sub

    End If

    ActiveCell.Offset(1, 0).Select

Loop

WSCad.Select
Range("B5").Select

End Sub




'***MODULO 1 (VBA)***

Option Explicit

Sub AttComboBox1()
'
' /Modulo1 do 26-Registros_Cad.xlsm
' Sub para atualizar o "ComboBox1" da Plan "MODO 1"
'

Application.ScreenUpdating = False

Dim WS As Worksheet

Set WS = Sheets("Clientes")

WS.Select
WS.Range("A2").Select

'É uma boa prática limpar valores de objetos como o CBox antes de usá-lo
Sheets("MODO 1").ComboBox1.Clear
Sheets("MODO 1").CBoxBusca.Clear

Do While ActiveCell.Value <> ""

    'Adicina celula ativa em ComboBox1 da plan "MODO1"
    Sheets("MODO 1").ComboBox1.AddItem ActiveCell.Value
    Sheets("MODO 1").CBoxBusca.AddItem ActiveCell.Value
    'Célula ativa desce uma linha
    ActiveCell.Offset(1, 0).Select

Loop

Sheets("MODO 1").Select
Range("B5").Select

Application.ScreenUpdating = True

End Sub




'***MODULO 2 (VBA)***

Option Explicit

Sub btCadastro2()
'
' /Modulo1 do 27-Registros_Cad.xlsm
' Sub para o botão "Cadastro" da Plan "MODO 2"
'

Dim WS As Worksheet
Set WS = Sheets("MODO 2")
     
Application.ScreenUpdating = False
    
WS.Select
    
Range("B5:G5").Select
Selection.Copy

Sheets("Clientes").Select
Range("A1048576").Select
Selection.End(xlUp).Select
ActiveCell.Offset(1, 0).Select
'Range("A3").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False   'Colar somente valores
'ActiveSheet.Paste

WS.Select
Application.CutCopyMode = False
Selection.ClearContents
Range("B5").Select

MsgBox "Cadastrado...", vbOKOnly, "Status"

Call AttVD2

Application.ScreenUpdating = True
    
End Sub



Sub btBusca2()
'
' /Modulo2 do 27-Registros_Cad.xlsm
' Sub para buscar cliente para a Plan "MODO 2"
'

Dim WS As Worksheet
Dim WSCad As Worksheet
Dim Nome As String

Set WS = Sheets("Clientes")
Set WSCad = Sheets("MODO 2")
Nome = Sheets("MODO 2").Range("F14").Text
' .Text retira o dado visível da célula, necessário para
' arrancar de células validadas

Application.ScreenUpdating = False

'Checa se validação de dados está vazia, se sim encerra
If Nome = "" Then
    
    MsgBox "Nenhum nome selecionado, busca abortada", vbOKOnly, "Status"
    Exit Sub

End If

WS.Select
WS.Range("A2").Select

'Percorre lista clientes enquanto <> ""
Do While ActiveCell.Value <> ""

    If Nome = ActiveCell.Value Then
       
       'WSCad.Range("B5").Value = ActiveCell.Value
       'WSCad.Range("C5").Value = ActiveCell.Offset(0, 1).Value
       'WSCad.Range("D5").Value = ActiveCell.Offset(0, 2).Value
       'WSCad.Range("E5").Value = ActiveCell.Offset(0, 3).Value
       'WSCad.Range("F5").Value = ActiveCell.Offset(0, 4).Value
       'WSCad.Range("G5").Value = ActiveCell.Offset(0, 5).Value
       
       WSCad.Range("B5:G5").Value = _
       WS.Range(ActiveCell.Address, ActiveCell.Offset(0, 5).Address).Value
       
       Exit Do
    
    End If
    
    ActiveCell.Offset(1, 0).Select
    
Loop

WSCad.Select
WSCad.Range("B5").Select

Application.ScreenUpdating = True

End Sub



Sub btAltera2()

' /Modulo2 do 27-Registros_Cad.xlsm
' Sub para alterar os registros do cliente na Plan "MODO 2"
'
' Verifica (pelo nome) se o cliente existe
' Se existir, atualiza seus dados e finaliza

Application.ScreenUpdating = False

Dim WS As Worksheet
Dim WSCad As Worksheet
Dim Nome As String
    
Set WS = Sheets("Clientes")
Set WSCad = Sheets("MODO 2")
Nome = Sheets("MODO 2").Range("F14").Value
    
WS.Select
WS.Range("A2").Select

Do While ActiveCell.Value <> ""

    If ActiveCell = Nome Then
        
        'Modifica a linha do cliente de acordo com a plan "MODO 2"
        WS.Range(ActiveCell.Address, ActiveCell.Offset(0, 5).Address).Value = _
        WSCad.Range("B5:G5").Value
        
        MsgBox "Registros atualizados com sucesso", vbOKOnly, "Status"
        
        WSCad.Select
        WSCad.Range("B5:G5").ClearContents
        WSCad.Range("B5").Select
        
        Call AttVD2
        
        'Sai da Sub
        Exit Sub

    End If

    ActiveCell.Offset(1, 0).Select

Loop

WSCad.Select
Range("B5").Select

Application.ScreenUpdating = True

End Sub


Sub btExcluir2()
'
' /Modulo2 do 27-Registros_Cad.xlsm
' Sub para excluir cliente pela célula com validação
' de dados da Plan "MODO 2"
'

Dim WS As Worksheet
Dim Nome As String

Set WS = Sheets("Clientes")
'Text no lugar de Value ou String "Nome" ficará vazia
'https://stackoverflow.com/questions/16820553/excel-cell-value-as-string-wont-store-as-string
Nome = Range("B14").Text

Application.ScreenUpdating = False

'Checa se B14 está vazio, se sim encerra
If Nome = "" Then
    
    MsgBox "Nenhum nome selecionado, exclusão abortada", vbOKOnly, "Status"
    Exit Sub

End If

WS.Select
WS.Range("A2").Select

Do While ActiveCell.Value <> ""

    If Nome = ActiveCell.Value Then
        
        'Deleta linha da Tabela
        ActiveCell.EntireRow.Delete
        MsgBox "Cadastro Apagado", vbOKOnly, "Status"
        
        Call AttVD2
                
        Exit Do
    
    End If
    
    ActiveCell.Offset(1, 0).Select
    
Loop


Sheets("MODO 2").Select

Application.ScreenUpdating = True

End Sub


Sub AttVD2()

Sheets("MODO 2").Select

Range("F14").Value = ""
Range("B14").Value = ""

End Sub