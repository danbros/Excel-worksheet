' Author			Date
' danbros			
'
'https://github.com/danbros/Excel-worksheet
'
'AULA 24 Bloquear acesso a uma plan
'Sub na plan selecionada ao inserir sub abaixo

Private Sub Worksheet_Activate()

	'desativa atualização de tela
	Application.ScreenUpdating = False
	    
	    'Se username atual(em maiusculo por causa do Ucase) for diferente de "USER OK" então...
	    If UCase(Environ("username")) <> "USER OK" Then
	    
	        Sheets("Planilha1").Activate
	        MsgBox "Seu usuário não tem permissão de abrir essa plan", vbOKOnly, "Atenção"
	                
	    End If
	    
	'ativa atualização de tela
	Application.ScreenUpdating = True

End Sub