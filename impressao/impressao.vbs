'  Script para impressão de documentos ( Windows ) 
'  Modificado por Rafael Zottesso
'  Adaptado por Daniel Mendes
'  ******************************

'Cria a variável para definir a impressora padrão

'Try
Set objPrinter = CreateObject("WScript.Network")

'Para impressora da rede utilize "\servidorNome da Impressora". Sem passar este parametro, o script pega a impressora padrão instalada no computador onde este script irá ser executado

'objPrinter.SetDefaultPrinter "NUMERO-IP\\NOME-IMPRESSORA\"



' Define o diretório
TargetFolder = "C:\arquivos" 'Alterar caso seja necessário, pois esta é a pasta onde o script lê para buscar os arquivos e enviar para a impressora

Set objShell = CreateObject("Shell.Application")

Set objFolder = objShell.Namespace(TargetFolder)

' Lista os arquivos
Set colItems = objFolder.Items

'Lê os arquivos encontrados na pasta indicada na variável TargetFolder
For Each objItem in colItems
	' Imprime os arquivos encontrados, ou seja, chama a impressora 
	objItem.InvokeVerbEx("Print")
	MsgBox("Impressao do arquivo realiza")

Next