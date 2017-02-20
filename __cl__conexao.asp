<%
'*************************************************************************
'*************************************************************************
'****                                                                 ****
'****        Exemplo de utilização da Classe                          ****
'****                                                                 ****
'****        Dim oConexao                                             ****
'****        Set oConexao = New Conexao                               ****
'****        oConexao.AbreConexao()                                   ****
'****                                                                 ****
'*************************************************************************
'*************************************************************************

Class Conexao
	'*************************************************************************
	'*************************************************************************
	'****                                                                 ****
	'****        Declaração das variáveis                                 ****
	'****                                                                 ****
	'*************************************************************************
	'*************************************************************************
	Private hostBaseDeDados
	Private userBaseDeDados
	Private dataBaseDeDados
	Private senhaBaseDeDados
	Private debug

	public conex
	Private conectionString
	Private driveODBC

	Private rsRecuperaDados

	'*************************************************************************

	Private Sub VariaveisDaConexao()
	'*************************************************************************
	'*************************************************************************
	'****                                                                 ****
	'****        Informe aqui os dados da Base de dados MySql criada      ****
	'****        no servidor.                                             ****
	'****                                                                 ****
	'*************************************************************************
	'*************************************************************************
		hostBaseDeDados 	= "localhost" 	'Incluir o Host
		userBaseDeDados 	= "root"		'Incluir o User
		senhaBaseDeDados 	= ""			'Incluir o Password
		dataBaseDeDados 	= "agenda"		'Incluir o DataBase

		debug 				= "" 			'Valor default é VAZIO que seria o mesmo que desativado.
				   				 			'caso seja necessário ativar o debug, utilizar "ativado".
	End Sub

	'*************************************************************************

	Function AbreConexao()
	'*************************************************************************
	'*************************************************************************
	'****                                                                 ****
	'****        Rotina identifica o Drive ODBC disponível no servidor    ****
	'****        e abre a conexão com a base de dados MySQL informada     ****
	'****        na sub Function VariaveisDaConexao()                     ****
	'****                                                                 ****
	'*************************************************************************
	'*************************************************************************
		VariaveisDaConexao()		
		Set conex = CreateObject("ADODB.Connection")
		tiposDeconexoes = Array("MySQL ODBC 2.5 Driver",_
								"MySQL ODBC 3.5 Driver",_
								"MySQL ODBC 3.51 Driver",_
								"MySQL ODBC 5.1 Driver",_
								"MySQL ODBC 5.2 Driver",_
								"MySQL ODBC 5.3 Driver",_
								"MySQL ODBC 2.5 UNICODE Driver",_
								"MySQL ODBC 3.5 UNICODE Driver",_
								"MySQL ODBC 3.51 UNICODE Driver",_
								"MySQL ODBC 5.1 UNICODE Driver",_
								"MySQL ODBC 5.2 UNICODE Driver",_
								"MySQL ODBC 5.3 UNICODE Driver",_
								"MySQL ODBC 2.5 ANSI Driver",_
								"MySQL ODBC 3.5 ANSI Driver",_
								"MySQL ODBC 3.51 ANSI Driver",_
								"MySQL ODBC 5.1 ANSI Driver",_
								"MySQL ODBC 5.2 ANSI Driver",_
								"MySQL ODBC 5.3 ANSI Driver")

		For Each driveODBC in tiposDeconexoes
			
			On Error Resume Next			
			conectionString = "driver="& driveODBC &";server="&hostBaseDeDados&";uid="&userBaseDeDados&";pwd="&senhaBaseDeDados&";database="&dataBaseDeDados
			conex.Open conectionString
			if 0=Err.Number Then
				ef_Teste = "drvEncontrado"
				Exit For
			Else
				if debug="ativado" then
					response.write("Error: "&Err.Description&" - O Driver testado foi: <span style='color:red'>"&driveODBC&"</span><br>")
				end if
			End if
			On Error Goto 0

		Next

		if debug="ativado" then
			if ef_Teste = "drvEncontrado" Then
				response.write("Conectado com o MySql usando o Driver: "&driveODBC&"<br>")
			Else
				response.write("A conexão com o MySql falhou<br>")
			End if
		End if
	End Function

	'*************************************************************************

	Function FechaConexao()
		conex.Close
		if debug="ativado" then
			response.write("<br>Conexao Fechada!")
		end if
	End Function

End Class	

%>