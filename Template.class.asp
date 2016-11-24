<%
' Template em ASP Classic
'
' O Template permite manter o código HTML livres de códigos ASP.
' Desta forma, é possível manter a programação lógica (código ASP) longe da estrutura visual (HTML, CSS, etc).
'
' @author  Plecyo Nahay (plecyonahay@gmail.com)
' @version 1.0
'
Class Template

	' coleção de variáveis de documentos existentes
	Private p_vars

	' coleção de variáveis e valores definidos pelo usuário
	Private p_values

	' coleção de variáveis com propriedades de objeto existentes no documento
	Private p_properties

	' coleção das instâncias de objeto definidas pelo usuário
	Private p_instances

	' coleção de modificadores (funções)
	Private p_modifiers

	' coleção dos blocos reconhecidos automaticamente
	Private p_blocks

	' coleção dos blocos que contém pelo menos um bloco filho
	Private p_parents

	' coleção de blocos analisados
	Private p_parsed

	' coleção de blocos utilizando finally 
	Private p_finally

	' Expressão regular para localizar nomes de variáveis e blocos
	' Somente caracteres alfanuméricos e underline são permitidos
	Private	p_regex
	Private p_regex_search
	
	
	' construtor da classe 
	Private Sub Class_Initialize
		Set p_vars         = CreateObject("Scripting.Dictionary")
		Set p_values       = CreateObject("Scripting.Dictionary")
		Set p_properties   = CreateObject("Scripting.Dictionary")
		Set p_instances    = CreateObject("Scripting.Dictionary")
		Set p_modifiers    = CreateObject("Scripting.Dictionary")
		Set p_blocks       = CreateObject("Scripting.Dictionary")
		Set p_parents      = CreateObject("Scripting.Dictionary")
		Set p_parsed       = CreateObject("Scripting.Dictionary")
		Set p_finally      = CreateObject("Scripting.Dictionary")
		
		Set p_regex        = New RegExp
		    p_regex_search = "([A-Z0-9_])+"
	End Sub
	
	
	' informa o Template padrão a ser utilizado
	Public Function addTemplate(filename)
		Call loadFile(".", filename)
	End Function
	
	
	' insere o conteúdo do "filename" dentro da variável "varname"
	Public Function addFile(varname, filename)
		If p_vars.Exists(varname) Then
			Call setError("addFile", "Variável '"&varname&"' não existe")
		End If
		
		Call loadFile(varname, filename)
	End Function
	
	
	' insere o valor da variável/objeto dentro da coleção
	Public Function setVariable(varname, value)
		If NOT p_vars.Exists("{"&varname&"}") Then
			Call setError("setVariable", "Variável '"&varname&"' não existe")
		End If
		
		Call setValue(varname, value)
	End Function
	
	
	' verifica se existe uma variável de template
	Public Function exists(varname)
		If p_vars.Exists("{"&varname&"}") Then
			exists = True
		Else
			exists = False
		End If
	End Function
	
	
	' carrega um arquivo identificado pelo "filename"
	Private Function loadFile(varname, filename)
		Dim objFSO, FilePath, objTextFile, contentFile
		
		If filename = "" Then
			Call setError("loadFile", "Informe o nome do arquivo para a variável '"&varname&"'")
		Else
			FilePath = Server.MapPath(filename)
		End If
		
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		If  objFSO.FileExists(FilePath) Then
			
			Set objTextFile = CreateObject("ADODB.Stream")
				objTextFile.CharSet = "utf-8"
				objTextFile.Open
				objTextFile.LoadFromFile(FilePath)
				contentFile = objTextFile.ReadText()
			Set objTextFile = nothing
			
			If contentFile = "" Then
				Call setError("loadFile", "Arquivo '"&filename&"' está vazio")
			End If
			
			Call setValue(varname, contentFile)
			
			Dim blocks_identified
			Set blocks_identified = identify(contentFile, varname)
			
			Call createBlocks(blocks_identified)
		Else
			Call setError("loadFile", "Arquivo '"&filename&"' não existe no caminho '"&FilePath&"'")
		End If
		
		Set objFSO = nothing
	End Function
	
	
	' identifica todos os blocos e variáveis automaticamente
	Private Function identify(content, varname)
		Dim blocks, queued_blocks, lines, line, matches, match_value, parent, last_block
		
		Set blocks        = CreateObject("Scripting.Dictionary")
		Set queued_blocks = CreateObject("Scripting.Dictionary")
	
		Call identifyVars(content)
		
		lines = Split(content, vbCrLf)
		For each line in lines
			If InStr(line, "<!--") > 0 Then
				
				'BEGIN
				with p_regex
					.Global = False
					.MultiLine = False
					.IgnoreCase = False
					.Pattern = "\<!--\s+(BEGIN\s+("&p_regex_search&"))\s+\-->"
				End with		
				Set matches = p_regex.Execute(line)
				If matches.Count > 0 Then
					match_value = matches.Item(0).SubMatches.Item(1)
					
					If queued_blocks.Count = 0 Then
						parent = varname
					Else
						parent = queued_blocks.Keys()(queued_blocks.Count - 1)
					End If

					If NOT blocks.Exists(parent) Then
						blocks.Add parent, match_value
					Else
						blocks.Item(parent) = blocks.Item(parent)&","&match_value
					End If
					
					queued_blocks.Add match_value, match_value
				End If
				
				'END
				with p_regex
					.Global = False
					.MultiLine = False
					.IgnoreCase = False
					.Pattern = "\<!--\s+(END\s+("&p_regex_search&"))\s+\-->"
				End with
				Set matches = p_regex.Execute(line)
				If matches.Count > 0 Then
					If queued_blocks.Count > 0 Then
						last_block = queued_blocks.Keys()(queued_blocks.Count - 1)
						If queued_blocks.Exists(last_block) Then
							queued_blocks.Remove(last_block)
						End If
					End If
				End If
				
			End If
		Next
		
		Set identify = blocks
	End Function
	
	
	' identifica todas as variáveis definidas no documento
	Private Function identifyVars(content)
		Dim matches, matchVar
		
		with p_regex
			.Global = True
			.MultiLine = True
			.IgnoreCase = False
			 .Pattern = "\{("&p_regex_search&")((\-\>("&p_regex_search&"))*)?((\|.*?)*)?\}"
		End with
		Set matches = p_regex.Execute(content)
		
		For each objMatch in matches
			If objMatch.SubMatches.Count > 0 Then
				matchVar        = Trim(objMatch.SubMatches(0))
				matchProperties = ""
				matchModifiers  = ""
				
				If objMatch.SubMatches.Count > 2 Then
					matchProperties = Trim(objMatch.SubMatches(2))
				End If
				
				If objMatch.SubMatches.Count > 6 Then
					matchModifiers = Trim(objMatch.SubMatches(6))
				End If
				
				' Objetos
				If matchProperties <> "" Then
					If NOT p_properties.Exists(matchVar) Then
						p_properties.Add matchVar, matchProperties
					Else
						If InStr(p_properties.Item(matchVar), matchProperties) = 0 Then
							p_properties.Item(matchVar) = p_properties.Item(matchVar) &","& matchProperties
						End If
					End If
				End If
				
				' Modificadores
				If matchModifiers <> "" Then
					If NOT p_modifiers.Exists(matchVar) Then
						p_modifiers.Add matchVar&matchProperties, matchVar&matchProperties&matchModifiers
					Else
						If InStr(p_modifiers.Item(matchVar&matchProperties), matchModifiers) = 0 Then 
							p_modifiers.Item(matchVar&matchProperties) = p_modifiers.Item(matchVar&matchProperties) &","& matchVar&matchProperties&matchModifiers
						End If
					End If
				End If
			
				' Variáveis comuns
				If NOT p_vars.Exists("{"&matchVar&"}") Then
					p_vars.Add "{"&matchVar&"}", "{"&matchVar&"}"
				End If
			End If
		Next
	End Function
	
	
	' Cria todos os blocos identificados por identifyBlocks()
	Private Function createBlocks(blocks)
		Dim i, j, parent, block_, arr_block, child
		
		For i = 0 to blocks.Count - 1
			parent = blocks.Keys()(i)
			block_ = blocks.Items()(i)
			
			If NOT p_parents.Exists(parent) Then
				p_parents.Add parent, block_
			End If
			
			arr_block = Split(block_, ",")
			
			For j = 0 to Ubound(arr_block)
				child = arr_block(j)
				
				If p_blocks.Exists(child) Then
					Call setError("createBlocks", "Bloco '"&child&"' está duplicado")
				End If
				
				p_blocks.Add child, child
				
				Call setBlock(parent, child)
			Next
		Next
	End Function
	
	
	' Uma variável "pai" pode conter um bloco de variável definido por:
	' <!---{BEGIN _varname_}---> content <!---{END _varname_}--->
	' Esse método remove o bloco "pai" e o substitui por uma referência variável denominada "block"
	Private Function setBlock(parent, block)
		Dim str, name, block_begin, block_end, block_finally, block_content, v_begin, v_end, v_finally, block_content_finally, block_replace
	
		str  = getVar(parent)
		name = block&"_value"
		
		Call setValue(name, "")

		'BEGIN
		with p_regex
			.Global = False
			.IgnoreCase = False
			.Pattern = "\<!--\s+(BEGIN\s+("&block&"))\s+\-->"
		End with		
		Set matches = p_regex.Execute(str)
		If matches.Count > 0 Then
			block_begin = matches.Item(0)
		Else
			Call setError("setBlock", "Prefixo 'BEGIN' não encontrado no bloco '"&block&"'")
		End If
		
		'END
		with p_regex
			.Global = False
			.IgnoreCase = False
			.Pattern = "\<!--\s+(END\s+("&block&"))\s+\-->"
		End with		
		Set matches = p_regex.Execute(str)
		If matches.Count > 0 Then
			block_end = matches.Item(0)
		Else
			Call setError("setBlock", "Prefixo 'END' não encontrado no bloco '"&block&"'")
		End If
		
		'FINALLY
		with p_regex
			.Global = False
			.IgnoreCase = False
			.Pattern = "\<!--\s+(FINALLY\s+("&block&"))\s+\-->"
		End with		
		Set matches = p_regex.Execute(str)
		If matches.Count > 0 Then
			block_finally = matches.Item(0)
		Else
			block_finally = ""
		End If
		
		v_begin       = InStr(str, block_begin) + Len(block_begin)
		v_end         = InStr(str, block_end) - v_begin
		
		block_content = Mid(str, v_begin, v_end)
		
		Call setValue(block, block_content)
		
		If block_finally = "" Then
			v_finally = 0
		Else
			v_finally = InStr(str, block_finally)
		End If
		
		If v_finally > 0 Then
			v_end     = InStr(str, block_end) + Len(block_end)
			v_finally = v_finally - v_end
			
			block_content_finally = Mid(str, v_end, v_finally)
			block_replace         = Replace(str, Mid(str, InStr(str, block_begin), ((InStr(str, block_finally)+Len(block_finally))-InStr(str, block_begin))), "{"&name&"}" )
			
			p_finally.Add block, block_content_finally
		Else
			block_content_finally = ""
			block_replace         = Replace(str, Mid(str, InStr(str, block_begin), ((InStr(str, block_end)+Len(block_end))-InStr(str, block_begin))), "{"&name&"}" )
		End If
	    
		Call setValue(parent, block_replace)
	End Function
	
	
	' definindo o valor de uma variável
	Private Function setValue(varname, value)
		If IsObject(value) Then
			Set p_instances.Item(varname) = value
		Else
			p_values.Item("{"&varname&"}") = value & ""
		End If
	End Function
	
	
	' Retorna o valor da variável
	Private Function getVar(varname)
		getVar = p_values.Item("{"&varname&"}")
	End Function
	
	
	' Limpa o valor de uma variável
	Public Function clear(varname)
		p_values.Item("{"&varname&"}") = ""
	End Function
	
	
	' atribuir manualmente um bloco "filho" a um bloco "pai"
	Public Function setParent(parent, block)
		p_parents.Item(parent) = block
	End Function
	
	
	' substituir modificadores de conteúdo
	Public Function substModifiers(value, exp)
		Dim i, statements, temp, function_name, function_value, function_param, subst_value
		
		statements = Split(exp, "|")
		function_value = Chr(34) & value & Chr(34)
		
		For i = 1 to Ubound(statements)
			temp = Split(statements(i), ":")
			function_name = temp(0)
			temp(0) = function_value
			function_param = Join(temp, ",")
			value = function_name&"("&function_param&")"
			function_value = value
		Next

		On Error Resume Next
		Err.Clear
		
		subst_value = Eval(value)
		
		If Err.Number <> 0 Then
			Call setError("substModifiers", Err.Source&" - "&Err.Description&" ("&value&")")
			Err.Clear
		End If
		On Error Goto 0
		
		substModifiers = subst_value
	End Function
	
	
	' retorna em formato camelCase a propriedade. EX: .getIdExemplo
	Private Function camelCasePropertyObject(property_value)
		Dim i, arrProperty, property_name
		
		arrProperty = Split(property_value, "_")
		For i = 0 to Ubound(arrProperty)
			arrProperty(i) = Left(arrProperty(i), 1) & lCase(Right( arrProperty(i), Len(arrProperty(i))-1 ))
		Next
		property_name = ".get" & Join(arrProperty, "") & "()"
		
		camelCasePropertyObject = property_name
	End Function
	
	
	' Preenche todas as variáveis contidas na variável
	Private Function subst(value)
		Dim i, j, k, subst_content, var, arrProperties, arrPropertiesIn, objClass, objPropertyName, pointer, value_attribute
		
		subst_content = value
		
		' Variáveis comuns
		If p_values.Count > 0 Then
			For i = 0 to p_values.Count - 1
				subst_content = Replace(subst_content, p_values.Keys()(i), p_values.Items()(i))
			Next
		End If
		
		' Modificadores
		If p_modifiers.Count > 0 Then
			For i = 0 to p_modifiers.Count - 1				
				If InStr(subst_content, "{"&p_modifiers.Keys()(i)&"|") > 0 Then
					If p_values.exists("{"&p_modifiers.Keys()(i)&"}") Then
						subst_content = Replace(subst_content, "{"&p_modifiers.Items()(i)&"}", substModifiers(p_values.Item("{"&p_modifiers.Keys()(i)&"}"), p_modifiers.Items()(i)) )
					End If
				End If
			Next
		End If
		
		' Objetos
		If p_instances.Count > 0 Then
			For i = 0 to p_instances.Count - 1
				var = p_instances.Keys()(i)
				If p_properties.Exists(var) Then
					arrProperties = Split(p_properties.Item(var), ",")
					For j = 0 to Ubound(arrProperties)
						If InStr(subst_content, "{"&var&arrProperties(j)&"}") > 0 OR InStr(subst_content, "{"&var&arrProperties(j)&"|") > 0 Then
							pointer = ""
							Set objClass = p_instances.Item(var)
							If NOT IsNull(objClass) Then
								arrPropertiesIn = Split(arrProperties(j), "->")
								For k = 1 to Ubound(arrPropertiesIn)
									objPropertyName = camelCasePropertyObject(arrPropertiesIn(k))
									
									On Error Resume Next
									Err.Clear
									
									If IsObject( objClass ) Then
										If IsObject( Eval("objClass"&objPropertyName) ) Then
											Set objClass = Eval("objClass"&objPropertyName)
										Else
											pointer = Eval("objClass"&objPropertyName)
										End If
									Else
										pointer = Eval("objClass"&objPropertyName)
									End If
		
									If Err.Number <> 0 Then
										Call setError("subst", Err.Source&" - "&Err.Description&" ("&TypeName(objClass)&objPropertyName&")")
										Err.Clear
									End If
									On Error Goto 0
		
								Next
							End If
							
							If TypeName( pointer ) = "Integer" OR _
							   TypeName( pointer ) = "Long" OR _
							   TypeName( pointer ) = "Single" OR _
							   TypeName( pointer ) = "Double" OR _
							   TypeName( pointer ) = "Currency" OR _
							   TypeName( pointer ) = "Decimal" OR _
							   TypeName( pointer ) = "Date" OR _
							   TypeName( pointer ) = "String" OR _
							   TypeName( pointer ) = "Boolean" OR _
							   TypeName( pointer ) = "Empty" OR _
							   TypeName( pointer ) = "Null" Then
								value_attribute = pointer
							Else
								value_attribute = TypeName(pointer)
							End If
							
							subst_content = Replace(subst_content, "{"&p_instances.Keys()(i)&arrProperties(j)&"}", value_attribute )

							If p_modifiers.exists(p_instances.Keys()(i)&arrProperties(j)) Then
								subst_content = Replace(subst_content, "{"&p_modifiers.item(p_instances.Keys()(i)&arrProperties(j))&"}", substModifiers(pointer, p_modifiers.item(p_instances.Keys()(i)&arrProperties(j))) )
							End If
						End If
					Next
				End If
			Next
		End If
		
		subst = subst_content
	End Function
	
	
	' exibe um bloco
	Public Function block(value)
		Dim i, arr_children, child
		
		If NOT p_blocks.Exists(value) Then
			Call setError("block", "Bloco "&value&" não existe")
		End If
		
		' verifica os blocos dentro de outros blocos
		If p_parents.Exists(value) Then
			children = p_parents.Item(value)
			
			arr_children = Split(children, ",")
			
			For i = 0 to Ubound(arr_children)
				child = arr_children(i)
				
				If p_finally.Exists(child) AND NOT p_parsed.Exists(child) Then
					Call setValue(child&"_value", subst(p_finally.Item(child)))
					p_parsed.Item(value) = value
				End If
			Next
		End If
		
		Call setValue(value&"_value", getVar(value&"_value")&subst(getVar(value)))
		
		If NOT p_parsed.Exists(value) Then
			p_parsed.Item(value) = value
		End If
		
		' limpando blocos filhos
		If p_parents.Exists(value) Then
			children = p_parents.Item(value)
			
			arr_children = Split(children, ",")
			For j = 0 to Ubound(arr_children)
				child = arr_children(j)
				Call clear(child&"_value")
			Next
		End If
	End Function
	
	
	' retorna o conteúdo final
	Public Function parse()
		Dim parent, children, arr_children, child, content_finally, container_html
	
		' auto-assistência para blocos "filhos"
		If p_parents.Count > 0 Then
			p_parents_qtd = p_parents.Count - 1
			For i = p_parents_qtd to 0 step -1
				parent   = p_parents.Keys()(i)
				children = p_parents.Items()(i)
				
				arr_children = Split(children, ",")
				For j = 0 to Ubound(arr_children)
					child = arr_children(j)
					If p_blocks.Exists(parent) AND p_parsed.Exists(child) AND NOT p_parsed.Exists(parent) Then
						Call setValue(parent&"_value", subst(getVar(parent)))
						p_parsed.Item(parent) = parent
					End If
				Next
			Next
		End If
		
		' exibindo os blocos "finally", caso não tenha sido chamado algum bloco "filho" ou o "pai"
		If p_finally.Count > 0 Then
			For i = 0 to p_finally.Count - 1
				If NOT p_parsed.Exists(p_finally.Keys()(i)) Then
					content_finally = subst(p_finally.Items()(i))
					Call setValue(p_finally.Keys()(i)&"_value", content_finally)
				End If
			Next
		End If

		' remove as variáveis vazias
		container_html = subst(getVar("."))
		with p_regex
			.Global = True
			.MultiLine = True
			.IgnoreCase = False
			.Pattern = "\{("&p_regex_search&")((\-\>("&p_regex_search&"))*)?((\|.*?)*)?\}"
		End with
		
		parse = p_regex.Replace(container_html, "")
	End Function
	
	
	' imprime o conteúdo final
	Public Function show()
		Response.Write parse()
	End Function
	
	
	' exibe na tela um erro ocorrido pela classe de Template
	Private Sub setError(method, msg)
		Response.Write("<pre><b>Template.class.asp - "&method&"()</b>" & vbCrLf & Server.HTMLEncode(msg) &"</pre>")
		Response.End()
	End Sub
	
	
	' destrutor da classe 
	Private Sub Class_Terminate()
		Set p_vars       = nothing
		Set p_values     = nothing
		Set p_properties = nothing
		Set p_instances  = nothing
		Set p_modifiers  = nothing
		Set p_blocks     = nothing
		Set p_parents    = nothing
		Set p_parsed     = nothing
		Set p_finally    = nothing
		Set p_regex      = nothing
	End Sub
	
End class
%>
