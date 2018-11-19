<%
' Template em ASP Classic
'
' Licensed under MIT (https://github.com/plecyonahay/ASP-MVC-Template/blob/master/LICENSE)
'
' O Template permite manter o código HTML livres de códigos ASP.
' Desta forma, é possível manter a programação lógica (código ASP) longe da estrutura visual (HTML, CSS, etc).
'
' @author  Plecyo Nahay (plecyonahay@gmail.com)
' @version 1.3
' @project https://github.com/plecyonahay/ASP-MVC-Template
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

	' definição do modo de comparação das chaves das coleções
	Private p_dictionary_compare_mode

	' Expressão regular para localizar nomes de variáveis e blocos
	' Somente caracteres alfanuméricos e underline são permitidos
	Private	p_regex
	Private p_regex_search
	
	
	' construtor da classe 
	Private Sub Class_Initialize
		'0 -> Binary Comparison
		'1 -> Text Comparison
		'2 -> Compare information inside database
		p_dictionary_compare_mode = 1
		
		Set p_vars = CreateObject("Scripting.Dictionary")
			p_vars.CompareMode = p_dictionary_compare_mode
			
		Set p_values = CreateObject("Scripting.Dictionary")
			p_values.CompareMode = p_dictionary_compare_mode
			
		Set p_properties = CreateObject("Scripting.Dictionary")
			p_properties.CompareMode = p_dictionary_compare_mode
			
		Set p_instances = CreateObject("Scripting.Dictionary")
			p_instances.CompareMode = p_dictionary_compare_mode
		
		Set p_modifiers = CreateObject("Scripting.Dictionary")
			p_modifiers.CompareMode = p_dictionary_compare_mode
			
		Set p_blocks = CreateObject("Scripting.Dictionary")
			p_blocks.CompareMode = p_dictionary_compare_mode
			
		Set p_parents = CreateObject("Scripting.Dictionary")
			p_parents.CompareMode = p_dictionary_compare_mode
			
		Set p_parsed = CreateObject("Scripting.Dictionary")
			p_parsed.CompareMode = p_dictionary_compare_mode
		
		Set p_finally = CreateObject("Scripting.Dictionary")
			p_finally.CompareMode = p_dictionary_compare_mode
		
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
		If Not p_vars.Exists("{"&varname&"}") Then
			Call setError("setVariable", "Variável '"&varname&"' não existe")
		End If
		
		Call setValue(varname, value)
	End Function
	
	
	' retorna o valor da variável/objeto dentro da coleção
	Public Function getVariable(varname)
		If p_values.Exists("{"&varname&"}") Then
			getVariable = p_values.Item("{"&varname&"}")
		Else
			Call setError("getVariable", "Variável '"&varname&"' não existe")
		End If
	End Function
	
	
	' verifica se existe uma variável de template
	Public Function exists(varname)
		If p_vars.Exists("{"&varname&"}") Then
			exists = True
		Else
			exists = False
		End If
	End Function
	
	
	' verifica se existe um bloco
	Public Function existsBlock(blockname)
		If p_blocks.Exists(blockname) Then
			existsBlock = True
		Else
			existsBlock = False
		End If
	End Function
	
	
	' carrega um arquivo identificado pelo "filename"
	Private Function loadFile(varname, filename)
		If filename = "" Then
			Call setError("loadFile", "Informe o nome do arquivo para a variável '"&varname&"'")
		End If

		Dim filepath : filepath = Server.MapPath(filename)
		
		Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
		If  objFSO.FileExists(filepath) Then
			Dim objTextFile : Set objTextFile = CreateObject("ADODB.Stream")
				objTextFile.CharSet = "utf-8"
				objTextFile.Open
				objTextFile.LoadFromFile(filepath)
			
			Dim contentFile : contentFile = objTextFile.ReadText()

			Set objTextFile = Nothing
			
			If (contentFile&"") = "" Then
				Call setError("loadFile", "Arquivo '"&filepath&"' está vazio")
			End If
			
			Call setValue(varname, contentFile)
			
			Dim blocks_identified : Set blocks_identified = identify(contentFile, varname)
			
			Call createBlocks(blocks_identified)
		Else
			Call setError("loadFile", "Arquivo não existe no caminho '"&filepath&"'")
		End If
		
		Set objFSO = Nothing
	End Function
	
	
	' identifica todos os blocos e variáveis automaticamente
	Private Function identify(content, varname)
		Dim blocks : Set blocks = CreateObject("Scripting.Dictionary")
			blocks.CompareMode = p_dictionary_compare_mode

		Dim queued_blocks : Set queued_blocks = CreateObject("Scripting.Dictionary")
			queued_blocks.CompareMode = p_dictionary_compare_mode
	
		Call identifyVars(content)
		
		Dim line
		Dim lines : lines = Split(content, vbCrLf)
		For Each line In lines
			If InStr(line, "<!--") > 0 Then
				
				'BEGIN Block
				With p_regex
					.Global = False
					.MultiLine = False
					.IgnoreCase = False
					.Pattern = "\<!--\s+(BEGIN\s+("&p_regex_search&"))\s+\-->"
				End With
				Dim matches_begin : Set matches_begin = p_regex.Execute(line)
				If matches_begin.Count > 0 Then
					Dim match_value : match_value = matches_begin.Item(0).SubMatches.Item(1)
					
					Dim parent

					If queued_blocks.Count = 0 Then
						parent = varname
					Else
						parent = queued_blocks.Keys()(queued_blocks.Count - 1)
					End If

					If Not blocks.Exists(parent) Then
						blocks.Add parent, match_value
					Else
						blocks.Item(parent) = blocks.Item(parent)&","&match_value
					End If
					
					If queued_blocks.Exists(match_value) Then
						Call setError("identify", "Bloco '"&match_value&"' está duplicado")
					Else
						queued_blocks.Add match_value, match_value
					End If
				End If
				Set matches_begin = Nothing
				

				'END Block
				With p_regex
					.Global = False
					.MultiLine = False
					.IgnoreCase = False
					.Pattern = "\<!--\s+(END\s+("&p_regex_search&"))\s+\-->"
				End With
				Dim matches_end : Set matches_end = p_regex.Execute(line)
				If matches_end.Count > 0 Then
					If queued_blocks.Count > 0 Then
						Dim last_block : last_block = queued_blocks.Keys()(queued_blocks.Count - 1)
						If queued_blocks.Exists(last_block) Then
							queued_blocks.Remove(last_block)
						End If
					End If
				End If
				Set matches_end = Nothing
				
			End If
		Next
		
		Set identify = blocks
	End Function
	
	
	' identifica todas as variáveis definidas no documento
	Private Function identifyVars(content)
		With p_regex
			.Global = True
			.MultiLine = True
			.IgnoreCase = False
			.Pattern = "\{("&p_regex_search&")((\-\>("&p_regex_search&"))*)?((\|.*?)*)?\}"
		End With
		Dim matches : Set matches = p_regex.Execute(content)

		Dim objMatch
		For Each objMatch In matches
			If objMatch.SubMatches.Count > 0 Then
				Dim matchVar : matchVar = Trim(objMatch.SubMatches(0))

				' Variáveis como Objetos: {OBJ->ATTRIBUTE}
				Dim matchProperties : matchProperties = ""
				If objMatch.SubMatches.Count > 2 Then
					matchProperties = Trim(objMatch.SubMatches(2))

					If Not p_properties.Exists(matchVar) Then
						p_properties.Add matchVar, matchProperties
					Else
						If InStr(p_properties.Item(matchVar), matchProperties) = 0 Then
							p_properties.Item(matchVar) = p_properties.Item(matchVar) &","& matchProperties
						End If
					End If
				End If
				
				
				' Variáveis com Modificadores: {OBJ->ATTRIBUTE|lCase} ou {ATTRIBUTE|lCase}
				Dim matchModifiers : matchModifiers = ""
				If objMatch.SubMatches.Count > 6 Then
					matchModifiers = Trim(objMatch.SubMatches(6))

					If Not p_modifiers.Exists(matchVar) Then
						p_modifiers.Add matchVar&matchProperties, matchVar&matchProperties&matchModifiers
					Else
						If InStr(p_modifiers.Item(matchVar&matchProperties), matchModifiers) = 0 Then 
							p_modifiers.Item(matchVar&matchProperties) = p_modifiers.Item(matchVar&matchProperties) &","& matchVar&matchProperties&matchModifiers
						End If
					End If
				End If
			
				' Variáveis comuns: {ATTRIBUTE}
				If Not p_vars.Exists("{"&matchVar&"}") Then
					p_vars.Add "{"&matchVar&"}", "{"&matchVar&"}"
				End If
			End If
		Next
		set objMatch = Nothing
		Set matches = Nothing
	End Function
	
	
	' Cria todos os blocos identificados por identifyBlocks()
	Private Function createBlocks(blocks)
		Dim i
		For i = 0 to blocks.Count - 1
			Dim parent : parent = blocks.Keys()(i)
			Dim block_ : block_ = blocks.Items()(i)
			
			If Not p_parents.Exists(parent) Then
				p_parents.Add parent, block_
			End If
			
			Dim arr_block : arr_block = Split(block_, ",")
			
			Dim j
			For j = 0 to Ubound(arr_block)
				Dim child : child = arr_block(j)
				
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
		Dim str : str = getVar(parent)
		Dim name : name = block&"_value"
		
		Call setValue(name, "")


		'BEGIN Block
		With p_regex
			.Global = False
			.IgnoreCase = False
			.Pattern = "\<!--\s+(BEGIN\s+("&block&"))\s+\-->"
		End With

		Dim matches_begin : Set matches_begin = p_regex.Execute(str)
		Dim block_begin
		If matches_begin.Count > 0 Then
			block_begin = matches_begin.Item(0)
		Else
			Call setError("setBlock", "Prefixo 'BEGIN' não encontrado no bloco '"&block&"'")
		End If
		Set matches_begin = Nothing
		

		'END Block
		With p_regex
			.Global = False
			.IgnoreCase = False
			.Pattern = "\<!--\s+(END\s+("&block&"))\s+\-->"
		End With		

		Dim matches_end : Set matches_end = p_regex.Execute(str)
		Dim block_end
		If matches_end.Count > 0 Then
			block_end = matches_end.Item(0)
		Else
			Call setError("setBlock", "Prefixo 'END' não encontrado no bloco '"&block&"'")
		End If
		

		'FINALLY Block
		With p_regex
			.Global = False
			.IgnoreCase = False
			.Pattern = "\<!--\s+(FINALLY\s+("&block&"))\s+\-->"
		End With

		Dim matches_finally : Set matches_finally = p_regex.Execute(str)
		Dim block_finally
		If matches_finally.Count > 0 Then
			block_finally = matches_finally.Item(0)
		Else
			block_finally = ""
		End If
		

		' Verifica onde inicia e termina cada block
		Dim v_begin : v_begin = InStr(str, block_begin) + Len(block_begin)
		Dim v_end   : v_end   = InStr(str, block_end)   - v_begin
		
		Dim block_content : block_content = Mid(str, v_begin, v_end)
		
		Call setValue(block, block_content)
		
		Dim v_finally
		If block_finally = "" Then
			v_finally = 0
		Else
			v_finally = InStr(str, block_finally)
		End If
		
		Dim block_replace
		If v_finally > 0 Then
			v_end     = InStr(str, block_end) + Len(block_end)
			v_finally = v_finally - v_end
		
			p_finally.Add block, Mid(str, v_end, v_finally)

			block_replace = Replace(str, Mid(str, InStr(str, block_begin), ((InStr(str, block_finally)+Len(block_finally))-InStr(str, block_begin))), "{"&name&"}" )
		Else
			block_replace = Replace(str, Mid(str, InStr(str, block_begin), ((InStr(str, block_end)+Len(block_end))-InStr(str, block_begin))), "{"&name&"}" )
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
		Dim statements : statements = Split(exp, "|")

		Dim function_value : function_value = Chr(34) & value & Chr(34)

		Dim i
		For i = 1 to Ubound(statements)
			Dim temp : temp = Split(statements(i), ":")
			
			Dim function_name : function_name = temp(0)
			temp(0) = function_value

			Dim function_param
				function_param = Join(temp, ",")
				function_param = Replace(function_param, "'", """")

			value = function_name&"("&function_param&")"
			function_value = value
		Next

		On Error Resume Next
		Err.Clear
		
		Dim subst_value : subst_value = Eval(value)
		
		If Err.Number <> 0 Then
			Call setError("substModifiers", Err.Source&" - "&Err.Description&" ("&value&")")
			Err.Clear
		End If
		On Error Goto 0
		
		substModifiers = subst_value
	End Function
	
	
	' retorna em formato camelCase a propriedade. EX: .getIdExemplo
	Private Function camelCasePropertyObject(property_value)
		Dim arrProperty : arrProperty = Split(property_value, "_")
		
		Dim i
		For i = 0 to Ubound(arrProperty)
			arrProperty(i) = Left(arrProperty(i), 1) & lCase(Right( arrProperty(i), Len(arrProperty(i))-1 ))
		Next

		Dim property_name : property_name = ".get" & Join(arrProperty, "") & "()"
		
		camelCasePropertyObject = property_name
	End Function
	
	
	' Preenche todas as variáveis contidas na variável
	Private Function subst(value)
		Dim subst_content : subst_content = value
		
		' Variáveis comuns
		If p_values.Count > 0 Then
			Dim iValue
			For iValue = 0 to p_values.Count - 1
				subst_content = Replace(subst_content, p_values.Keys()(iValue), p_values.Items()(iValue))
			Next
		End If
		
		' Variáveis com Modificadores
		If p_modifiers.Count > 0 Then
			Dim iModify
			For iModify = 0 to p_modifiers.Count - 1				
				If InStr(subst_content, "{"&p_modifiers.Keys()(iModify)&"|") > 0 Then
					If p_values.exists("{"&p_modifiers.Keys()(iModify)&"}") Then
						subst_content = Replace(subst_content, "{"&p_modifiers.Items()(iModify)&"}", substModifiers(p_values.Item("{"&p_modifiers.Keys()(iModify)&"}"), p_modifiers.Items()(iModify)) )
					End If
				End If
			Next
		End If
		
		' Variáveis como Objeto
		If p_instances.Count > 0 Then
			Dim i
			For i = 0 to p_instances.Count - 1
				Dim var : var = p_instances.Keys()(i)
				If p_properties.Exists(var) Then
					Dim arrProperties : arrProperties = Split(p_properties.Item(var), ",")
					Dim j
					For j = 0 to Ubound(arrProperties)
						If InStr(subst_content, "{"&var&arrProperties(j)&"}") > 0 OR InStr(subst_content, "{"&var&arrProperties(j)&"|") > 0 Then
							Dim pointer : pointer = ""
							Dim objClass : Set objClass = p_instances.Item(var)
							If Not IsNull(objClass) Then
								Dim arrPropertiesIn : arrPropertiesIn = Split(arrProperties(j), "->")
								Dim k
								For k = 1 to Ubound(arrPropertiesIn)
									Dim objPropertyName : objPropertyName = camelCasePropertyObject(arrPropertiesIn(k))
									
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
							
							Dim value_attribute

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
		If Not p_blocks.Exists(value) Then
			Call setError("block", "Bloco "&value&" não existe")
		End If
		
		Dim arr_children

		' verifica os blocos dentro de outros blocos
		If p_parents.Exists(value) Then
			children = p_parents.Item(value)
			
			arr_children = Split(children, ",")
			
			Dim i
			For i = 0 to Ubound(arr_children)
				Dim child : child = arr_children(i)
				
				If p_finally.Exists(child) AND NOT p_parsed.Exists(child) Then
					Call setValue(child&"_value", subst(p_finally.Item(child)))
					p_parsed.Item(value) = value
				End If
			Next
		End If
		
		Call setValue(value&"_value", getVar(value&"_value")&subst(getVar(value)))
		
		If Not p_parsed.Exists(value) Then
			p_parsed.Item(value) = value
		End If
		
		' limpando blocos filhos
		If p_parents.Exists(value) Then
			children = p_parents.Item(value)
			
			arr_children = Split(children, ",")
			For j = 0 to Ubound(arr_children)
				Call clear(arr_children(j)&"_value")
			Next
		End If
	End Function
	
	
	' retorna o conteúdo final
	Public Function parse()
		' auto-assistência para blocos "filhos"
		If p_parents.Count > 0 Then
			p_parents_qtd = p_parents.Count - 1
			For i = p_parents_qtd to 0 step -1
				Dim parent : parent = p_parents.Keys()(i)
				Dim children : children = p_parents.Items()(i)
				
				Dim arr_children : arr_children = Split(children, ",")
				For j = 0 to Ubound(arr_children)
					If p_blocks.Exists(parent) AND p_parsed.Exists(arr_children(j)) AND Not p_parsed.Exists(parent) Then
						Call setValue(parent&"_value", subst(getVar(parent)))
						p_parsed.Item(parent) = parent
					End If
				Next
			Next
		End If
		
		' exibindo os blocos "finally", caso não tenha sido chamado algum bloco "filho" ou o "pai"
		If p_finally.Count > 0 Then
			For i = 0 to p_finally.Count - 1
				If Not p_parsed.Exists(p_finally.Keys()(i)) Then
					Call setValue(p_finally.Keys()(i)&"_value", subst(p_finally.Items()(i)))
				End If
			Next
		End If

		' remove as variáveis vazias
		Dim container_html : container_html = subst(getVar("."))
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
		Err.Raise 1, method, msg
	End Sub
	
	
	' destrutor da classe 
	Private Sub Class_Terminate()
		Set p_vars       = Nothing
		Set p_values     = Nothing
		Set p_properties = Nothing
		Set p_instances  = Nothing
		Set p_modifiers  = Nothing
		Set p_blocks     = Nothing
		Set p_parents    = Nothing
		Set p_parsed     = Nothing
		Set p_finally    = Nothing
		Set p_regex      = Nothing
	End Sub
	
End class
%>
