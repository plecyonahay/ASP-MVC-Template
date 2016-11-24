# ASP-MVC-Template

Tutorial de Templates em ASP
========

Através do uso de templates, deixamos toda a estrutura visual (HTML ou XML, CSS, etc) separado da lógica de programação (código ASP), o que melhora e muito tanto a construção quanto a manutenção de sistemas web. Existem poucos mecanismo de template para ASP no mercado. 

Com base nisso, eu resolvi criar um tutorial baseado em um mecanismo de template muito simples, baseado numa biblioteca que eu mesmo desenvolvi, e uso em meus projetos, na qual gastei bastante tempo desenvolvendo e melhorando, sempre com o foco principal na facilidade de uso.

Você verá que será preciso entender apenas duas idéias básicas: variáveis e blocos.

Então, vamos lá.


## Requisitos Necessários

É preciso ter implantando um servidor ASP clássico (IIS - Internet Information Services) versão igual ou superior a 6.0.


## Instalação e Uso

Incluir a classe Template da seguinte forma:

```asp
<!--#include file="Template.class.asp" -->
<%
    Dim tpl
    Set tpl = New Template
    Set tpl = Nothing
%>
```


## Exemplo e explicação: Olá Mundo

O funcionamento básico do mecanismo de templates é baseado na seguinte idéia: você tem um arquivo ASP (que usará a biblioteca), e este arquivo não conterá código HTML algum. O código HTML ficará separado, em um arquivo que conterá somente códigos HTML. Este arquivo HTML será lido pelo arquivo ASP.

Para evitar problemas com codificação de caracteres, procure definir o charset (encoding) UTF-8 para todos os documentos que você irá importar para o seu projeto (.HTML, .ASP, ...).

Dito isso, vamos ao primeiro exemplo, o já manjado Olá mundo. Vamos criar 2 arquivos: o ASP responsável por toda a lógica, e o arquivo HTML com nosso layout.

Então, crie um arquivo HTML, chamado hello.html com o conteúdo abaixo:

```html
<!DOCTYPE html>
<html lang="pt-br">
    <head>
        <meta charset="utf-8">
        <title>Olá Mundo</title>
    </head>
    <body>
        Olá Mundo, com templates ASP!
    </body>
</html>
```

Agora, crie o arquivo ASP, hello.asp:

```asp
<!--#include file="Template.class.asp" -->
<%
    Dim tpl
    Set tpl = New Template
        tpl.addTemplate "hello.html"
        tpl.show()
    Set tpl = Nothing
%>
```

Agora basta executar em seu navegador, o script hello.asp, e verificar que ele irá exibir o conteúdo lido de hello.html.

Se algo deu errado, consulte a seção sobre as Mensagens de Erro.


## Variáveis

Vamos agora a um conceito importante: variáveis de template. Como você pode imaginar, vamos querer alterar várias partes do arquivo HTML. Como fazer isso? Simples: no lado do HTML, você cria as chamadas variáveis de template. Veja o exemplo abaixo:

```html
<!DOCTYPE html>
<html lang="pt-br">
    <head>
        <meta charset="utf-8">
        <title>Olá Mundo</title>
    </head>
    <body>
        Olá {FULANO}, com templates ASP!
    </body>
</html>
```

Repare na variável FULANO, entre chaves. Ela vai ter seu valor atribuído no código ASP.

Variáveis só podem contêr em seu nome: letras, números e underscore (_). O uso de maiúsculas é apenas uma convenção, pra facilitar a identificação quando olhamos o código HTML. Mas, obrigatoriamente tem que estar entre chaves, sem nenhum espaço em branco.

Então, como ficaria o código ASP que atribui valor a ela? Vamos a ele:

```asp
<!--#include file="Template.class.asp" -->
<%
    Dim tpl
    Set tpl = New Template
        tpl.addTemplate "hello.html"
        tpl.setVariable "FULANO", "Plécyo"
        tpl.show()
    Set tpl = Nothing
%>
```

Execute então novamente o script, e você verá que o código final gerado no navegador será:

```html
<!DOCTYPE html>
<html lang="pt-br">
    <head>
        <meta charset="utf-8">
        <title>Olá Mundo</title>
    </head>
    <body>
        Olá Plécyo, com templates ASP!
    </body>
</html>
```

Variáveis de template que não tiverem um valor atribuído, serão limpas do código final gerado.

Outra coisa sobre variáveis: você pode repetir as variáveis de template (ou seja, usar a mesma variável em vários lugares). Mas, óbvio, todas mostrarão o mesmo valor.

Repare que usando as variáveis de template, você pode continuar editando o arquivo HTML em seu editor favorito: as variáveis de template serão reconhecidas como um texto qualquer, e o arquivo HTML não ficará poluído de código ASP. O contrário também é verdade: seu arquivo ASP ficará limpo de código HTML.


## Checando se Variáveis Existem

Caso você queira atribuir valor pra uma variável de template, mas não tem certeza se a variável existe, você pode usar o método exists() para fazer essa checagem.

Como é de se esperar, ele retorna true caso a variável exista. Caso não, retorna false:

```asp
<!--#include file="Template.class.asp" -->
<%
    Dim tpl
    Set tpl = New Template
        tpl.addTemplate "hello.html"
        
        If tpl.exists("FULANO") Then
            tpl.setVariable "FULANO", "Plécyo"
        End If

        tpl.show()
    Set tpl = Nothing
%>
```


## Variáveis com Modificadores

É possível, dentro do arquivo HTML, chamarmos algumas funções do ASP, desde que elas atendam as duas condições:

- A função tem que sempre retornar um valor;
- A função deve ter sempre como primeiro parâmetro uma string;

Vamos supor o seguinte arquivo ASP, que atribui as variáveis de template NOME e VALOR:

```asp
<!--#include file="Template.class.asp" -->
<%
    Dim tpl
    Set tpl = New Template
        tpl.addTemplate "hello.html"
        tpl.setVariable "NOME", "Fulano Ciclano da Silva"
        tpl.setVariable "VALOR", 100
        tpl.show()
    Set tpl = Nothing
%>
```

E o seguinte HTML, já fazendo uso dos modificadores:

```html
<!DOCTYPE html>
<html lang="pt-br">
    <head>
        <meta charset="utf-8">
        <title>Olá Mundo - Modificadores</title>
    </head>
    <body>
        <div>Nome: {NOME|replace:"Fulano":"Plécyo"}</div>
        <div>Valor: {VALOR|FormatCurrency:2}</div>
    </body>
</html>
```

Explicando: a linha `{NOME|replace:"Fulano":"Plécyo"}` faz a mesma coisa que a função do ASP [`Replace`](http://www.w3schools.com/asp/func_replace.asp) (substitui um texto por outro).

Já no segundo exemplo, usamos a função do ASP, [`FormatCurrency`](http://www.w3schools.com/asp/func_formatcurrency.asp). Neste caso, estamos usando ela para formatar como moeda com duas casas decimais.


## Blocos

Esse é o segundo e último conceito que você precisa saber sobre essa biblioteca de templates: os blocos.

Imagine que você gostaria de listar o total de produtos cadastrados em um banco de dados. Se não houver nenhum produto, você irá exibir uma aviso de que nenhum produto foi encontrado.

Vamos utilizar então, dois blocos: um que mostra a quantidade total; e outro que avisa que não existem produtos cadastrados, caso realmente o banco esteja vazio. O código HTML para isso é:

```html
<!DOCTYPE html>
<html lang="pt-br">
    <head>
        <meta charset="utf-8">
        <title>Olá Mundo - Blocos</title>
    </head>
    <body>
        <p>Quantidade de produtos cadastrados no sistema:</p>

        <!-- BEGIN BLOCK_QUANTIDADE -->
            <div class="destaque">
                Existem {QUANTIDADE} produtos cadastrados.
            </div>
        <!-- END BLOCK_QUANTIDADE -->

        <!-- BEGIN BLOCK_VAZIO -->
            <div class="vazio">
                Não existe nenhum produto cadastrado.
            </div>
        <!-- END BLOCK_VAZIO -->

    </body>
</html>
```

Repare que o início e final do bloco são identificados por um comentário HTML, com a palavra BEGIN (para identificar início) ou END (para identificar fim) e o nome do bloco em seguida.

As palavras BEGIN e END sempre devem ser maiúsculas. O nome do bloco deve conter somente letras, números ou underscore.

E então, no lado ASP, vamos checar se os produtos existem. Caso sim, mostraremos o bloco BLOCK_QUANTIDADE. Caso não, vamos exibir o bloco BLOCK_VAZIO.

```asp
<!--#include file="Template.class.asp" -->
<%
    Dim tpl
    Set tpl = New Template
        tpl.addTemplate "hello.html"
        
        ' Vamos supor que esta quantidade veio do banco de dados
        quantidade = 5
        
        ' Se existem produtos cadastrados, vamos exibir a quantidade
        if quantidade > 0 then
            tpl.setVariable "QUANTIDADE", quantidade
            tpl.block "BLOCK_QUANTIDADE"
        else ' Caso não exista nenhum produto, exibimos a mensagem de vazio
            tpl.block "BLOCK_VAZIO"
        end if
        
        tpl.show()
    Set tpl = Nothing
%>
```

Como você pode reparar, blocos podem conter variáveis de template. E blocos só são exibidos se no código ASP pedirmos isso, através do método block(). Caso contrário, o bloco não é exibido no conteúdo final gerado.

Outro detalhe importante: ao contrário das variáveis de template, cada bloco deve ser único, ou seja, não podemos usar o mesmo nome para vários blocos.

Repare que mesmo com o uso de blocos, podemos continuar editando o arquivo HTML em qualquer editor HTML: os comentários que indicam início e fim de bloco não irão interferir em nada.

Agora vamos a outro exemplo usando blocos: imagine que você precisa mostrar os dados dos produtos que existem no seu cadastro. Vamos então, usando blocos, montar o HTML para isso:

```html
<!DOCTYPE html>
<html lang="pt-br">
    <head>
        <meta charset="utf-8">
        <title>Olá Mundo - Blocos</title>
    </head>
    <body>
        <p>Produtos cadastrados no sistema:</p>
        <table border="1">
            <thead>
                <tr>
                    <th>Nome</th>
                    <th>Quantidade</th>
                </tr>
            </thead>
            <tbody>
                <!-- BEGIN BLOCK_PRODUTO -->
                    <tr>
                        <td> {NOME} </td>
                        <td> {QUANTIDADE} </td>
                    </tr>
                <!-- END BLOCK_PRODUTO -->
            </tbody>
        </table>
    </body>
</html>
```

Repare que temos apenas uma linha de tabela HTML para os dados dos produtos, dentro de um bloco. Vamos então atribuir valor a estas variáveis, e ir duplicando o conteúdo do bloco conforme listamos os produtos:

```asp
<!--#include file="Template.class.asp" -->
<%
    Dim tpl
    Set tpl = New Template
        tpl.addTemplate "hello.html"

        ' Simulando produtos cadastrados no banco de dados
        produtos = Array( _
            Array("Sabão em Pó", 15), _
            Array("Escova de Dente", 53), _
            Array("Creme Dental", 37) _
        )

        ' Listando os produtos
        For Each p In produtos
            tpl.setVariable "NOME", p(0)
            tpl.setVariable "QUANTIDADE", p(1)

            tpl.block "BLOCK_PRODUTO"
        Next

        tpl.show()
    Set tpl = Nothing
%>
```

O comportamento padrão do método block() é manter o conteúdo anterior do bloco, somado (ou melhor, concatenado) ao novo conteúdo que acabamos de atribuir.

No exemplo acima, os dados dos produtos vieram do array "produtos". Caso estes dados estivessem armazenados em um banco de dados, então bastaríamos fazer como no exemplo abaixo:

```asp
<!--#include file="Template.class.asp" -->
<%
    Dim tpl
    Set tpl = New Template
        tpl.addTemplate "hello.html"
        
        ' ... Conectar ao banco, selecionar database, etc
        
        ' Produtos do database
        Set result = conexao.Execute("SELECT nome, quantidade FROM produtos")
		
        ' Listando os produtos
        While Not result.EOF Then
            tpl.setVariable "NOME", result("nome")
            tpl.setVariable "QUANTIDADE", result("quantidade")

            tpl.block "BLOCK_PRODUTO"
        Wend
		
        tpl.show()
    Set tpl = Nothing
%>
```


## Blocos aninhados

Vamos agora então juntar os 2 exemplos de uso de blocos que vimos: queremos mostrar os dados dos produtos em um bloco, mas caso não existam produtos cadastrados, exibiremos uma mensagem com um aviso. Vamos fazer isso agora usando blocos aninhados, ou seja, blocos dentro de outros blocos:

```html
<!DOCTYPE html>
<html lang="pt-br">
    <head>
        <meta charset="utf-8">
        <title>Olá Mundo - Blocos</title>
    </head>
    <body>
        <p>Produtos cadastrados no sistema:</p>

        <!-- BEGIN BLOCK_PRODUTOS -->
            <table border="1">
                <thead>
                    <tr>
                        <th>Nome</th>
                        <th>Quantidade</th>
                    </tr>
                </thead>
                <tbody>
                    <!-- BEGIN BLOCK_DADOS -->
                        <tr>
                            <td> {NOME} </td>
                            <td> {QUANTIDADE} </td>
                        </tr>
                    <!-- END BLOCK_DADOS -->
                </tbody>
            </table>
        <!-- END BLOCK_PRODUTO -->

        <!-- BEGIN BLOCK_VAZIO -->
            <div class="vazio">
                Nenhum registro encontrado.
            </div>
        <!-- END BLOCK_VAZIO -->

    </body>
</html>
```

E então, caso existam produtos, nós exibimos o bloco PRODUTOS. Caso contrário, exibimos o bloco VAZIO.
Um detalhe muito importante em relação a sua antecessora: se um bloco aninhado é exibido, todos os blocos pais serão automaticamente exibidos.

```asp
<!--#include file="Template.class.asp" -->
<%
    Dim tpl
    Set tpl = New Template
        tpl.addTemplate "hello.html"
        
        ' Produtos cadastrados
        produtos = Array( _
            Array("Sabão em Pó", 15), _
            Array("Escova de Dente", 53), _
            Array("Creme Dental", 37) _
        )
		
        ' Listando os produtos
        For Each p In produtos
            tpl.setVariable "NOME", p(0)
            tpl.setVariable "QUANTIDADE", p(1)

            tpl.block "BLOCK_DADOS"
        Next
		
        ' Se não existem produtos, mostramos o bloco com o aviso de nenhum cadastrado
        If Not IsArray(produtos) Then
            tpl.block "BLOCK_VAZIO"
        End If

        tpl.show()
    Set tpl = Nothing
%>
```

Ou seja, se existem produtos, e por consequência o BLOCK_DADOS foi exibido, o BLOCK_PRODUTOS automaticamente será.

Mais continue lendo, vamos conseguir fazer mais coisas automaticamente, e com menos código, usando os Blocos FINALLY.


## Blocos FINALLY

No exemplo anterior, usamos o BLOCK_PRODUTOS para exibir os produtos, e caso não existam produtos, usamos o BLOCK_VAZIO para exibir uma mensagem amigável de que não existem produtos cadastrados.

Podemos fazer isso de forma mais automática: usandos os blocos FINALLY.

Veja como ficaria o arquivo HTML neste caso:

```html
<!DOCTYPE html>
<html lang="pt-br">
    <head>
        <meta charset="utf-8">
        <title>Olá Mundo - Blocos</title>
    </head>
    <body>
        <p>Produtos cadastrados no sistema:</p>

        <!-- BEGIN BLOCK_PRODUTOS -->
            <table border="1">
                <thead>
                    <tr>
                        <th>Nome</th>
                        <th>Quantidade</th>
                    </tr>
                </thead>
                <tbody>
                    <!-- BEGIN BLOCK_DADOS -->
                        <tr>
                            <td> {NOME} </td>
                            <td> {QUANTIDADE} </td>
                        </tr>
                    <!-- END BLOCK_DADOS -->
                </tbody>
            </table>
        <!-- END BLOCK_PRODUTO -->

            <div class="vazio">
                Nenhum registro encontrado.
            </div>
        <!-- FINALLY BLOCK_PRODUTO -->

    </body>
</html>
```

E o arquivo ASP? Bem, ele vai ficar mais simples ainda:

```asp
<!--#include file="Template.class.asp" -->
<%
    Dim tpl
    Set tpl = New Template
        tpl.addTemplate "hello.html"
        
        ' Produtos cadastrados
        produtos = Array( _
            Array("Sabão em Pó", 15), _
            Array("Escova de Dente", 53), _
            Array("Creme Dental", 37) _
        )
		
        ' Listando os produtos
        If IsArray(produtos) Then
            For Each p In produtos
                tpl.setVariable "NOME", p(0)
                tpl.setVariable "QUANTIDADE", p(1)
                tpl.block "BLOCK_DADOS"
            Next
        End If

        tpl.show()
    Set tpl = Nothing
%>
```

Primeiro detalhe importante: o bloco FINALLY nunca precisa ser invocado no arquivo ASP. Caso ele exista no HTML, ele sempre será chamado se o bloco relacionado não for exibido.

Ou seja, se não houverem produtos, o bloco FINALLY exibirá automaticamente o aviso de que não exitem registros encontrados. Prático, né?

E segundo detalhe: no arquivo ASP, o bloco BLOCK_PRODUTOS nem foi chamado. Mas como o bloco mais interno, BLOCK_DADOS, foi chamado, o bloco pai automaticamente será mostrado.


## Blocos com HTML Select

Uma das dúvidas mais comuns é: como usar a classe Template com o elemento Select do HTML? Ou melhor: como fazer um elemento Option ficar selecionado, usando Template?

Vamos então montar nossa página HTML com o elemento Select e os devidos Options, representando cidades de uma lista:

```html
<!DOCTYPE html>
<html lang="pt-br">
    <head>
        <meta charset="utf-8">
        <title>Olá Mundo - Blocos</title>
    </head>
    <body>
        <select name="cidades">
            <!-- BEGIN BLOCK_OPTION -->
                <option value="{VALUE}" {SELECTED}>
                    {TEXT}
                </option>
            <!-- END BLOCK_OPTION -->
        </select>
    </body>
</html>
```

Agora vamos ao respectivo arquivo ASP:

``` asp
<!--#include file="Template.class.asp" -->
<%
    Dim tpl
    Set tpl = New Template
        tpl.addTemplate "hello.html"
        
        ' Array de cidades
        cidades = Array( _
            Array(0, "Cidade 0"), _
            Array(1, "Cidade 1"), _
            Array(2, "Cidade 2") _
        )

        ' Valor selecionado
        atual = 1
        
        ' Listando as cidades
        For Each c In cidades
            tpl.setVariable "VALUE", c(0)
            tpl.setVariable "TEXT", c(1)

            If atual = c(0) Then
                tpl.setVariable "SELECTED", "selected"
            Else
                tpl.clear "SELECTED"
            End If
			
            tpl.block "BLOCK_OPTION"
        Next
		
        tpl.show()
    Set tpl = Nothing
%>
```

Como resultado, o navegador exibirá o seguinte código:

```html
<!DOCTYPE html>
<html lang="pt-br">
    <head>
        <meta charset="utf-8">
        <title>Olá Mundo - Blocos</title>
    </head>
    <body>
        <select name="cidades">
            <option value="0" >Cidade 0</option>
            <option value="1" selected>Cidade 1</option>
            <option value="2" >Cidade 2</option>
        </select>
    </body>
</html>
```

Reparou que no arquivo ASP chamamos o método clear? Se não chamarmos este método (que limpa o valor de uma variável), todas as opções (Options) ficariam com a propriedade "selected" (obviamente, efeito não desejado):

```html
<!DOCTYPE html>
<html lang="pt-br">
    <head>
        <meta charset="utf-8">
        <title>Olá Mundo - Blocos</title>
    </head>
    <body>
        <select name="cidades">
            <option value="0" selected>Cidade 0</option>
            <option value="1" selected>Cidade 1</option>
            <option value="2" selected>Cidade 2</option>
        </select>
    </body>
</html>
```


## Usando vários arquivos HTML

Um uso bastante comum de templates é usarmos um arquivo HTML que contenha a estrutura básica do nosso site: cabeçalho, rodapé, menus, etc. E outro arquivo com o conteúdo da página que desejamos mostrar, ou seja, o "miolo". Dessa forma, não precisamos repetir em todos os arquivos HTML os elementos comuns (cabeçalho, rodapé, etc), e as páginas HTML que terão o conteúdo (o "miolo") ficarão mais limpas, menores e mais fáceis de serem mantidas.

Como fazer isso com templates? Em primeiro lugar, vamos criar nosso arquivo "base" HTML, o arquivo base.html:

```html
<!DOCTYPE html>
<html lang="pt-br">
    <head>
        <meta charset="utf-8">
        <title>Olá Mundo</title>
    </head>
    <body>
        <div>{FULANO}, seja bem vindo!</div>
	
        <div>{CONTEUDO}</div>
	
        <div>Deseja maiores informações? Clique <a href="#">aqui</a> para saber</div>
    </body>
</html>
```

Agora, vamos criar o arquivo que contém o "miolo" de nossa página HTML, o arquivo miolo.html:

```html
    <p>Produtos cadastrados no sistema:</p>

    <!-- BEGIN BLOCK_PRODUTOS -->
        <table border="1">
            <thead>
                <tr>
                    <th>Nome</th>
                    <th>Quantidade</th>
                </tr>
            </thead>
            <tbody>
                <!-- BEGIN BLOCK_DADOS -->
                    <tr>
                        <td> {NOME} </td>
                        <td> {QUANTIDADE} </td>
                    </tr>
                <!-- END BLOCK_DADOS -->
            </tbody>
        </table>
    <!-- END BLOCK_PRODUTO -->

        <div class="vazio">
            Nenhum registro encontrado.
        </div>
    <!-- FINALLY BLOCK_PRODUTO -->
```

No arquivo ASP então, usamos o método addFile(), onde informamos duas coisas: em qual variável do template o conteúdo do novo arquivo será jogado, e qual o caminho desse arquivo. Depois disso, basta usar as variáveis e blocos normalmente, independente de qual arquivo HTML eles estejam:

```asp
<!--#include file="Template.class.asp" -->
<%
    Dim tpl
    Set tpl = New Template
        tpl.addTemplate "base.html"
        
        ' Adicionando mais um arquivo HTML
        tpl.addFile "CONTEUDO", "miolo.html"
		
        tpl.setVariable "FULANO", "Plécyo"
        
        ' Produtos cadastrados
        produtos = Array( _
            Array("Sabão em Pó", 15), _
            Array("Escova de Dente", 53), _
            Array("Creme Dental", 37) _
        )
        
        ' Listando os produtos
        For Each p In produtos
            tpl.setVariable "NOME", p(0)
            tpl.setVariable "QUANTIDADE", p(1)

            tpl.block "BLOCK_DADOS"
        Next
		
        tpl.show()
    Set tpl = Nothing
%>
```


## Guardando o conteúdo do template

Até agora exibimos o conteúdo gerado pelo template na tela, através do método show(). Mas, e quisermos fazer outro uso para esse conteúdo, como salvá-lo em arquivo ou outra coisa do tipo? Basta usarmos o método parse(), que gera o conteúdo final e o retorna:

```asp
<!--#include file="Template.class.asp" -->
<%
    Dim tpl
    Set tpl = New Template
        tpl.addTemplate "base.html"
        
        ' Adicionando mais um arquivo HTML
        tpl.addFile "CONTEUDO", "miolo.html"
		
        tpl.setVariable "FULANO", "Plécyo"
        
        ' Produtos cadastrados
        produtos = Array( _
            Array("Sabão em Pó", 15), _
            Array("Escova de Dente", 53), _
            Array("Creme Dental", 37) _
        )
        
        ' Listando os produtos
        For Each p In produtos
            tpl.setVariable "NOME", p(0)
            tpl.setVariable "QUANTIDADE", p(1)

            tpl.block "BLOCK_DADOS"
        Next

    ' Pega o conteúdo final do template
    html = tpl.parse()
    
    Set tpl = Nothing

    ' Cria um objeto manipulação de arquivos
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Seleciona a pasta TEMP do Windows
    Dim temp_folder
    Set temp_folder = fso.GetSpecialFolder(2)

    ' Salva o arquivo no diretório
    Dim temp_file
    Set temp_file = temp_folder.CreateTextFile(fso.GetTempName)
        temp_file.WriteLine(html)
        temp_file.Close
    Set temp_file = Nothing
%>
```


## Usando Objetos

A classe Template suporta a atribuição de objetos para as variáveis de template.

Isso faz com que o código nos arquivos ASP fique bastante reduzido (claro, desde que você use objetos), e devido a isso, há uma melhora (quase imperceptível) em desempenho.

Para que os exemplos fiquem mais claros, vamos trabalhar com uma suposta página que exibe detalhes de um produto. Os produtos possuem como atributos: `id` e `name`.

Primeiro, vamos a um exemplo de classe Produtos:

```asp
<%
Class Product
	
    Private var_id
    Private var_name
	
    Public Property Get getId()
        getId = var_id
    End Property

    Public Property Let setId(p_id)
        var_id = p_id
    End Property

    Public Property Get getName()
        getName = var_name
    End Property

    Public Property Let setName(p_name)
        var_name = p_name
    End Property

End Class
%>
```

Vamos então modificar o arquivo ASP para carregar um produto, e usar o suporte a objetos de Template:

```asp
<!--#include file="Template.class.asp" -->
<%
    Dim tpl
    Set tpl = New Template
        tpl.addTemplate "base.html"
        
        ' Adicionando mais um arquivo HTML
        tpl.addFile "CONTEUDO", "miolo.html"

        tpl.setVariable "FULANO", "Plécyo"
        
        Set objProduct_1 = New Product
            objProduct_1.setId   = 0
            objProduct_1.setName = "Sabão em Pó"
		
        Set objProduct_2 = New Product
            objProduct_2.setId   = 1
            objProduct_2.setName = "Escova de Dente"
		
        Set objProduct_3 = New Product
            objProduct_3.setId   = 2
            objProduct_3.setName = "Creme Dental"
		
        ' Produtos cadastrados
        produtos = Array( objProduct_1, objProduct_2, objProduct_3 )
        
        ' Listando os produtos
        For Each obj In produtos
            tpl.setVariable "PRODUTO", obj

            tpl.block "BLOCK_DADOS"
        Next
		
        tpl.show()
    Set tpl = Nothing
%>
```

O arquivo HTML também deve ser modificado pra exibir as propriedades de Produto:

```html
    <p>Produtos cadastrados no sistema:</p>

    <!-- BEGIN BLOCK_PRODUTOS -->
        <table border="1">
            <thead>
                <tr>
                    <th>Id</th>
                    <th>Name</th>
                </tr>
            </thead>
            <tbody>
                <!-- BEGIN BLOCK_DADOS -->
                    <tr>
                        <td> {PRODUTO->ID} </td>
                        <td> {PRODUTO->NAME} </td>
                    </tr>
                <!-- END BLOCK_DADOS -->
            </tbody>
        </table>
    <!-- END BLOCK_PRODUTO -->

        <div class="vazio">
            Nenhum registro encontrado.
        </div>
    <!-- FINALLY BLOCK_PRODUTO -->
```

A instrução PRODUTO->NAME chamará o método Produto.getName(), caso ele exista. Se não existir esse método na classe, um erro será disparado.

Isso vale para qualquer atributo que tentarmos chamar no HTML: será traduzido para `meuObjeto.getAtributo()`.
Se o nome do método ASP for composto, como por exemplo `meuObjeto.getExpirationDate()`, basta usar underscore `_` no HTML como separador dos nomes: no caso do exemplo, ficaria `PRODUTO->EXPIRATION_DATE`.


## Comentários

A exemplo das linguagens de programação, a classe Template suporta comentários no HTML. Comentários são úteis para várias coisas, entre elas, identificar o autor do HTML, versão, incluir licenças, etc.

Diferentemente dos comentários HTML, que são exibidos no código fonte da página, os comentários da classe Template são extraídos do HTML final. Na verdade os comentários de Template são extraídos antes mesmo de qualquer processamento, e tudo que estiver entre os comentários será ignorado.

Os comentários ficam entre as tags `<!---` e `--->`.
Repare que usamos 3 tracinhos, ao invés de 2 (que identificam comentários HTML).
A razão é simples: permitir diferenciarmos entre um e outro, e permitir que os editores continuem reconhecendo o conteúdo entre `<!---` e `--->` como comentários.

Veja o exemplo abaixo:

```
    <!---
        Listagem de produtos.
	
        @author Plécyo
        @version 1.0
    --->
```
```html
    <p>Produtos cadastrados no sistema:</p>

    <!-- BEGIN BLOCK_PRODUTOS -->
        <table border="1">
            <thead>
                <tr>
                    <th>Id</th>
                    <th>Name</th>
                </tr>
            </thead>
            <tbody>
                <!-- BEGIN BLOCK_DADOS -->
                    <tr>
                        <td> {PRODUTO->ID} </td>
                        <td> {PRODUTO->NAME} </td>
                    </tr>
                <!-- END BLOCK_DADOS -->
            </tbody>
        </table>
    <!-- END BLOCK_PRODUTO -->

        <div class="vazio">
            Nenhum registro encontrado.
        </div>
    <!-- FINALLY BLOCK_PRODUTO -->
```


## Criando XML, CSV e outros

O uso mais comum de templates é com arquivos HTML. Mas como essa biblioteca é direcionada ao uso de qualquer tipo de arquivo de texto, podemos usá-la com vários outros formatos de arquivo, como XMLs e arquivos CSVs.

Como fazer isso? Mais simples impossível: não muda nada, basta apenas ao invés de indicar um arquivo HTML para o Template, indicar qualquer outro arquivo de texto. E usar variáveis e blocos nele conforme já vimos, exibindo o conteúdo na tela, ou salvando em arquivos.


## Escapando Variáveis

Vamos supor que por algum motivo você precise manter uma variável de template no resultado final de seu HTML. Como por exemplo: você está escrevendo um sistema que gera os templates automaticamente pra você.

Para isso, vamos supor que você tenha o HTML abaixo:

```html
<!DOCTYPE html>
<html lang="pt-br">
    <head>
        <meta charset="utf-8">
        <title>Olá Mundo</title>
    </head>
    <body>
        {CONTEUDO}
    </body>
</html>
```

E você precisa que `{CONTEUDO}` não seja substituído (ou removido), mas que permaneça no HTML final.

Para isso, faça o *escape* incluindo `{_}` dentro da variável:

```html
<!DOCTYPE html>
<html lang="pt-br">
    <head>
        <meta charset="utf-8">
        <title>Olá Mundo</title>
    </head>
    <body>
        {{_}CONTEUDO}
    </body>
</html>
```

E pronto: no HTML final `{CONTEUDO}` ainda estará presente.


## Conclusão

O uso de mecanismos de Template é um grande avanço no desenvolvimento de aplicações web, pois nos permite manter a estrutura visual de nosso aplicativo separado da programação ASP.

Eu tentei incluir neste tutorial todos os tópicos que cobrem o uso de templates. Se você tiver problemas, procure primeiro em todos os tópicos aqui. Lembre-se que este trabalho é voluntário, e que eu gastei muito tempo escrevendo este tutorial, além do tempo gasto nesta biblioteca. Portanto, antes de me enviar um email com algum problema, tente resolvê-lo sozinho: grande parte do aprendizado está nisso. Se você não conseguir, vá em frente, e fale comigo.
