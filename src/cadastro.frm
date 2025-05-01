VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cadastro 
   Caption         =   "Nova Programação"
   ClientHeight    =   3315
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13170
   OleObjectBlob   =   "cadastro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public tamanhoPadrao As Double
Public espacamentoPadrao As Double
Public defaultFormHeight As Double
Public itensPorPagina As Integer
Public contagem As Range
Private Function confirmaImportacao() As Boolean

    confirmaImportacao = True
    
    'Loop em todos controle do form
    For Each Obj In Me.Controls
        'Se for algum dos campos de texto
        If InStr(1, Obj.Name, "txtDescricao") >= 1 Or InStr(1, Obj.Name, "txtParticipantes") >= 1 Or InStr(1, Obj.Name, "txtInstrumentos") >= 1 Or InStr(1, Obj.Name, "txtProjecao") >= 1 Or InStr(1, Obj.Name, "listBoxTipoProjecao") >= 1 Then
            
            If Obj.Value <> "" Then
                'Confirmação
                resposta = MsgBox("O formulário atual possui campos preenchidos!" & vbNewLine & "Essas informações serão perdidas!" & vbNewLine & "Deseja continuar?", vbYesNo Or vbExclamation, "CONFIRMAÇÃO")
    
                If resposta = vbYes Then
                    confirmaImportacao = True
                    Exit Function
                Else
                    confirmaImportacao = False
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Private Function criaNovaPlanilha()
    Workbooks.Add 'Cria uma planilha
End Function
Private Function verificaPreenchimento() As Boolean
    Dim preenchido As Boolean
    Dim firstEmpty As Boolean
    
    firstEmpty = False
    
    'Se não foi informado um titulo
    If Me.txtTitulo.Value = "" Then
        verificaPreenchimento = False
        MsgBox "Insira um título para a programação!", vbInformation, "Título Necessário!"
        Me.txtTitulo.SetFocus
        Exit Function
    End If
    
    For Each Obj In Me.Controls 'Loop em cada controle do formulario
        
        'Caso seja textBox e não seja o titulo e a data
        If InStr(1, Obj.Name, "txt") >= 1 And InStr(1, Obj.Name, "Titulo") <= 0 And InStr(1, Obj.Name, "Data") <= 0 Then
            If Obj.Value <> "" Then
                verificaPreenchimento = True
                preenchido = True
                Exit Function
            Else
                If firstEmpty = False Then
                    Obj.SetFocus
                    firstEmpty = True
                End If
            End If
        End If
    Next
    
    'Se não foi preenchido os campos necessarios exibe uma mensagem
    If preenchido = False Then
        MsgBox "Preencha ao menos um campo!", vbInformation, "Informações Insuficientes!"
    End If
End Function
Private Function paginacao(ByVal pagAtual As Integer)
    Dim numItem As Integer
    
    numItem = 0
    
    'Iteração em cada objeto
    For Each Obj In Me.Controls
        'Se os dois ultimos caracteres do obj são numéricos
        If IsNumeric(Right(Obj.Name, 2)) = True Then
            numItem = Right(Obj.Name, 2)
        'Se o ultimo caractere do obj é numérico
        ElseIf IsNumeric(Right(Obj.Name, 1)) = True Then
            numItem = Right(Obj.Name, 1)
        End If
        
        'Se for algum dos campos criados dinamicamente
        If InStr(1, Obj.Name, "lbContador") >= 1 Or InStr(1, Obj.Name, "txtDescricao") >= 1 Or _
        InStr(1, Obj.Name, "txtParticipantes") >= 1 Or InStr(1, Obj.Name, "txtInstrumentos") >= 1 Or _
        InStr(1, Obj.Name, "txtProjecao") >= 1 Or InStr(1, Obj.Name, "listBoxTipoProjecao") >= 1 Then
            'Se o objeto está entre os que devem ser renderizados conforme a página atual
            If numItem > itensPorPagina * pagAtual - itensPorPagina _
            And numItem <= itensPorPagina * pagAtual Then
                Obj.Visible = True 'Torna o obj visivel
            Else
                Obj.Visible = False 'Torna o obj invisivel
            End If
        End If
    Next
    
End Function
Private Sub btnGerar_Click()
    Dim aba As Worksheet
    Dim paginaAtual As Range
    Dim newLabel As Control
    Dim newTxtDesc As Control
    Dim newTxtParticipants As Control
    Dim newTxtInstrumentos As Control
    Dim newTxtProjecao As Control
    Dim newListBox As Control

    Set paginaAtual = ThisWorkbook.Worksheets("controle").Range("B2") 'Pega a pagina atual
    
    'Se ja tem 8 itens criados
    If contagem.Value >= itensPorPagina * paginaAtual.Value Then
        paginaAtual.Value = paginaAtual.Value + 1
        Me.lNumPagina.Caption = Str(paginaAtual.Value)
        Call paginacao(paginaAtual.Value)  'Chama a função que deixa os objetos com a devida visibilidade
    End If
    
    'redimensiona e posiciona os objetos
    Call redimensiona
    
    contagem.Value = contagem.Value + 1 'Soma o contador de itens criados
    
    Set newLabel = Me.Controls.Add("Forms.Label.1", "lbContador" & contagem) 'Cria o Label da contagem'
    Set newTxtDesc = Me.Controls.Add("Forms.TextBox.1", "txtDescricao" & contagem) 'Cria o txtBox da descrição'
    Set newTxtParticipants = Me.Controls.Add("Forms.TextBox.1", "txtParticipantes" & contagem) 'Cria o txtBox dos participantes'
    Set newTxtInstrumentos = Me.Controls.Add("Forms.TextBox.1", "txtInstrumentos" & contagem) 'Cria o txtBox dos instrumentos'
    Set newTxtProjecao = Me.Controls.Add("Forms.TextBox.1", "txtProjecao" & contagem) 'Cria o txtBox da projecao'
    Set newListBox = Me.Controls.Add("Forms.ListBox.1", "listBoxTipoProjecao" & contagem) 'Cria o ListBox do tipo da projeção'
    
    Set aba = ThisWorkbook.Worksheets("controle") 'define a aba onde esta os valores
    uLin = aba.Cells(1, 3).End(xlDown).Row 'define a ultima linha preenchida'
   
   'Configura o label
   With newLabel
       .Tag = "lbContador" & Str(contagem)
       .Caption = contagem
       .Font.bold = Me.lbContador1.Font.bold
       .Font.Size = Me.lbContador1.Font.Size
       .Top = Me.lbContador1.Top + tamanhoPadrao * (contagem - 1 - (itensPorPagina * (paginaAtual - 1)))
       .Left = Me.lbContador1.Left
       .Width = Me.lbContador1.Width
       .Height = Me.lbContador1.Height
       .TextAlign = Me.lbContador1.TextAlign
   End With
    
    'Configura o textbox Descricao
    With newTxtDesc
        .Tag = "txtDescricao" & contagem
        .Top = Me.txtDescricao1.Top + tamanhoPadrao * (contagem - 1 - (itensPorPagina * (paginaAtual - 1)))
        .Left = Me.txtDescricao1.Left
        .Width = Me.txtDescricao1.Width
        .Height = Me.txtDescricao1.Height
        .TextAlign = Me.txtDescricao1.TextAlign
    End With
    
    'Configura o textbox Participantes
    With newTxtParticipants
        .Tag = "txtParticipantes" & contagem
        .Top = Me.txtParticipantes1.Top + tamanhoPadrao * (contagem - 1 - (itensPorPagina * (paginaAtual - 1)))
        .Left = Me.txtParticipantes1.Left
        .Width = Me.txtParticipantes1.Width
        .Height = Me.txtParticipantes1.Height
        .TextAlign = Me.txtParticipantes1.TextAlign
    End With
    
    'Configura o textbox Instrumentos
    With newTxtInstrumentos
        .Tag = "txtInstrumentos" & contagem
        .Top = Me.txtInstrumentos1.Top + tamanhoPadrao * (contagem - 1 - (itensPorPagina * (paginaAtual - 1)))
        .Left = Me.txtInstrumentos1.Left
        .Width = Me.txtInstrumentos1.Width
        .Height = Me.txtInstrumentos1.Height
        .TextAlign = Me.txtInstrumentos1.TextAlign
    End With
    
    'Configura o textbox Projecao
    With newTxtProjecao
        .Tag = "txtProjecao" & contagem
        .Top = Me.txtProjecao1.Top + tamanhoPadrao * (contagem - 1 - (itensPorPagina * (paginaAtual - 1)))
        .Left = Me.txtProjecao1.Left
        .Width = Me.txtProjecao1.Width
        .Height = Me.txtProjecao1.Height
        .TextAlign = Me.txtProjecao1.TextAlign
    End With
    
    'Configura o listbox Tipo de Projeção
    With newListBox
        .Tag = "listBoxTipoProjecao" & contagem
        .Width = Me.listBoxTipoProjecao1.Width
        .Height = Me.listBoxTipoProjecao1.Height
        .Top = Me.listBoxTipoProjecao1.Top + tamanhoPadrao * (contagem - 1 - (itensPorPagina * (paginaAtual - 1)))
        .Left = Me.listBoxTipoProjecao1.Left
        
        'faz loop nas celulas com as opções e coloca no ListBox'
        For i = 2 To uLin
            .AddItem aba.Cells(i, 3)
        Next i
    End With

End Sub
Private Sub redimensiona()
    Dim itensNaPagina As Integer
    Dim paginasTotal As Integer
    Dim paginaAtual As Range
    Set paginaAtual = ThisWorkbook.Worksheets("controle").Range("B2")
    
    paginasTotal = WorksheetFunction.RoundUp((contagem / itensPorPagina) + 0.1, 0) 'Quantidade total de paginas
    
    If paginaAtual >= paginasTotal Then 'Se for a ultima pagina
        If paginasTotal > 1 Then 'Se haver mais de uma pagina
            itensNaPagina = contagem - (itensPorPagina * (paginasTotal - 1)) '
        Else 'Se for a primeira pagina
            itensNaPagina = contagem
        End If
    Else 'Se não for a ultima pagina
        itensNaPagina = 7 '+ 1
    End If
    
    Me.Height = defaultFormHeight + (tamanhoPadrao * itensNaPagina) 'Ajusta o tamanho do form'
    
    Me.fSave.Top = Me.Height - 85 'Reposiciona o frame dos métodos de salvamento e todos botões
    Me.fAcoes.Top = Me.Height - 85 'Reposiciona o frame das ações e os botões dentro
    Me.fPaginacao.Top = Me.Height - 85 'Reposiciona o frame da navegação de páginas
    
    'Se tem mais de uma pagina
    If paginaAtual.Value > 1 Then
        Me.Width = 705
        Me.btnGerar.Width = 56
        Me.fPaginacao.Visible = True
    End If
    
End Sub

Private Sub btnImportar_Click()
    Dim planilha As Workbook
    Dim cabecalho As Range
    Dim contagemColunas As Integer
    Dim quantItens As Integer
    Dim objNumero As Integer
    Dim confirmado As Boolean
    
    confirmado = confirmaImportacao
    
    If confirmado = False Then
        MsgBox "Importação cancelada!", vbOKOnly Or vbInformation, "ANULADO"
        Exit Sub
    End If
    
    contagemColunas = 0
    
    ' Abre o explorador de arquivos
    caminho = Application.GetOpenFilename("Planilha Eletrônica do Excel (*.xlsx), *.xlsx", , "Selecionar Arquivo", , False)
    
    If caminho <> False Then 'Se selecionou algum arquivo
        
        ' Tratativa de erro
        On Error GoTo ErroAoAbrir:
        
        Set planilha = Workbooks.Open(caminho, ReadOnly:=False, editable:=True)  'Abre a planilha
        Set cabecalho = planilha.Worksheets(1).Range("A3:F3") 'Define o cabeçalho
    
        On Error GoTo 0
        
        For Each celula In cabecalho 'Loop no cabecalho
            If InStr(1, celula.Value, "N°") >= 1 Then 'Se é a coluna N°
                contagemColunas = contagemColunas + 1
            ElseIf InStr(1, celula.Value, "DESCRIÇÃO") >= 1 Then 'Se é a coluna DESCRIÇÃO
                contagemColunas = contagemColunas + 1
            ElseIf InStr(1, celula.Value, "PARTICIPANTES") >= 1 Then 'Se é a coluna PARTICIPANTES
                contagemColunas = contagemColunas + 1
            ElseIf InStr(1, celula.Value, "INSTRUMENTOS") >= 1 Then 'Se é a coluna INSTRUMENTOS
                contagemColunas = contagemColunas + 1
            ElseIf InStr(1, celula.Value, "MÍDIA") >= 1 Then 'Se é a coluna MÍDIAS
                contagemColunas = contagemColunas + 1
            ElseIf InStr(1, celula.Value, "PROJEÇÃO") >= 1 Then 'Se é a coluna PROJEÇÃO
                contagemColunas = contagemColunas + 1
            End If
        Next
        
        If contagemColunas >= 6 Then 'Se todas as colunas foram encontradas na planilha
            quantItens = planilha.ActiveSheet.Range("A3").End(xlDown).Row - 3 'Ultima linha preenchida
            MsgBox quantItens & " Linhas serão importadas!"
            
            'Loop nas linhas da planilha que será importada
            For i = 4 To quantItens + 3
                'Se restam objetos a serem criados então os cria
                If quantItens - contagem.Value > 0 Then
                    'Loop na quantidade restante de itens
                    For j = 1 To quantItens - contagem
                        Call btnGerar_Click
                    Next j
                End If
            
                'Loop para ver se ja existem objetos criados com a respectiva linha da planilha
                For Each Obj In Me.Controls
                
                    'Obtém a sequência do objeto
                    If IsNumeric(Right(Obj.Name, 2)) = True Then 'Se os dois ultimos caracteress do nome é numérico
                        objNumero = Right(Obj.Name, 2)
                    ElseIf IsNumeric(Right(Obj.Name, 1)) = True Then 'Se o ultimo caractere do nome é numerico
                        objNumero = Right(Obj.Name, 1)
                    Else
                        objNumero = Empty 'Se não for um campo gerado em tempo de execução então fica vazio
                    End If
                    
                    'Se é a linha da planilha referente ao objeto
                    If objNumero = i - 3 Then
                        If InStr(1, Obj.Name, "txtDescricao") >= 1 Then 'Se é a coluna DESCRIÇÃO
                            Obj.Value = planilha.ActiveSheet.Cells(i, 2)
                            restante = restante - 1
                        ElseIf InStr(1, Obj.Name, "txtParticipantes") >= 1 Then 'Se é a coluna PARTICIPANTES
                            Obj.Value = planilha.ActiveSheet.Cells(i, 3)
                            restante = restante - 1
                        ElseIf InStr(1, Obj.Name, "txtInstrumentos") >= 1 Then 'Se é a coluna INSTRUMENTOS
                            Obj.Value = planilha.ActiveSheet.Cells(i, 4)
                            restante = restante - 1
                        ElseIf InStr(1, Obj.Name, "txtProjecao") >= 1 Then 'Se é a coluna MÍDIAS
                            Obj.Value = planilha.ActiveSheet.Cells(i, 5)
                            restante = restante - 1
                        ElseIf InStr(1, Obj.Name, "listBoxTipoProjecao") >= 1 Then 'Se é a coluna PROJEÇÃO
                            For j = 1 To Obj.ListCount 'Loop nas opções do listbox até encontrar a correta
                                Obj.ListIndex = j - 1
                                If Obj.Value = planilha.ActiveSheet.Cells(i, 6) Then 'Se for o valor correto
                                    restante = restante - 1
                                    Exit For 'Sai do for
                                Else
                                    If Obj.ListIndex = Obj.ListCount Then 'Se for o ultimo item do listbox
                                        MsgBox "Valor do tipo de projeção não encontrado!", vbCritical Or vbOKOnly, "ERRO"
                                    End If
                                End If
                            Next j
                        End If
                    ElseIf objNumero > quantItens Then 'Se é objeto sobrando
                        If InStr(1, Obj.Name, "listBoxTipoProjecao") >= 1 Then 'se for o ultimo objeto da linha
                            contagem.Value = contagem.Value - 1 'Subtrai a linha excluida
                        End If
                        Me.Controls.Remove Obj.Name 'Remove o objeto
                    End If
                Next
            Next i
            ThisWorkbook.Worksheets("controle").Range("B2").Value = 1 'Define a pagina atual como a primeira
            Call redimensiona
            Call paginacao(1)
            Me.lNumPagina.Caption = 1
            
        Else
            MsgBox "Planilha com formato inválido!" & vbNewLine & "A planilha a ser importada deve ter o mesmo formato das planilhas geradas por esse formulário!", vbInformation Or vbOKOnly, "FALHA AO IMPORTAR"
        End If
        planilha.Close 'Fecha a planilha que foi importada
    End If
Exit Sub
ErroAoAbrir:
    MsgBox "Erro ao tentar abrir a planilha!" & vbNewLine & Err.Description, vbCritical Or vbOKOnly, "ERRO"
    On Error Resume Next
    planilha.Close 'Fecha a planilha que foi importada
End Sub
Private Sub btnPreencher_Click()

        For Each Obj In Me.Controls 'Loop em cada controle do formulario'
            
            'se é um numero na casa da dezena
            If IsNumeric(Right(Obj.Name, 2)) = True Then
                linha = Int(Right(Obj.Name, 2)) 'define a linha da vez'
            'se é um numero < 10
            ElseIf IsNumeric(Right(Obj.Name, 1)) = True Then
                linha = Int(Right(Obj.Name, 1)) 'define a linha da vez'
            End If
            
            'Insere a sequência'
            If InStr(1, Obj.Name, "lbContador") >= 1 Then
                linha = linha + 1
            'Insere as descrições'
            ElseIf InStr(1, Obj.Name, "txtDescricao") >= 1 Then
                Obj.Value = Obj.Value & " Descrição " & linha
                linha = linha + 1
                
            'Insere os participantes'
            ElseIf InStr(1, Obj.Name, "txtParticipantes") >= 1 Then
                Obj.Value = Obj.Value & "Participantes " & linha
                linha = linha + 1
                    
            'Insere os Instrumentos'
            ElseIf InStr(1, Obj.Name, "txtInstrumentos") >= 1 Then
                Obj.Value = Obj.Value & "Instrumentos " & linha
                linha = linha + 1
                
            'Insere as mídias'
            ElseIf InStr(1, Obj.Name, "txtProjecao") >= 1 Then
                Obj.Value = Obj.Value & "Projeção " & linha
                linha = linha + 1
                
            'Insere o tipo de projeção'
            ElseIf InStr(1, Obj.Name, "listBoxTipoProjecao") >= 1 Then
                Obj.Value = Obj.List(1, 0)
                linha = linha + 1
            End If
            
        Next
End Sub
Private Sub btnHtml_Click()
    Dim preenchido As Boolean
    Dim salvo As Boolean
    Dim exporter As clsExporter
    Dim sysManager As clsSystem
    
    preenchido = verificaPreenchimento() 'Verifica se os campos necessários foram preenchidos
    
    If preenchido = True Then
        ' Instâncias
        Set exporter = New clsExporter
        Set sysManager = New clsSystem
    
        criaNovaPlanilha 'Cria uma nova planilha
        
        exporter.writeInfos ' Escreve e estiliza os dados
        Call exporter.adaptaArquivo("html")  ' Realiza os ajustes conforme a extensão
        sysManager.verificaPasta ' Verifica a existência da pasta de salvamento
        nomeValido = exporter.verificaNome(Trim(Me.txtTitulo), "html") 'Retorna um nome válido
        salvo = sysManager.salvaComo(nomeValido, "html") 'Salva o arquivo

        'Se foi salvo com sucesso
        If salvo = True Then
            Call sysManager.abrirArquivos(nomeValido, "html")
            MsgBox "HTML criado com sucesso!", vbOKOnly Or vbInformation
        End If
    End If
End Sub
Private Sub btnTxt_Click()
    Dim preenchido As Boolean
    Dim salvo As Boolean
    Dim exporter As clsExporter
    Dim sysManager As clsSystem
    
    preenchido = verificaPreenchimento() 'Verifica se os campos necessários foram preenchidos
    
    If preenchido = True Then
        ' Instâncias
        Set exporter = New clsExporter
        Set sysManager = New clsSystem
    
        Call criaNovaPlanilha 'Cria uma nova planilha
        
        exporter.writeInfos ' Escreve e estiliza os dados
        exporter.adaptaArquivo ("txt") ' Realiza os ajustes conforme a extensão
        sysManager.verificaPasta ' Verifica a existência da pasta de salvamento
        nomeValido = exporter.verificaNome(Trim(Me.txtTitulo), "txt") 'Retorna um nome válido
        salvo = sysManager.salvaComo(nomeValido, "txt") 'Salva o arquivo

        'Se foi salvo com sucesso
        If salvo = True Then
            Call sysManager.abrirArquivos(nomeValido, "txt")
            MsgBox "TXT criado com sucesso!", vbOKOnly Or vbInformation
        End If
    End If
End Sub
Private Sub btnPdf_Click()
    Dim preenchido As Boolean
    Dim exportado As Boolean
    Dim exporter As clsExporter
    Dim sysManager As clsSystem
    
    preenchido = verificaPreenchimento() 'Verifica se os campos necessários foram preenchidos
    
    If preenchido = True Then
        ' Instâncias
        Set exporter = New clsExporter
        Set sysManager = New clsSystem
    
        Call criaNovaPlanilha 'Cria uma nova planilha
        
        exporter.writeInfos ' Escreve e estiliza os dados
        sysManager.verificaPasta 'Verifica a existencia da pasta e cria caso nao exista
        nomeValido = exporter.verificaNome(Trim(Me.txtTitulo), "pdf") 'Retorna um nome válido
        exportado = exporter.exportaPDF(nomeValido) 'Exporta em PDF
        
        'Se foi salvo com sucesso
        If exportado = True Then
            Call sysManager.abrirArquivos(nomeValido, "pdf") 'Abre o arquivo
            ThisWorkbook.Activate
            MsgBox "PDF criado com sucesso!", vbOKOnly Or vbInformation
        End If
    End If
End Sub
Private Sub btnDoc_Click()
    Dim preenchido As Boolean
    Dim exportado As Boolean
    Dim exporter As clsExporter
    Dim sysManager As clsSystem
    
    preenchido = verificaPreenchimento() 'Verifica se os campos necessários foram preenchidos
    
    If preenchido = True Then
        ' Instâncias
        Set exporter = New clsExporter
        Set sysManager = New clsSystem
    
        Call criaNovaPlanilha 'Cria uma nova planilha
        
        exporter.writeInfos ' Escreve e estiliza os dados
        sysManager.verificaPasta 'Verifica a existencia da pasta e cria caso nao exista
        nomeValido = exporter.verificaNome(Trim(Me.txtTitulo), "doc") 'Retorna um nome válido
        exportado = exporter.exportaDOC(nomeValido, UCase(Me.txtTitulo), Me.txtData) 'Exporta como DOC
        
        'Se foi exportado com sucesso
        If exportado = True Then
            Call sysManager.abrirArquivos(nomeValido, "doc") 'Abre o arquivo
            MsgBox "Arquivo Word (DOC) criado e salvo com sucesso!", vbOKOnly Or vbInformation
        End If
    End If
End Sub
Private Sub btnXlsx_Click()
    Dim preenchido As Boolean
    Dim salvo As Boolean
    Dim exporter As clsExporter
    Dim sysManager As clsSystem

    preenchido = verificaPreenchimento() 'Verifica se os campos necessários foram preenchidos
    
    If preenchido = True Then
        ' Instâncias
        Set exporter = New clsExporter
        Set sysManager = New clsSystem
        
        Call criaNovaPlanilha 'Cria uma nova planilha
        
        exporter.writeInfos  'Escreve e estiliza os dados
        sysManager.verificaPasta 'Verifica a existencia da pasta e cria caso nao exista
        nomeValido = exporter.verificaNome(Trim(Me.txtTitulo), "xlsx") 'Retorna um nome válido
        salvo = sysManager.salvaComo(nomeValido, "xlsx")  'Salva o arquivo
        
        'Se foi salvo com sucesso
        If salvo = True Then
            ThisWorkbook.Activate
            MsgBox "Planilha criada com sucesso!", vbOKOnly Or vbInformation
        End If
    End If
    
End Sub
Private Sub btnImprimir_Click()

    Dim preenchido As Boolean
    Dim exporter As clsExporter
    Dim sysManager As clsSystem

    preenchido = verificaPreenchimento() 'Verifica se os campos necessários foram preenchidos
    
        If preenchido = True Then
            ' Instâncias
            Set exporter = New clsExporter
            Set sysManager = New clsSystem
            
            Call criaNovaPlanilha 'Cria uma nova planilha
            exporter.writeInfos 'Escreve e estiliza os dados
            
            'Se der erro vai para erro:
            On Error GoTo erro
            ActiveWorkbook.ActiveSheet.PrintOut 'Imprime a aba ativa
        End If
        
Exit Sub
erro:
    MsgBox "Não foi possível realizar a impressão automática!" & vbNewLine & "Verifique se alguma impressora está configurada ou gere o arquivo PDF, DOC ou XLSX e imprima manualmente!", vbCritical Or vbOKOnly
End Sub
Private Sub btnPrevious_Click()
    Dim paginaAtual As Range 'Tipagem
    
    'Define a variavel
    Set paginaAtual = ThisWorkbook.Worksheets("controle").Range("B2")
    
    'Se n for a primeira pagina
    If paginaAtual.Value > 1 Then
        paginaAtual.Value = paginaAtual.Value - 1 'Define a pagina atual
        Me.lNumPagina.Caption = Str(paginaAtual) 'Muda o caption do label da página atual
        paginacao paginaAtual.Value 'Chama a função de paginação
        Call redimensiona
    End If
    
End Sub
Private Sub btnNext_Click()
    Dim paginaAtual As Range

    Set paginaAtual = ThisWorkbook.Worksheets("controle").Range("B2")
    
    'Se n for a ultima pagina
    If paginaAtual.Value < WorksheetFunction.RoundUp(contagem.Value / itensPorPagina, 0) Then
        paginaAtual.Value = paginaAtual.Value + 1 'Define a pagina atual
        Me.lNumPagina.Caption = Str(paginaAtual) 'Muda o caption do label da página atual
        paginacao paginaAtual.Value 'Chama a função de paginação
        Call redimensiona
    End If

End Sub
Private Sub UserForm_Initialize()
    Dim aba As Worksheet
    Dim uLin As Integer
    
    Set aba = Worksheets("controle") 'define a aba onde esta os valores
    uLin = aba.Cells(1, 3).End(xlDown).Row 'define a ultima linha preenchida'
    
    ThisWorkbook.Worksheets("controle").Range("B1") = 1 'restaura o contador de itens criados'
    ThisWorkbook.Worksheets("controle").Range("B2") = 1 'restaura a pagina atual'
    
    ' Define os valores das variaveis globais
    tamanhoPadrao = Me.txtDescricao1.Height + 10
    espacamentoPadrao = 20
    Set contagem = ThisWorkbook.Worksheets("controle").Range("B1")
    itensPorPagina = 8 'Itens por pagina
    defaultFormHeight = 195 'Tamanho padrão inicial do form
    
    'Adiciona os valores no listbox
    For i = 2 To uLin
        Me.listBoxTipoProjecao1.AddItem aba.Cells(i, 3)
    Next i

End Sub
Private Sub UserForm_Terminate()
    ThisWorkbook.Worksheets("controle").Range("B1") = 1 'restaura o contador de itens criados'
    ThisWorkbook.Worksheets("controle").Range("B2") = 1 'restaura a pagina atual'
End Sub
