VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub estiliza()
    Dim descs As String
    Dim cabecalho As String
    Dim efeitar As clsFormatter
    
    cabecalho = "A1:F1" 'range do cabeçalho
    descs = "A3:F3" 'range das descrições
    
    Set enfeitar = New clsFormatter 'instancia a classe
    
    Call enfeitar.cabecalho(cabecalho)
        
    Call enfeitar.titulos(descs)
        
    Set enfeitar = Nothing 'Destroi a classe

End Sub
Public Sub writeInfos()
    Dim enfeitar As clsFormatter
    Dim sysManager As clsSystem
    Dim contador As Integer
    Dim zerados As Integer
    Dim objs_linha As Integer

    contador = 0 'Acompanha se todos os objetos da linha foram verificados
    zerados = 0 'Acompanha a quantidade de linhas com objetos vazios que foram excluidas
    objs_linha = 0 'Acompanha se todos objetos da linha foram verificados
    
    Set sysManager = New clsSystem
    sysManager.configura_impressao 'Configura parâmetros de impressão
    Set sysManager = Nothing
    
    Call estiliza   'Estiliza a planilha gerada

    With ActiveWorkbook
        'valores cabeçalho
        Cells(1, 1).Value = UCase(cadastro.txtTitulo.Value)
        Cells(1, 5).Value = cadastro.txtData.Value
        
        'titulos
        Cells(3, 1).Value = "N°"
        Cells(3, 2).Value = "DESCRIÇÃO"
        Cells(3, 3).Value = "PARTICIPANTES"
        Cells(3, 4).Value = "INSTRUMENTOS"
        Cells(3, 5).Value = "MÍDIA"
        Cells(3, 6).Value = "PROJEÇÃO"
        
        Set enfeitar = New clsFormatter 'instancia a classe de estilização
        
        'corpo
        For Each Obj In cadastro.Controls 'Loop em cada controle do formulario'
            
            'se o nome do objeto termina com um número
            If IsNumeric(Right(Obj.Name, 2)) = True Then
                linha = Int(Right(Obj.Name, 2) + 3) - zerados 'define o n° do objeto da vez
            'se é um numero < 10
            ElseIf IsNumeric(Right(Obj.Name, 1)) = True Then
                linha = Int(Right(Obj.Name, 1) + 3) - zerados 'define a linha da vez'
            End If
            
            'Insere o n° da sequência'
            If InStr(1, Obj.Name, "lbContador") >= 1 Then
                Cells(linha, 1).Value = Str(Int(Obj.Caption) - zerados) & " -"
                enfeitar.corpo linha, 1, True
                objs_linha = objs_linha + 1
                
            'Insere as descrições'
            ElseIf InStr(1, Obj.Name, "txtDescricao") >= 1 Then 'Se o objeto é um txtBox de descrição
                If Len(Obj.Value) > 0 Then 'Se esta preenchido
                    Cells(linha, 2).Value = RTrim(LTrim(Obj.Value))
                Else
                    contador = contador + 1
                End If
                
                objs_linha = objs_linha + 1
                enfeitar.corpo linha, 2
                
            'Insere os participantes'
            ElseIf InStr(1, Obj.Name, "txtParticipantes") >= 1 Then 'Se é um txtBox dos participantes
                If Len(Obj.Value) > 0 Then 'Se esta preenchido
                    Cells(linha, 3).Value = RTrim(LTrim(Obj.Value))
                Else
                    contador = contador + 1
                End If
                
                objs_linha = objs_linha + 1
                enfeitar.corpo linha, 3
                
            'Insere os Instrumentos'
            ElseIf InStr(1, Obj.Name, "txtInstrumentos") >= 1 Then 'Se é um txtBox dos instrumentos
                If Len(Obj.Value) > 0 Then 'Se esta preenchido
                    Cells(linha, 4).Value = RTrim(LTrim(Obj.Value))
                Else
                    contador = contador + 1
                End If
                
                objs_linha = objs_linha + 1
                enfeitar.corpo linha, 4
                
            'Insere as mídias'
            ElseIf InStr(1, Obj.Name, "txtProjecao") >= 1 Then 'Se é um txtBox da projeção
                If Len(Obj.Value) > 0 Then 'Se esta preenchido
                    Cells(linha, 5).Value = RTrim(LTrim(Obj.Value))
                Else
                    contador = contador + 1
                End If
                
                objs_linha = objs_linha + 1
                enfeitar.corpo linha, 5
                
            'Insere o tipo de projeção'
            ElseIf InStr(1, Obj.Name, "listBoxTipoProjecao") >= 1 Then 'Se é o listBox do tipo de projeção
                If Obj.ListIndex >= 0 Then 'Se tem algum valor selecionado
                    Cells(linha, 6).Value = RTrim(LTrim(Obj.Value))
                Else
                    contador = contador + 1
                End If
                
                objs_linha = objs_linha + 1
                'Se todos campos estiverem vazio, estiliza a linha sem o parametro de espaçamento de linha
                If contador >= 5 Then enfeitar.corpo linha, 6, True, 10 Else enfeitar.corpo linha, 6, True, 10, True
                
            End If
            
            If objs_linha >= 6 Then 'Se todos os objetos da linha foram verificados no for
                If contador >= 5 Then 'se nenhum foi preenchido
                    zerados = zerados + 1 'Aumenta um na variavel de linhas zeradas
                End If
                contador = 0
                objs_linha = 0
            End If
        Next
        
        Set enfeitar = Nothing 'destroi a instância da classe
    End With
    
End Sub
Public Sub adaptaArquivo(ByVal extensao As String)
    Dim uLin As Integer
    Dim coluna As Range
    
    With ActiveWorkbook.ActiveSheet
    
        If extensao = "html" Then
            
            uLin = .Range("A3").End(xlDown).Row 'Ultima linha preenchida
            
            'Iteracao em cada celula das linhas preenchidas
            For Each celula In .Range("A1:F" & uLin)
                celula.WrapText = False
            Next
                
            Range("A1:F" & uLin).Columns.AutoFit 'ajuste automatica das colunas
            ' largura extra
            For Each coluna In Range("A1:F" & uLin).Columns
                coluna.ColumnWidth = coluna.ColumnWidth + 10
            Next coluna
            Range("A4:F" & uLin).Rows.AutoFit 'ajuste automatico das linhas
        'Se for txt exclui o cabeçalho
        ElseIf extensao = "txt" Then
            Range("A3:F3").Clear
        End If
    End With
    
End Sub
Public Function verificaNome(ByVal nome As String, ByVal extensao As String) As String
    Dim arquivo As Variant
    Dim novoNome As String
    Dim contador As Integer
    Dim verificado As Boolean
    
    contador = 2 'Acompanha contagem de itens com mesmo nome
    
    nome = Left(nome, 100) '100 primeiros caracteres
    novoNome = nome
    
    arquivo = ThisWorkbook.Path & "\Programas\" & UCase(extensao) & "\" 'caminho do arquivo
    
    'enquanto não encontrar um nome que não existe
    Do While Dir(arquivo & novoNome & "." & extensao) <> vbNullString 'Se o arquivo com esse nome já existir
        novoNome = nome & "(" & contador & ")"
        contador = contador + 1
    Loop
    
    verificaNome = Trim(novoNome) 'Retorna o nome verificado
    
End Function
Public Function exportaPDF(ByVal nome As String) As Boolean
    exportaPDF = True
    
    On Error GoTo restart: 'Se der erro vai para restart:
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ThisWorkbook.Path & "/Programas/PDF/" & nome & ".PDF" 'Exporta como PDF
    ActiveWorkbook.Close False 'Fecha a planilha temporária
    
Exit Function

restart:
    exportaPDF = False
    MsgBox "Erro ao salvar o arquivo!" & vbNewLine & vbNewLine & "Verifique o nome (titulo) e tente novamente ou salve manualmente a planilha criada!", vbOKOnly Or vbCritical, "ERRO AO SALVAR"
    
End Function
Public Function exportaDOC(ByVal nome, titulo, data As String) As Boolean
    Dim uLin As Integer
    Dim cabecalho As Word.Range
    Dim app As Word.Application
        
    exportaDOC = True
    uLin = ActiveWorkbook.ActiveSheet.Range("A3").End(xlDown).Row 'Atribui a ultima linha preenchida na variavel
    
    ActiveWorkbook.ActiveSheet.Range("A3:F" & uLin).Copy 'copia todos os valores
    
    ActiveWorkbook.Close False 'fecha a planilha
        
    'Se der erro continua para a próxima tentativa
    On Error Resume Next
    Set app = GetObject(, "Word.Application") 'Define o app word caso ja esteja instanciado
    
    'Se der erro exibe a mensagem e cancela a operação
    On Error GoTo ErroAoInstanciar
    
    'Se a primeira tentativa não deu certo, tenta novamente
    If app Is Nothing Then
        Set app = New Word.Application 'Instancia uma nova aplicação word
    End If
        
    app.Visible = True 'torna visivel
    app.Activate 'Ativa a janela
    app.Documents.Add 'Cria o documento
    app.ActiveDocument.Activate
    
    'configura o documento
    app.ActiveDocument.PageSetup.LineNumbering.Active = False
    app.ActiveDocument.PageSetup.Orientation = wdOrientPortrait
    app.ActiveDocument.PageSetup.TopMargin = CentimetersToPoints(0.2)
    app.ActiveDocument.PageSetup.BottomMargin = CentimetersToPoints(0.5)
    app.ActiveDocument.PageSetup.LeftMargin = CentimetersToPoints(1)
    app.ActiveDocument.PageSetup.RightMargin = CentimetersToPoints(1)
    app.ActiveDocument.PageSetup.Gutter = CentimetersToPoints(0)
    app.ActiveDocument.PageSetup.HeaderDistance = CentimetersToPoints(1.25)
    app.ActiveDocument.PageSetup.FooterDistance = CentimetersToPoints(1.25)
    app.ActiveDocument.PageSetup.FirstPageTray = wdPrinterDefaultBin
    app.ActiveDocument.PageSetup.OtherPagesTray = wdPrinterDefaultBin
    app.ActiveDocument.PageSetup.SectionStart = wdSectionNewPage
    app.ActiveDocument.PageSetup.VerticalAlignment = wdAlignVerticalTop
    app.ActiveDocument.PageSetup.SuppressEndnotes = False
    app.ActiveDocument.PageSetup.MirrorMargins = False
    app.ActiveDocument.PageSetup.TwoPagesOnOne = False
    
    Set cabecalho = app.ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range 'define o cabecalho
    
    'estiliza o cabecalho
    With cabecalho
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Text = titulo & vbNewLine & data
        .Font.bold = True
        .Font.Size = 20
        .Font.Name = "Arial"
        .Font.Color = vbBlack
    End With
    
    app.Selection.TypeParagraph 'adiciona um parágrafo
    app.Selection.TypeParagraph 'adiciona um parágrafo
    
    app.Selection.PasteAndFormat wdFormatOriginalFormatting 'cola as informações do excel no word
    
    app.ActiveDocument.SaveAs2 ThisWorkbook.Path & "\Programas\DOC\" & nome & ".doc" 'salva o documento
    
    Set app = Nothing 'Destroi a instancia

Exit Function
ErroAoInstanciar:
    exportaDOC = False
    MsgBox "Erro ao abrir o word:" & vbNewLine & vbNewLine & Err.Description, vbCritical Or vbOKOnly, "ERRO"
End Function
