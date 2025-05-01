VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function salvaComo(ByVal nome, extensao As String) As Boolean
    salvaComo = True
    
    On Error GoTo restart: 'Se der erro vai para restart
    
    If extensao = "txt" Then
        ActiveWorkbook.SaveAs ThisWorkbook.Path & "/Programas/" & UCase(extensao) & "/" & nome & "." & extensao, xlText 'Salva como Txt
    ElseIf extensao = "html" Then
        ActiveWorkbook.SaveAs ThisWorkbook.Path & "/Programas/" & UCase(extensao) & "/" & nome & "." & extensao, xlHtml 'Salva como Html
    ElseIf extensao = "xlsx" Then
        ActiveWorkbook.SaveAs ThisWorkbook.Path & "/Programas/" & UCase(extensao) & "/" & nome & "." & extensao, xlWorkbookDefault 'Salva como xlsx
    End If
    
    If extensao <> "xlsx" Then ActiveWorkbook.Close False 'Fecha a planilha se não foi selecionado o botão XLSX
    
Exit Function
restart:
    salvaComo = False
    MsgBox "Erro ao salvar arquivo ." & extensao & vbNewLine & vbNewLine & "Verifique o nome (titulo) e tente novamente ou salve manualmente a planilha criada!", vbOKOnly Or vbCritical, "ERRO AO SALVAR"
End Function
Public Sub verificaPasta()
    pasta = ThisWorkbook.Path & "\Programas\"
    
    'Verifica se a pasta que salva os arquivos existe
    If Dir(pasta, vbDirectory) = vbNullString Then 'cria a pasta se não existe
        MkDir ThisWorkbook.Path & "\Programas" ' cria a pasta programas
        MkDir ThisWorkbook.Path & "\Programas\HTML" 'cria a pasta HTML
        MkDir ThisWorkbook.Path & "\Programas\PDF" 'cria a pasta PDF
        MkDir ThisWorkbook.Path & "\Programas\TXT" 'cria a pasta TXT
        MkDir ThisWorkbook.Path & "\Programas\DOC" 'cria a pasta DOC
        MkDir ThisWorkbook.Path & "\Programas\XLSX" 'cria a pasta XLSX
    End If
    
End Sub

Public Sub abrirArquivos(ByVal arquivo As String, ByVal extensao)

Dim objShell As Object

Set objShell = CreateObject("Shell.Application") 'Instancia objeto shell do windows

caminho = ThisWorkbook.Path & "\Programas\" & UCase(extensao) & "\" & arquivo & "." & extensao 'Caminho do arquivo a ser aberto

objShell.Open (caminho) 'Abre o arquivo

End Sub
Public Sub configura_impressao()
    
    'propriedades de impressao EXCEL'
    With ActiveSheet.PageSetup
        ' margens ---
        .LeftMargin = Application.InchesToPoints(0)
        .RightMargin = Application.InchesToPoints(0)
        .TopMargin = Application.InchesToPoints(0.3)
        .BottomMargin = Application.InchesToPoints(0.3)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        ' ----
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintNotes = False
        .CenterHorizontally = True
        .CenterVertically = False
        .Draft = False
        .BlackAndWhite = True
        .FitToPagesWide = 1 'força uma largura de página
        .FitToPagesTall = False
    End With
    
End Sub
