VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "estilizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub cabecalho(ByVal intervalo As String)
    
    With ActiveWorkbook.ActiveSheet.Range(intervalo)
        .Font.bold = True 'Negrito
        .Font.Size = 13 'Tamanho da fonte
        .HorizontalAlignment = xlCenter 'Alinhamento horizontal
        .VerticalAlignment = xlCenter 'Alinhamento vertical
        .BorderAround Color:=vbBlack 'Bordas ao redor
        .RowHeight = 22 'altura da linha
    End With
    
    'mescla o cabecalho
    ActiveWorkbook.ActiveSheet.Cells(1, 1).Resize(1, 4).Merge
    ActiveWorkbook.ActiveSheet.Cells(1, 5).Resize(1, 2).Merge

End Sub

Public Sub titulos(ByVal intervalo As String)
    Dim tamanhoColunaPadrao As Double
    tamanhoColunaPadrao = 8.43

    'redimensiona o tamanho das colunas
    ActiveWorkbook.ActiveSheet.Cells(3, 2).ColumnWidth = tamanhoColunaPadrao * 3 'Descrição
    
    For i = 3 To 5 'colunas participantes, instrumentos e midia
        ActiveWorkbook.ActiveSheet.Cells(1, i).ColumnWidth = tamanhoColunaPadrao * 2
    Next i
    
    'Tira comprimento da coluna A e coloca para a F'
    ActiveWorkbook.ActiveSheet.Cells(1, 1).ColumnWidth = Cells(1, 1).ColumnWidth - 3
    ActiveWorkbook.ActiveSheet.Cells(1, 6).ColumnWidth = Cells(1, 6).ColumnWidth + 3
    
    With ActiveWorkbook.ActiveSheet.Range(intervalo)
        .Font.bold = True 'Negrito
        .Font.Size = 11 'Tamanho da fonte
        .HorizontalAlignment = xlCenter 'Alinhamento horizontal
        .VerticalAlignment = xlCenter 'Alinhamento vertical
        .BorderAround Color:=vbBlack 'Bordas ao redor
    End With
    
    
End Sub

Public Sub corpo(ByVal linha As Integer, ByVal coluna As Integer, Optional ByVal bold As Boolean = False, Optional ByVal tamanho As Integer = 0, Optional ByVal espacamentoLinha As Boolean)
    Dim espacamento_linhas As Double
    
    espacamento_linhas = 15
    
    With ActiveWorkbook.ActiveSheet.Cells(linha, coluna)
        If bold = True Then .Font.bold = True
        If tamanho > 0 Then .Font.Size = tamanho
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Borders(xlEdgeBottom).Color = vbBlack
        If espacamentoLinha = True Then .RowHeight = .RowHeight + espacamento_linhas
    End With
    
End Sub
