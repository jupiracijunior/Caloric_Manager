Attribute VB_Name = "Utils"
Sub Botão1_Clique()
    Menu.Show
End Sub

Public Sub formatWidthColumns() 'formata a largura de todas as colunas preenchidas para exibir todo o conteudo
    For j = 1 To Worksheets("TMB").Range("A1").End(xlToRight).Column
        Worksheets("TMB").Columns(j).AutoFit
    Next j
End Sub

Public Sub formatStyleCells()
    Range(Range("A1").End(xlDown), ActiveSheet.Range("A1").End(xlToRight)).Select
    
    With Selection.Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = RGB(0, 0, 0)
    End With
    
    Selection.HorizontalAlignment = xlCenter
    Range("A1").Select
End Sub

Public Sub insertValuesOnSheets(nome As String, peso As Double, altura As Integer, idade As Integer, genero As String, fator As Integer, resultadoTMB As Double, gTotal As Double)
    Dim ultimaLinha As Integer
    
    On Error Resume Next
    ultimaLinha = Range("A1").End(xlDown).Row + 1 'dispara um erro caso a primeira linha nao esteja preenchida
    
    'verifica se a primeira linha nao esta vazia
    'primeiro if verifica se nao houve erro na busca da ultima linha
    If Err.Number = 0 Then
        'preenche com null caso o campo nome esteja vazio, se nao preenche com o valor do txtBox
        If nome = "" Then
            Cells(ultimaLinha, 1).Value = "Null"
        Else
            For i = 2 To ultimaLinha
                If nome = Cells(i, 1).Value Then
                    MsgBox "Este nome já foi registrado."
                    Exit Sub 'encerra a execucao do sub caso o nome já exista
                End If
            Next i
            Cells(ultimaLinha, 1).Value = nome
            Cells(ultimaLinha, 2).Value = peso
            Cells(ultimaLinha, 3).Value = altura
            Cells(ultimaLinha, 4).Value = idade
            Cells(ultimaLinha, 5).Value = genero
            Cells(ultimaLinha, 6).Value = intFactorToString(fator)
            Cells(ultimaLinha, 7).Value = resultadoTMB
            Cells(ultimaLinha, 8).Value = gTotal
        End If
        
    Else
        'insere o nome na primeira linha caso esteja vazia
        If nome = "" Then
            Cells(2, 1).Value = "Null"
        Else
            Cells(2, 1).Value = nome
            Cells(2, 2).Value = peso
            Cells(2, 3).Value = idade
            Cells(2, 4).Value = altura
            Cells(2, 5).Value = genero
            Cells(2, 6).Value = intFactorToString(fator)
            Cells(2, 7).Value = resultadoTMB
            Cells(2, 8).Value = gTotal
        End If
    End If
    Err.Clear
    On Error GoTo 0

    Call formatWidthColumns
    Call formatStyleCells
End Sub

'em planejamento
'Public Sub modValuesOnSheets()
'
'End Sub

Function StrFactorToInteger(nameFactor As String)
    Dim result As Integer
    
    Select Case nameFactor
        Case "Sedentário"
            result = 0
        Case "Levemente ativo"
            result = 1
        Case "Moderadamente ativo"
            result = 2
        Case "Altamente ativo"
            result = 3
        Case "Extremamente ativo"
            result = 4
    End Select
    
    StrFactorToInteger = result
End Function

Function intFactorToString(factor As Integer) As String
    Dim lvlActivity() As String
    
    ReDim lvlActivity(0 To 4)
    lvlActivity(0) = "Sedentário"
    lvlActivity(1) = "Levemente ativo"
    lvlActivity(2) = "Moderadamente ativo"
    lvlActivity(3) = "Altamente ativo"
    lvlActivity(4) = "Extremamente ativo"
    
    intFactorToString = lvlActivity(factor)
End Function
