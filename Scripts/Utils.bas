Attribute VB_Name = "Utils"
Sub Botão1_Clique()
    Menu.Show
End Sub

Public Sub formatWidthColumns() 'formata a largura de todas as colunas preenchidas para exibir todo o conteudo
    For j = 1 To Worksheets("Registros").Range("A1").End(xlToRight).Column
        Worksheets("Registros").Columns(j).AutoFit
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
    Dim ultimaLinha As Integer, data As String, newSheets As Worksheet
    
    data = Now
    
    For i = 1 To Sheets.Count
        If Sheets(i).name = nome Then
            Sheets(i).Select
            Exit For
        Else
            If i = Sheets.Count Then
                'Cria uma nova aba com o nomeDaAba
                Sheets.Add(After:=Sheets(Sheets.Count)).name = nome
                Sheets(nome).Select
                Exit For
            End If
        End If
    Next i
    
    'Link voltar
    Range("K2").Select
    Range("K2").HorizontalAlignment = xlCenter
    ActiveSheet.Hyperlinks.Add _
        Anchor:=Selection, Address:="", SubAddress:= _
        "Registros!A1", TextToDisplay:="Voltar"
        'Range("F2").Font.Size = 15
    
    On Error Resume Next
    ultimaLinha = Range("A1").End(xlDown).Row + 1
    
    
    'verifica se a primeira linha nao esta vazia
    'primeiro if verifica se nao houve erro na busca da ultima linha
    If Err.Number = 0 Then
        'preenche com null caso o campo nome esteja vazio, se nao preenche com o valor do txtBox
        If nome = "" Then
            MsgBox "Campo nome é obrigatório"
            Exit Sub
        Else
            Worksheets(nome).Cells(ultimaLinha, 1).Value = nome
            Worksheets(nome).Cells(ultimaLinha, 2).Value = peso
            Worksheets(nome).Cells(ultimaLinha, 3).Value = altura
            Worksheets(nome).Cells(ultimaLinha, 4).Value = idade
            Worksheets(nome).Cells(ultimaLinha, 5).Value = genero
            Worksheets(nome).Cells(ultimaLinha, 6).Value = intFactorToString(fator)
            Worksheets(nome).Cells(ultimaLinha, 7).Value = resultadoTMB
            Worksheets(nome).Cells(ultimaLinha, 8).Value = gTotal
            Worksheets(nome).Cells(ultimaLinha, 9).Value = Split(data, " ")(0)
            Worksheets(nome).Cells(ultimaLinha, 10).Value = Split(data, " ")(1)
        End If
        
    Else
        'insere o nome na primeira linha caso esteja vazia
        If nome = "" Then
            Cells(2, 1).Value = "Null"
        Else
            Worksheets(nome).Cells(2, 1).Value = nome
            Worksheets(nome).Cells(2, 2).Value = peso
            Worksheets(nome).Cells(2, 3).Value = idade
            Worksheets(nome).Cells(2, 4).Value = altura
            Worksheets(nome).Cells(2, 5).Value = genero
            Worksheets(nome).Cells(2, 6).Value = intFactorToString(fator)
            Worksheets(nome).Cells(2, 7).Value = resultadoTMB
            Worksheets(nome).Cells(2, 8).Value = gTotal
            Worksheets(nome).Cells(2, 9).Value = Split(Now, " ")(0)
            Worksheets(nome).Cells(2, 10).Value = Split(Now, " ")(1)
        End If
    End If
    Err.Clear
    On Error GoTo 0

    Call formatWidthColumns
    Call formatStyleCells
    
    Sheets("Registros").Select
    Call criarIndice
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

Sub criarIndice()
    Dim planilha As Worksheet
    Dim linha As Integer
    
    linha = 2
    
    For Each planilha In Worksheets
        
        If planilha.name <> "Registros" And planilha.name <> "Dashboard" Then
                    
            Sheets("Registros").Select
            Sheets("Registros").Cells(linha, 1).Select
            ActiveSheet.Hyperlinks.Add _
            Anchor:=Selection, Address:="", SubAddress:= _
            planilha.name & "!A1", TextToDisplay:=planilha.name
            Sheets("Registros").Range("A" & linha).Font.Size = 15
            
            linha = linha + 1
                    
        End If
        
    Next planilha
    
    Columns(1).AutoFit

End Sub
