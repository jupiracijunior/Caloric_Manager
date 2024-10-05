VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub cbNames_Change()
    Dim grafic As ChartObject
    Dim name As String
    Dim ultimaLinha As Integer
    Dim chartsCount As Integer
    Dim chartsPoint As Integer
    
    name = cbNames.List(cbNames.ListIndex)
    
    If Sheets(name).Cells(2, 1).Value <> "" Then
        ultimaLinha = Sheets(name).Range("A1").End(xlDown).Row
    Else
        MsgBox "Não há registros suficientes para esta pessoa"
        Exit Sub
    End If
    
    chartsCount = Worksheets("Dashboard").ChartObjects.Count
    If chartsCount = 0 Then
        Set grafic = Worksheets("Dashboard").ChartObjects.Add(Range("A1").Width * 3.8, 5, 500, 250)
        grafic.Chart.SetSourceData Source:=Sheets(name).Range("H1:H" & ultimaLinha)
        Exit Sub
    End If
    
    Set grafic = Worksheets("Dashboard").ChartObjects.Add(Range("A1").Width * 3.8, (Worksheets("Dashboard").ChartObjects.Item(chartsCount).Height * chartsCount) + 15, 500, 250)
    grafic.Chart.SetSourceData Source:=Sheets(name).Range("H1:H" & ultimaLinha)
    grafic.Chart.ChartType = xlLine
    grafic.Chart.SeriesCollection(1).ApplyDataLabels ShowValue:=Sheets(name).Range("I2", Sheets(name).Range("I2").End(xlDown)).Value
    
    chartsPoint = grafic.Chart.SeriesCollection(1).Points.Count
    
    For i = 2 To chartsPoint
        grafic.Chart.SeriesCollection(1).Points(i).DataLabel.Text = Sheets(name).Range("I" & i).Value
    Next i
End Sub

Public Sub updateNamesComboBox()
    Dim ultimaLinha As Integer
    
    If Sheets("Registros").Cells(2, 1).Value <> "" Then
        ultimaLinha = Sheets("Registros").Range("A1").End(xlDown).Row
    Else
        Sheets("Dashboard").Select
        Exit Sub
    End If
    
    cbNames.ListFillRange = "Registros!A2:A" & ultimaLinha
End Sub
