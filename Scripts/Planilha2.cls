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
    
    name = cbNames.List(cbNames.ListIndex)
    
    If Sheets(name).Cells(2, 1).Value <> "" Then
        ultimaLinha = Sheets(name).Range("A1").End(xlDown).Row
    Else
        MsgBox "Não há registros suficientes para esta pessoa"
        Exit Sub
    End If
    
    Set grafic = Worksheets("Dashboard").ChartObjects.Add(500, 30, 500, 250)
    
    grafic.Chart.SetSourceData Source:=Sheets(name).Range("H1:H" & ultimaLinha)
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
