VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Menu 
   Caption         =   "Alterar registros"
   ClientHeight    =   7068
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10704
   OleObjectBlob   =   "Menu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nome As String, peso As Double, altura As Integer, idade As Integer

Private Sub btnAddRegis_Click()
    formTMB.Show
End Sub

Private Sub btnAltRegis_Click()
    nome = ltNomes.List(ltNomes.ListIndex)
    
    For i = 2 To Range("A1").End(xlDown).Row
        If nome = Cells(i, 1).Value Then
            For j = 2 To 4
                Select Case j
                    Case 2
                        peso = Cells(i, j).Value
                    Case 3
                        altura = Cells(i, j).Value
                    Case 4
                        idade = Cells(i, j).Value
                End Select
            Next j
            Exit For
        End If
    Next i
    
    ModRegis.Show
End Sub

Private Sub btnRmResgistro_Click()
        Dim listWidth As Integer

        For i = ltNomes.ListCount - 1 To 0 Step -1
            If ltNomes.Selected(i) Then
                Cells(i + 2, 1).EntireRow.Delete
                ltNomes.RemoveItem (i)
            End If
        Next i
End Sub

Private Sub UserForm_Activate()
    Sheets("Registros").Select
    Call updateListBox
End Sub

Public Sub updateListBox()
    Dim ultimaLinha As Integer
    ltNomes.Clear
    
    If Cells(2, 1).Value <> "" Then
        ultimaLinha = Range("A1").End(xlDown).Row
        
        For i = 2 To ultimaLinha
            ltNomes.AddItem (Cells(i, 1).Value)
        Next i
    End If
End Sub

Private Sub ltNomes_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call up
End Sub

Sub up()
    Dim name As String, linhaFinal As Integer
    
    name = ltNomes.List(ltNomes.ListIndex)
    
    ltNomes.Clear
    
    Sheets(name).Select
    
    If Sheets(name).Cells(2, 1) <> "" Then
        linhaFinal = Range("A1").End(xlDown).Row
    End If
    
    ltNomes.ColumnCount = 8
    ltNomes.ColumnHeads = True
    ltNomes.RowSource = "A1:H" & linhaFinal
    ltNomes.ColumnWidths = "50pt; 30pt; 100pt; 100pt; 100pt; 120pt; 100pt; 100pt;"
    
    linhaFinal = 0
    Sheets("Registros").Select
    Exit Sub
End Sub
