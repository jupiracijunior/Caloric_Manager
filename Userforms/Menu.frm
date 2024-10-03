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
'Variaveis de componentes
Private WithEvents ltNomes1 As MSForms.listBox, WithEvents ltNomes2 As MSForms.listBox
Attribute ltNomes1.VB_VarHelpID = -1
Attribute ltNomes2.VB_VarHelpID = -1

Private Sub UserForm_Initialize()
    Set ltNomes1 = Me.Controls.Add("Forms.ListBox.1", "ltNomes", True)
    With ltNomes1
        .Width = 511.25
        .Height = 334.85
        .Top = 10
        .Left = 10
        .Font = "Tahoma"
        .Font.Size = 12
    End With
    
    
    Set ltNomes2 = Me.Controls.Add("Forms.ListBox.1", "ltNomes", False)
    With ltNomes2
        .Width = 511.25
        .Height = 330.85
        .Top = 10
        .Left = 10
        .Font = "Tahoma"
        .Font.Size = 12
    End With
End Sub

Private Sub UserForm_Activate()
    Sheets("Registros").Select
    Call updateOnlyNamesToListBox
End Sub

Private Sub ltNomes1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call modListBox
End Sub

Private Sub ltNomes2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ModRegis.name = ltNomes2.List(ltNomes2.ListIndex)
    ModRegis.indexLine = ltNomes2.ListIndex + 1
    ModRegis.Show
End Sub

'Public nome As String, peso As Double, altura As Integer, idade As Integer

'Private Sub btnAddRegis_Click()
 '   formTMB.Show
'End Sub

'Private Sub btnAltRegis_Click()
'    nome = ltNomes.List(ltNomes.ListIndex)
'
'    For i = 2 To Range("A1").End(xlDown).Row
'        If nome = Cells(i, 1).Value Then
'            For j = 2 To 4
'                Select Case j
'                    Case 2
'                        peso = Cells(i, j).Value
'                    Case 3
'                        altura = Cells(i, j).Value
'                    Case 4
'                        idade = Cells(i, j).Value
'                End Select
'            Next j
'            Exit For
'        End If
'    Next i
'
'    ModRegis.Show
'End Sub

'Private Sub btnRmResgistro_Click()
'        Dim listWidth As Integer
'
'        For i = ltNomes.ListCount - 1 To 0 Step -1
'            If ltNomes.Selected(i) Then
'                Cells(i + 2, 1).EntireRow.Delete
'                ltNomes.RemoveItem (i)
'            End If
'        Next i
'End Sub

Public Sub updateOnlyNamesToListBox()
    Dim ultimaLinha As Integer
    ltNomes1.Clear
    
    If Cells(2, 1).Value <> "" Then
        ultimaLinha = Range("A1").End(xlDown).Row
        
        For i = 2 To ultimaLinha
            ltNomes1.AddItem (Cells(i, 1).Value)
        Next i
    End If
End Sub

Public Sub fullUpdateListBox()
    Sheets(ModRegis.name).Select
    
    Dim ultimaLinha As Integer
    
    If Cells(2, 1).Value <> "" Then
        ultimaLinha = Range("A1").End(xlDown).Row
        
        ltNomes2.RowSource = "A1:H" & ultimaLinha
    End If
End Sub

Sub modListBox()
    Dim name As String, linhaFinal As Integer

    name = ltNomes1.List(ltNomes1.ListIndex)

    ltNomes1.Visible = False
    ltNomes2.Visible = True
    
    Sheets(name).Select

    If Sheets(name).Cells(2, 1) <> "" Then
        linhaFinal = Range("A1").End(xlDown).Row
    End If

    ltNomes2.ColumnCount = 8
    'ltNomes2.ColumnHeads = True
    ltNomes2.RowSource = "A1:H" & linhaFinal
    ltNomes2.ColumnWidths = "50pt; 35pt; 70pt; 40pt; 50pt; 125pt; 100pt; 100pt;"

    name = ""
    linhaFinal = 0
    Sheets("Registros").Select
End Sub
