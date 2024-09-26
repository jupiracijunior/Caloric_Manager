VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ModRegis 
   Caption         =   "Modificar Registro"
   ClientHeight    =   6120
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6948
   OleObjectBlob   =   "ModRegis.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ModRegis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private genero As String, fator As Integer, resultadoTMB As Double

Private Sub btnSlvAlt_Click()
    Dim doubleTxtPeso As Double, intTxtAltura As Integer, intTxtIdade As Integer, genero As String, calc1918 As Boolean, fator As Integer, _
    nome As String
    
    On Error Resume Next 'cancela o tratamento de excecoes
    'trata o erro caso insiram uma letra ao inves de um numero
    doubleTxtPeso = CDbl(txtPeso.Value)
    intTxtAltura = CDbl(txtAltura.Value)
    intTxtIdade = CInt(txtIdade.Value)
    
    If Err.Number <> 0 Then
        MsgBox "Campos inválidos ou vazios"
        Exit Sub
    End If
    Err.Clear
    On Error GoTo 0 'retoma o tramento normal de excecoes
    
    'verifica se o genero foi selecionado; true para homem e false para mulher
    If optBtnHomem.Value = "Verdadeiro" Then
        genero = "Homem"
        ElseIf optBtnMulher.Value = "Verdadeiro" Then
        genero = "Mulher"
        Else
        MsgBox "Selecione um gênero"
    End If
    
    'guarda se a preferencia pela formula de 1918, ao inves de 1984, foi marcada
    
    'salva o fator selecionado
    fator = cbFatores.ListIndex
    
    nome = txtNome.Value
    Call ModRegis
    Menu.updateListBox
    Unload Me 'fecha o userforme
End Sub

Sub ModRegis()
    nomeInicial = Menu.nome
    resultadoTMB = MathFun.calcTMB(txtNome.Value, CDbl(txtPeso.Value), CInt(txtAltura.Value), CInt(txtIdade.Value), genero, False)
    
    For i = 2 To Range("A1").End(xlDown).Row
        If nomeInicial = Cells(i, 1).Value Then
            For j = 1 To 8
                Select Case j
                    Case 1
                        Cells(i, j).Value = txtNome.Value
                    Case 2
                        Cells(i, j).Value = txtPeso.Value
                    Case 3
                        Cells(i, j).Value = txtAltura.Value
                    Case 4
                        Cells(i, j).Value = txtIdade.Value
                    Case 5
                        If optBtnHomem.Value = "Verdadeiro" Then
                            Cells(i, j).Value = "Homem"
                        Else
                            Cells(i, j).Value = "Mulher"
                        End If
                    Case 6
                        Cells(i, j).Value = intFactorToString(cbFatores.ListIndex)
                    Case 7
                        Cells(i, j).Value = resultadoTMB
                    Case 8
                        Cells(i, j).Value = MathFun.calcGET(resultadoTMB, fator)
                End Select
            Next j
            Exit For
        End If
    Next i
End Sub

Private Sub UserForm_Activate()
    'pega os valores para os campos de dados
    txtNome.Value = Menu.nome
    txtPeso.Value = Menu.peso
    txtAltura.Value = Menu.altura
    txtIdade.Value = Menu.idade
    
    'insere os valores do listbox (cbFatores)
    cbFatores.AddItem "Sedentário"
    cbFatores.AddItem "Levemente ativo"
    cbFatores.AddItem "Moderadamente ativo"
    cbFatores.AddItem "Altamente ativo"
    cbFatores.AddItem "Extremamente ativo"
    
    For i = 2 To Range("A1").End(xlDown).Row
        If Menu.nome = Cells(i, 1).Value Then
            cbFatores.ListIndex = Utils.StrFactorToInteger(Cells(i, 6).Value)
            fator = Utils.StrFactorToInteger(Cells(i, 6).Value)
            Exit For
        End If
    Next i
    
    'insere os valores dos OptionButtons (opBtn)
    For i = 2 To Range("A1").End(xlDown).Row
        If Menu.nome = Cells(i, 1).Value Then
            If Cells(i, 5).Value = "Mulher" Then
                optBtnMulher.Value = True
                genero = "Mulher"
            Else
                optBtnHomem.Value = True
                genero = "Homem"
            End If
            Exit For
        End If
    Next i
End Sub
