VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formTMB 
   Caption         =   "Formulário TMB"
   ClientHeight    =   7140
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9456.001
   OleObjectBlob   =   "formTMB.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formTMB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCalcTMB_Click()
    Dim nome As String, doubleTxtPeso As Double, intTxtAltura As Integer, intTxtIdade As Integer, _
    genero As String, fator As Integer, resultadoTMB As Double, gTotal As Double, calc1918 As Boolean

    
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
    If ckBox1918.Value = "Verdadeiro" Then
        calc1918 = True
    Else
        calc1918 = False
    End If
    
    'armazena o fator selecionado
    fator = cbFatores.ListIndex
    
    'armazena o valor do campo nome
    nome = txtNome.Value
    
    resultadoTMB = MathFun.calcTMB(nome, doubleTxtPeso, intTxtAltura, intTxtIdade, genero, calc1918)
    Call Utils.insertValuesOnSheets(nome, doubleTxtPeso, intTxtAltura, intTxtIdade, genero, fator, resultadoTMB, MathFun.calcGET(resultadoTMB, fator))
    Menu.updateListBox 'atualiza o listbox no userform Menu
End Sub

Public Sub UserForm_Activate()
    cbFatores.AddItem "Sedentário"
    cbFatores.AddItem "Levemente ativo"
    cbFatores.AddItem "Moderadamente ativo"
    cbFatores.AddItem "Altamente ativo"
    cbFatores.AddItem "Extremamente ativo"
    cbFatores.ListIndex = 0
End Sub
