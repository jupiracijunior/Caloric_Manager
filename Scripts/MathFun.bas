Attribute VB_Name = "MathFun"
Function calcTMB(nome As String, peso As Double, altura As Integer, idade As Integer, genero As String, calc1918 As Boolean) As Double
    Dim resultadoTMB As Double, gTotal As Double, tbFatores() As Double
    
    ReDim tbFatores(0 To 4)
    tbFatores(0) = 1.2
    tbFatores(1) = 1.375
    tbFatores(2) = 1.55
    tbFatores(3) = 1.725
    tbFatores(4) = 1.9
    
    If calc1918 = True Then
        'calculo de 1918
        If genero = "Homem" Then
            resultadoTMB = (66 + (13.7 * peso) + (5 * altura) - (6.8 * idade))
        Else
            resultadoTMB = (655 + (9.6 * peso) + (1.8 * altura) - (4.7 * idade))
        End If
    Else
        'calculo de 1984
        If genero = "Homem" Then
            resultadoTMB = (88.36 + (13.4 * peso) + (4.8 * altura) - (5.7 * idade))
        Else
            resultadoTMB = (447.6 + (9.2 * peso) + (3.1 * altura) - (4.3 * idade))
        End If
    End If
    
    calcTMB = resultadoTMB
    'Call Utils.insertValuesOnSheets(nome, peso, altura, idade, genero, fator, resultadoTMB, gTotal)
End Function

Function calcGET(resultadoTMB As Double, fator As Integer) As Double
    Dim tbFatores() As Double
    
    ReDim tbFatores(0 To 4)
    tbFatores(0) = 1.2
    tbFatores(1) = 1.375
    tbFatores(2) = 1.55
    tbFatores(3) = 1.725
    tbFatores(4) = 1.9
    
    calcGET = tbFatores(fator) * resultadoTMB
End Function









