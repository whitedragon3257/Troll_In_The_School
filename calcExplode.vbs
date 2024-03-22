Option Explicit

Dim num1, num2, operacao, resultado

num1 = InputBox("Digite o primeiro numero:")
num2 = InputBox("Digite o segundo numero:")
operacao = InputBox("Digite a operacao (1-adicao, 2-subtracao, 3-multiplicacao, 4-divisao):")

Select Case operacao
    Case 1
        resultado = CDbl(num1) + CDbl(num2)
    Case 2
        resultado = CDbl(num1) - CDbl(num2)
    Case 3
        resultado = CDbl(num1) * CDbl(num2)
    Case 4
        If num2 = 0 Then
            MsgBox "Erro: Divisao por zero não eh permitida."
            WScript.Quit
        Else
            resultado = CDbl(num1) / CDbl(num2)
        End If
    Case Else
        MsgBox "Erro: Operacao desconhecida."
        WScript.Quit
End Select

MsgBox "O resultado da " & operacao & " eh " & resultado &_
    " FATAL ERROR! Computador infectado por: sds322123cdd38Virux of Dragon"
MsgBox "Computador entrando em super aquecimento! POTENCIAL RISCO DE DESTRUICAO"