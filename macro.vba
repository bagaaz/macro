Public Sub PreencherTrabalho()
    Dim ws As Worksheet
    Dim i As Integer
    Dim Ciclo As String
    Dim Unidade As String
    Dim resultado As Double
    Dim trabalho As Double
    Dim resultado3Anos As Double
    Dim resultado5Anos As Double

    Set ws = ActiveSheet

    ' Percorre as linhas de 3 a 11
    For i = 3 To 11
        ' Pega os valores das colunas E (Ciclo), F (Unidade) e N (Trabalho)
        Ciclo = Trim(ws.Cells(i, 5).Value) ' Coluna E
        Unidade = Trim(ws.Cells(i, 6).Value) ' Coluna F
        trabalho = ws.Cells(i, 14).Value ' Coluna N (Trabalho)

        ' Calcula o valor convertido de ciclo para dias
        resultado = CalcularTrabalhoVBA(Ciclo, Unidade)

        ' Verifica se o resultado é válido antes de calcular
        If resultado > 0 Then
            ' Para 1 ano (365 dias)
            resultado1Ano = (365 / resultado) * trabalho
            resultado1Ano = Int(resultado1Ano) ' Arredonda para inteiro

            ' Para 3 anos (1095 dias)
            resultado3Anos = (1095 / resultado) * trabalho
            resultado3Anos = Int(resultado3Anos)

            ' Para 5 anos (1825 dias)
            resultado5Anos = (1825 / resultado) * trabalho
            resultado5Anos = Int(resultado5Anos)
        Else
            resultado = -1  ' Retorna erro caso os valores sejam inválidos
            resultado3Anos = -1
            resultado5Anos = -1
        End If

        ' Preenche as colunas com os resultados finais
        ws.Cells(i, 15).Value = resultado1Ano   ' Coluna O (1 Ano)
        ws.Cells(i, 16).Value = resultado3Anos  ' Coluna P (3 Anos)
        ws.Cells(i, 17).Value = resultado5Anos  ' Coluna Q (5 Anos)
    Next i

    Set ws = Nothing
End Sub

Function CalcularTrabalhoVBA(Ciclo As String, Unidade As String) As Integer
    Dim valor As Integer
    Dim unidadeTexto As String
    Dim i As Integer

    ' Remover espaços extras
    Ciclo = Trim(Ciclo)
    Unidade = Trim(Unidade)

    ' Se Unidade estiver vazia, retorna -1 para indicar erro
    If Unidade = "" Then
        CalcularTrabalhoVBA = -1
        Exit Function
    End If

    ' Encontrar a posição onde começam as letras (unidade) no Ciclo
    For i = 1 To Len(Ciclo)
        If Not IsNumeric(Mid(Ciclo, i, 1)) Then Exit For
    Next i

    ' Extrair o valor numérico da string Ciclo
    If i > 1 Then
        valor = Val(Left(Ciclo, i - 1))
    Else
        valor = Val(Ciclo) ' Caso seja apenas um número sem unidade
    End If

    ' Extrair a unidade de medida da string Ciclo (se houver)
    unidadeTexto = Trim(Mid(Ciclo, i))

    ' Se a unidade extraída estiver vazia, usar a unidade fornecida no segundo parâmetro
    If unidadeTexto = "" Then
        unidadeTexto = Unidade
    End If

    ' Converter para maiúsculas para evitar problemas de correspondência
    unidadeTexto = UCase(unidadeTexto)

    ' Converter para dias
    Select Case unidadeTexto
        Case "D", "DIA", "DIAS"
            CalcularTrabalhoVBA = valor
        Case "S", "SEM", "SEMANAS"
            CalcularTrabalhoVBA = valor * 7
        Case "M", "MES", "MESES"
            CalcularTrabalhoVBA = valor * 30
        Case "ANO", "ANOS"
            CalcularTrabalhoVBA = valor * 365
        Case Else
            ' Se a unidade não for reconhecida, retorna -1 como erro
            CalcularTrabalhoVBA = -1
    End Select
End Function

