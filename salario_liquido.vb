Function Calcular_Salario(SalarioBruto As Double) As Double
    Dim INSS As Double
    Dim BaseIRRF As Double
    Dim IRRF As Double
    Dim SalarioLiquido As Double
    Dim Excedente As Double
    

    ' Definição das faixas do INSS
    Dim Faixa1 As Double: Faixa1 = 1518
    Dim Faixa2 As Double: Faixa2 = 2793.88
    Dim Faixa3 As Double: Faixa3 = 4190.84
    Dim Faixa4 As Double: Faixa4 = 8157.41


    ' Alíquotas do INSS
    Dim Aliquota1 As Double: Aliquota1 = 0.075
    Dim Aliquota2 As Double: Aliquota2 = 0.09
    Dim Aliquota3 As Double: Aliquota3 = 0.12
    Dim Aliquota4 As Double: Aliquota4 = 0.14


    ' Definição das faixas do IRPF
    Dim Faixa1_IRPF As Double: Faixa1_IRPF = 2259.2
    Dim Faixa2_IRPF As Double: Faixa2_IRPF = 2826.65
    Dim Faixa3_IRPF As Double: Faixa3_IRPF = 3751.05
    Dim Faixa4_IRPF As Double: Faixa4_IRPF = 4664.68

    ' Alíquotas do IRPF
    Dim Aliquota1_IRPF As Double: Aliquota1_IRPF = 0
    Dim Aliquota2_IRPF As Double: Aliquota2_IRPF = 0.075
    Dim Aliquota3_IRPF As Double: Aliquota3_IRPF = 0.15
    Dim Aliquota4_IRPF As Double: Aliquota4_IRPF = 0.225
    Dim Aliquota5_IRPF As Double: Aliquota5_IRPF = 0.275

    'descontos fixos adicionais ao irpf
    Dim fixo_irpf1 As Double: fixo_irpf1 = 169.44
    Dim fixo_irpf2 As Double: fixo_irpf2 = 381.44
    Dim fixo_irpf3 As Double: fixo_irpf3 = 662.77
    Dim teto_IRPF As Double: teto_IRPF = 896


    ' Cálculo do INSS por faixas progressivas
    If SalarioBruto <= Faixa1 Then
        INSS = SalarioBruto * Aliquota1
    ElseIf SalarioBruto <= Faixa2 Then
        INSS = (Faixa1 * Aliquota1) + ((SalarioBruto - Faixa1) * Aliquota2)
    ElseIf SalarioBruto <= Faixa3 Then
        INSS = (Faixa1 * Aliquota1) + ((Faixa2 - Faixa1) * Aliquota2) + ((SalarioBruto - Faixa2) * Aliquota3)
    ElseIf SalarioBruto <= Faixa4 Then
        INSS = (Faixa1 * Aliquota1) + ((Faixa2 - Faixa1) * Aliquota2) + ((Faixa3 - Faixa2) * Aliquota3) + ((SalarioBruto - Faixa3) * Aliquota4)
    Else
        ' Soma das quatro faixas
        INSS = (Faixa1 * Aliquota1) + ((Faixa2 - Faixa1) * Aliquota2) + ((Faixa3 - Faixa2) * Aliquota3) + ((Faixa4 - Faixa3) * Aliquota4)
        ' Adiciona 14% sobre o valor que excede R$ 8.157,41
        'Excedente = SalarioBruto - Faixa4
        INSS = INSS
    End If
    
    ' Base de cálculo do IRRF
    BaseIRRF = SalarioBruto - INSS


    ' Cálculo do IRRF
    If BaseIRRF <= Faixa1_IRPF Then
        IRRF = Aliquota1_IRPF
    ElseIf BaseIRRF <= Faixa2_IRPF Then
        IRRF = BaseIRRF * Aliquota2_IRPF - fixo_irpf1
    ElseIf BaseIRRF <= Faixa3_IRPF Then
        IRRF = BaseIRRF * Aliquota3_IRPF - fixo_irpf2
    ElseIf BaseIRRF <= Faixa4_IRPF Then
        IRRF = BaseIRRF * Aliquota4_IRPF - fixo_irpf3
    Else
        IRRF = BaseIRRF * Aliquota5_IRPF - teto_IRPF
    End If
    
    ' Cálculo do Salário Líquido
    SalarioLiquido = SalarioBruto - INSS - IRRF
    'SalarioLiquido = INSS
    ' Retorno do valor final
    Calcular_Salario = SalarioLiquido
End Function
