' custom query for power query, advanced editor
let
    // Lista de ativos que você quer consultar (separados por vírgula)
    ativos = "bitcoin,solana,ethereum,cardano,pendle,ondo-finance,chainlink,litecoin",
    
    // Moeda de comparação
    moeda = "usd",
    
    // Construir a URL
    url = Text.Combine({
        "https://api.coingecko.com/api/v3/simple/price?ids=", 
        ativos, 
        "&vs_currencies=", 
        moeda,
        "&include_24hr_change=true"
    }),
    
    // Fazer a requisição
    fonte = Json.Document(Web.Contents(url)),
    
    // Transformar em tabela
    tabela = Record.ToTable(fonte),
    
    // Expandir os valores
    resultados = Table.ExpandRecordColumn(tabela, "Value", {moeda, moeda & "_24h_change"}, {"Preço", "Variação 24h"}),
    
    // Renomear colunas
    resultadoFinal = Table.RenameColumns(resultados, {{"Name", "Ativo"}, {"Preço", "Preço (" & moeda & ")"}, {"Variação 24h", "Variação 24h (%)"}})
in
    resultadoFinal
