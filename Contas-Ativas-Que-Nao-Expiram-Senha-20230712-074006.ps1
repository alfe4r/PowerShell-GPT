# Obtém a data e a hora atual
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"

# Define o nome do arquivo CSV com base na data e hora atual
$csvPath = "C:\CSV\Contas-Ativas-Que-Nao-Expiram-Senha-$timestamp.csv"

# Obtém informações de usuários do Active Directory
Get-ADUser -Filter * -Properties Name, sAMAccountName, PasswordNeverExpires, PasswordLastSet, EmployeeID, AccountExpirationDate, Enabled, LastLogonDate |
    # Filtra usuários com senhas que nunca expiram e contas ativadas
    Where-Object { $_.PasswordNeverExpires -eq $true -and $_.Enabled -eq $true } |
    # Seleciona as propriedades relevantes dos usuários
    Select-Object sAMAccountName, Name, PasswordLastSet, PasswordNeverExpires, EmployeeID, AccountExpirationDate, Enabled, LastLogonDate, 
        # Adiciona uma propriedade calculada para verificar o status da senha
        @{Name="StatusSenha"; Expression={if ((Get-Date) - $_.PasswordLastSet -lt (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge) { "Válida" } else { "Expirada" }}},
        # Adiciona uma propriedade calculada para verificar o status da conta
        @{Name="StatusConta"; Expression={if ($_.Enabled -eq $true) { "Ativa" } else { "Desativada" }}} |
    # Exporta os resultados para o arquivo CSV
    Export-Csv -Path $csvPath -Encoding UTF8 -NoTypeInformation

# Exibe o caminho do arquivo CSV gerado
Write-Host "Arquivo CSV gerado em: $csvPath"
