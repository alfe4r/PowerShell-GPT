# Verificar e importar o módulo ImportExcel (caso necessário)
#if (-not (Get-Module -Name ImportExcel)) {
#    Write-Host "O módulo 'ImportExcel' não está instalado. Instale-o usando o seguinte comando:"
#    Write-Host "Install-Module -Name ImportExcel"
#    return
#}

# Importando o módulo ActiveDirectory e ImportExcel
Import-Module ActiveDirectory
Import-Module ImportExcel

# Definindo o caminho do arquivo .xlsx
$CaminhoArquivo = "C:\Rotinas\Controle de Acesso\Arquivo Excel Exportado do CS\Funcionarios.xlsx"

# Carregando os dados do arquivo .xlsx
$dados = Import-Excel $CaminhoArquivo

# Inicializando um array para armazenar informações sobre as atualizações
$informacoesAtualizadas = @()

# Percorrendo os dados e atualizando o campo "manager" no AD
foreach ($linha in $dados) {
    # Convertendo o campo "RE" para um número inteiro
    $RE = [int]$linha.RE
    $GESTOR = $linha.GESTOR_IMEDIATO

    if ([string]::IsNullOrEmpty($GESTOR)) {
        Write-Host "CAMPO GESTOR VAZIO para RE $($linha.RE). Nenhuma atualização realizada."
        continue  # Pular esta iteração e continuar com a próxima linha de dados
    }

    # Buscando o usuário no AD com base no campo "EmployeeID" correspondente ao "RE"
    $usuarioAD = Get-ADUser -Filter { EmployeeID -eq $RE } -Properties SamAccountName, name, EmployeeID, manager

    if ($usuarioAD) {
        # Buscando o gestor imediato no AD com base no nome correspondente
        $gestorImediatoAD = Get-ADUser -Filter { cn -eq $GESTOR } -Properties distinguishedName, displayName

        if ($gestorImediatoAD) {
            # Obtendo o distinguishedName do gestor imediato
            $distinguishedNameGestor = $gestorImediatoAD.distinguishedName
            $displayNameGestor = $gestorImediatoAD.displayName

            # Atualizando o campo "manager" do usuário com o distinguishedName do gestor imediato
            Set-ADUser -Identity $usuarioAD.SamAccountName -Replace @{
                manager = $distinguishedNameGestor
            }

            # Adicionando informações atualizadas ao array
            $informacoesAtualizadas += @{
                "Usuário" = $usuarioAD.name
                "RE" = $usuarioAD.EmployeeID
                "Informações" = "Campo manager atualizado com o DistinguishedName do gestor imediato."
                "Gestor" = $displayNameGestor
            }

            # Exibindo informações atualizadas em tempo real
            Write-Host "Usuário $($usuarioAD.name) atualizado com sucesso. Gestor: $displayNameGestor."
        } else {
            Write-Host "Gestor imediato '$GESTOR' não encontrado no AD."
        }
    } else {
        Write-Host "Usuário com RE $($linha.RE) não encontrado no AD. Nenhuma atualização realizada."
    }
}

# Exibindo mensagem de sucesso e as informações atualizadas para cada usuário
if ($informacoesAtualizadas.Count -gt 0) {
    Write-Host "As seguintes informações foram adicionadas com sucesso para os usuários:"
    $informacoesAtualizadas | ForEach-Object {
        Write-Host "Usuário: $($_.Usuário)"
        Write-Host "Informações: $($_.Informações)"
        Write-Host "Gestor: $($_.Gestor)"
        Write-Host
    }
} else {
    Write-Host "Nenhuma informação foi adicionada, pois não foram encontrados usuários correspondentes no AD."
}
