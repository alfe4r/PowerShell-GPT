$relatorio = New-Object System.Data.DataTable
$relatorio.Columns.Add("Usuario","string") | Out-Null
$relatorio.Columns.Add("Login","string") | Out-Null
$relatorio.Columns.Add("Grupo","string") | Out-Null
$relatorio.Columns.Add("Status","string") | Out-Null
$relatorio.Columns.Add("OU","string") | Out-Null
$relatorio.Columns.Add("EmployeeID","string") | Out-Null

$usuarios = Get-ADUser -Filter * -Properties EmployeeID

foreach ($usuario in $usuarios) {
    $dn = $usuario.DistinguishedName
    $grupos = Get-ADGroup -LDAPFilter ("(member:1.2.840.113556.1.4.1941:={0})" -f $dn) | Select -ExpandProperty Name

    foreach ($grupo in $grupos) {
        $linha = $relatorio.NewRow()

        $linha.Usuario = $usuario.Name
        $linha.Login = $usuario.samAccountName
        $linha.Grupo = $grupo
        $linha.Status = $usuario.Enabled
        $linha.OU = $usuario.distinguishedname
        $linha.EmployeeID = $usuario.EmployeeID

        $relatorio.Rows.Add($linha)
    }
}

$relatorio | Sort Usuario | ft

# Obter a data atual no formato desejado (04-05-23 no exemplo) e adicionar ao nome do arquivo
$nomeArquivo = "GruposAD-" + (Get-Date -Format "dd-MM-yy") + ".csv"
$relatorio | Export-Csv -Path "C:\temp\$nomeArquivo" -NoTypeInformation -Encoding UTF8

Write-Host "Relatório exportado para C:\temp\$nomeArquivo"
