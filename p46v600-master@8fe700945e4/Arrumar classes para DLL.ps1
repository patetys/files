$thisScript = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
. ($thisScript + '.\functions.ps1')

# Declarando variáveis de ambiente
$CaminhoEntrada = Get-ScriptDirectory
$Inicio=Get-Date
Write-Host "=========================================================================="
Write-Host "== FERRAMENTA DE AJUSTE DE CLASSES PÚBLICAS E PRIVADAS PARA DLL D46O003 =="
Write-Host "=========================================================================="
Write-Host "Sendo executado a partir de $CaminhoEntrada"
Write-Host "=========================================================================="
Write-Host "Variáveis locais: "
$CaminhoDestinoProjeto	= $CaminhoEntrada

# Modificando as permissões dos arquivos.
Write-Host "Modificando as permissões dos arquivos..."
$readonly = [System.IO.FileAttributes]::ReadOnly
Get-ChildItem "$CaminhoDestinoProjeto\*.cls" | ForEach {$_.Set_IsReadOnly($False)}

# Trocando o atributo das classes para Private - A lista dos arquivos deve ser especificada após a chave -include
Write-Host "Trocando o atributo das classes para Private..."
Foreach ($file in Get-Childitem $CaminhoDestinoProjeto\*.cls -include A46V704rpt.cls,A46V705V.cls,A46V701A.cls,A46V703Brpt.cls,A46V715Brpt.cls,A46V720A.cls,A46V721A.cls,A46V722A.cls)
{
	$arquivoVelho = "$CaminhoDestinoProjeto\" + $file.name
	$arquivoNovo = "$CaminhoDestinoProjeto\" + $file.name + "novo"
	Get-Content -Path $arquivoVelho | ForEach-Object { $_ -replace "Attribute VB_Creatable = True", "Attribute VB_Creatable = False" } | Set-Content -Path $arquivoNovo
	Remove-Item -Path $arquivoVelho -force
	Rename-Item -Path $arquivoNovo -newname $arquivoVelho

	Get-Content -Path $arquivoVelho | ForEach-Object { $_ -replace "Attribute VB_Exposed = True", "Attribute VB_Exposed = False" } | Set-Content -Path $arquivoNovo
	Remove-Item -Path $arquivoVelho -force
	Rename-Item -Path $arquivoNovo -newname $arquivoVelho
}

# Trocando o atributo das classes para MultiUse - A lista dos arquivos deve ser especificada após a chave -exclude
Write-Host "Trocando o atributo das classes para MultiUse..."
Foreach ($file in Get-Childitem $CaminhoDestinoProjeto\*.cls -exclude A46V704rpt.cls,A46V705V.cls,A46V701A.cls,A46V703Brpt.cls,A46V715Brpt.cls,A46V720A.cls,A46V721A.cls,A46V722A.cls)
{
	$arquivoVelho = "$CaminhoDestinoProjeto\" + $file.name
	$arquivoNovo = "$CaminhoDestinoProjeto\" + $file.name + "novo"
	Get-Content -Path $arquivoVelho | ForEach-Object { $_ -replace "Attribute VB_Creatable = False", "Attribute VB_Creatable = True" } | Set-Content -Path $arquivoNovo
	Remove-Item -Path $arquivoVelho -force
	Rename-Item -Path $arquivoNovo -newname $arquivoVelho

	Get-Content -Path $arquivoVelho | ForEach-Object { $_ -replace "Attribute VB_Exposed = False", "Attribute VB_Exposed = True" } | Set-Content -Path $arquivoNovo
	Remove-Item -Path $arquivoVelho -force
	Rename-Item -Path $arquivoNovo -newname $arquivoVelho
}