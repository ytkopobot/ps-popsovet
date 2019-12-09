$manifest = @{
    Path              = '.\ExcelUtils\ExcelUtils.psd1'
    RootModule        = 'ExcelUtils.psm1'
    Author            = 'Kate Dolgikh'
}
New-ModuleManifest @manifest

#powershell ./createManifest.ps1
