filter Add-PSType {
    param(
        [string]$TypeName,
        [string[]]$DefaultDisplayPropertySet
    )

    $PSItem.PSTypeNames.Insert(0, $TypeName)
    if ($DefaultDisplayPropertySet -and -not (Get-TypeData $TypeName)) {
        Update-TypeData -TypeName $TypeName -DefaultDisplayPropertySet $DefaultDisplayPropertySet
    }

    return $PSItem
}
