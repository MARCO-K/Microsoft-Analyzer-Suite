function Remove-FalsyHashtableEntry {
    <#
    Copy a hashtable with all the falsy entries removed
    @{ x = 'x'; y = '' } -> @{ x = 'x' }
    #>
    param([hashtable]$Hashtable)

    $outTable = @{}

    foreach ($key in $Hashtable.Keys) {
        if ($Hashtable[$key]) {
            $outTable[$key] = $Hashtable[$key]
        }
    }
    $outTable
}