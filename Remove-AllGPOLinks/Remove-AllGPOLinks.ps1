#Written with love by Nate Bender

#Requires -Modules sdm-gpmc

param(
    [Parameter()][string]$GPOName,
    [Parameter()][switch]$FuzzySearch
)

if (-not $GPOName){
    $GPOName = Read-Host "Please enter the name of the GPO you want to remove links to: "
}

#Find all GPOs matching the input. If FuzzySearch is specified, use wildcards in the search.
$GPOs = Get-GPO -All | Where-Object -Property DisplayName -Like $(if ($FuzzySearch){"*$GPOName*"} else {"$GPOName"})

if ($GPOs.Count -gt 0){
    Write-Host "Found $($GPOs.Count) GPOs matching the input." -ForegroundColor Green -BackgroundColor Black
}
else {
     Write-Host "Didnt find any GPOs matching the input.`nUse the -FuzzySearch parameter to find all GPOs containing a term." -ForegroundColor Yellow -BackgroundColor Black
     exit
}


#For each GPO we found in the search:
Foreach ($gpo in $GPOs){
    $links = @()

    #Find the locations that the GPO is linked
    $links += (Get-SDMgplink -DisplayName $gpo.DisplayName)

    if ($links.Count -eq 0){Write-Host "GPO '$($gpo.DisplayName)' is not currently linked to any OUs. Continuing" -ForegroundColor Yellow -BackgroundColor Black ; Continue}
    
    Write-Host "$($links.Count) GPO Links found for GPO '$($gpo.DisplayName)'!`n" -ForegroundColor Green -BackgroundColor Black

    #List each link we found for the current GPO so the user can see them
    Foreach ($link in $links){Write-Host $link.Path }

    if ($(Read-Host "`nDelete the above links? Y or N?: ").ToUpper() -eq "Y"){
        
        #If the user responds affirmatively to the deletion response, try to remove each link in the order they were found.
        Foreach ($link in $links){
            try {
                Remove-SDMgplink -Scope $link.Path -DisplayName $gpo.DisplayName | Out-Null
                Write-Host "`nRemoved link at location $($link.Path) for GPO $($gpo.DisplayName)`n"  -ForegroundColor Green -BackgroundColor Black
            }
            catch {
                Write-Error "`nCouldn't Remove link at location $($link.Path) for GPO $($gpo.DisplayName)`n"
            }
        }
    }
    else {
        #If the user opts to not remove links for the current GPO, move on to the next GPO
        Write-Host "Not removing GPO Links for '$($gpo.DisplayName)'. Quitting." ; Continue
    }
}

Write-Host "Done!" -ForegroundColor Green -BackgroundColor Black