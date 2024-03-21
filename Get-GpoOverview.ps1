<#
.SYNOPSIS
    Creates an Excel file with all Group Policy Objects listed and some basic information for an overview.
.DESCRIPTION
    Creates an Excel file with all Group Policy Objects listed and some basic information for an overview.
.EXAMPLE
    .\Get-GpoOverview.ps1 -Verbose
    Creates an Excel file with all Group Policy Objects listed and some basic information for an overview - verbosly.
#>



#requires -modules ImportExcel
#Requires -runasadministrator
[cmdletBinding()]
param()

function Test-IsGpoComputerDefault {
    [cmdletBinding()]
    param(
        # The xmlgpo object
        [Parameter(Mandatory = $true)]
        [xml]
        $xmlGpo
    )
    #Return Boolean
    $Version = $xmlGpo.GPO.Computer.VersionDirectory
    if ($Version -eq 0) {
        return $true
    }
    elseif ($Version -gt 0 -and ($null -eq $xmlGpo.GPO.Computer.ExtensionData)) {
        return $true
    }
    else {
        return $false
    }
}

function Test-IsGpoUserDefault {
    [cmdletBinding()]
    param(
        # The xmlgpo object
        [Parameter(Mandatory = $true)]
        [xml]
        $xmlGpo
    )
    #Return Boolean
    $Version = $xmlGpo.GPO.User.VersionDirectory
    if ($Version -eq 0) {
        return $true
    }
    elseif ($Version -gt 0 -and ($null -eq $xmlGpo.GPO.User.ExtensionData)) {
        return $true
    }
    else {
        return $false
    }
}

function Get-GpoCounts {
    [cmdletBinding()]
    param(
        # The xmlgpo object
        [Parameter(Mandatory = $true)]
        [xml]
        $xmlGpo
    )
    #Return counts of links, enforcement, and blocks
    if ($null -eq $xmlGpo.GPO.LinksTo) {
        #return 0 linked, 0 enforced, 0 blocked as "0-0-0"
        return "0-0-0"
    }
    else {
        #return total linked, total enforced, total blocked as "x-x-x"
        $LinksCount = $xmlGpo.GPO.LinksTo.Count
        if ($null -eq $LinksCount) {
            $LinksCount = 1
        }
        if ($LinksCount -eq 1 -and $xmlGpo.GPO.LinksTo.Enabled -eq $true) {
            $LinksEnabledCount = 1
        }
        else {
            $LinksEnabledCount = 0
        }
        if ($LinksCount -eq 1 -and $xmlgpo.GPO.LinksTo.NoOverride -eq $true) {
            $LinksEnforcedCount = 1
        }
        else {
            $LinksEnforcedCount = 0
        }
        if ($LinksCount -gt 1) {
            $LinksEnabledCount = (($xmlGpo.GPO.LinksTo).Where({ $_.Enabled -eq $true })).Count
            $LinksEnforcedCount = (($xmlGpo.GPO.LinksTo).Where({ $_.NoOverride -eq $true })).Count
            $Results = "$LinksCount-$LinksEnabledCount-$LinksEnforcedCount"
        }
        else {
            $Results = "$LinksCount-$LinksEnabledCount-$LinksEnforcedCount"
        }
        return $Results
    }
}

#Get all the GPOs
Write-Verbose "Getting all Gpo's in the domain"
$AllGPOs = Get-GPO -All
#Get the GPO information from AD
$AdsiGpo = ([adsisearcher]"objectcategory=groupPolicyContainer").FindAll() | Select-Object -Property *

#Prepare a List variable to store the results
Write-Verbose "Preparing a variable to hold the results."
$AllResults = New-Object System.Collections.Generic.List[psobject]


foreach ($gpo in $AllGPOs) {
    Write-Verbose "Examining the $($gpo.DisplayName) Gpo policy."
    #Note the $_.Properties.displayname is case sensitive
    $GpoAdsiObj = ($AdsiGpo).Where({ $_.Properties.displayname.item(0) -eq $gpo.DisplayName }) 
    $WhenCreated = Get-Date($GpoAdsiObj.Properties.whencreated[0]) -Format "yyyy/MM/dd"
    $gpcfilesyspath = $GpoAdsiObj.Properties.gpcfilesyspath[0]
    $GpoSettings = [PSCustomObject]@{
        #GPO Object
        Name             = $gpo.DisplayName
        Comment          = $gpo.description
        Status           = $gpo.GpoStatus
        WhenCreated      = $WhenCreated
        Modified         = $gpo.ModificationTime
        Owner            = $gpo.Owner
        Comp_Is_Defaults = 'UNKNOWN'
        User_Is_Defaults = 'UNKNOWN'
        WmiFilter        = 'UNKNOWN'
        Link_Info        = 'UNKNOWN'
        Id               = $gpo.Id.ToString()
        gpcfilesyspath   = $gpcfilesyspath
    }
    if ($null -eq $gpo.WmiFilter) {
        $GpoSettings.WmiFilter = "None"
    }
    else {
        $GpoSettings.WmiFilter = $gpo.WmiFilter.Name
    }

    #Create an xml version of the GPO
    [xml]$xmlgpo = $gpo.GenerateReport("xml")

    #Must make the $gpo an xml object before this stage
    $GpoSettings.Comp_Is_Defaults = Test-IsGpoComputerDefault -xmlGpo $xmlgpo
    $GpoSettings.User_Is_Defaults = Test-IsGpoUserDefault -xmlGpo $xmlgpo
    $GpoSettings.Link_Info = Get-GpoCounts -xmlGpo $xmlgpo

    [void]$AllResults.Add($GpoSettings)
}
Write-Debug "Saving results to .\GpoReport.xlsx"
Remove-Item -Path ".\GpoReport.xlsx" -ErrorAction SilentlyContinue
Export-Excel -Path ".\GpoReport.xlsx" -InputObject $AllResults -WorksheetName "GPO Overview" -AutoSize -AutoFilter -FreezeTopRow -TableStyle Medium7
