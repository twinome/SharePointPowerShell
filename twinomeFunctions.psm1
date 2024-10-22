﻿<# 
  _               _                                    
 / |_            (_)                                   
`| |-'_   _   __ __  _ .--.   .--.  _ .--..--.  .---.  
 | | [ \ [ \ [  |  |[ `.-. |/ .'`\ [ `.-. .-. |/ /__\\ 
 | |, \ \/\ \/ / | | | | | || \__. || | | | | || \__., 
 \__/  \__/\__/ [___|___||__]'.__.'[___||__||__]'.__.'                                         
 
/_____/_____/_____/_____/_____/_____/_____/_____/_____/

Script: twinomeFunctions.ps1
Author: Matt Warburton
Date: 03/09/15
Comments: SharePoint functions
#>

#REQUIRES -Version 4.0
#REQUIRES -RunAsAdministrator

Function ApprovedVerb-TWPATTERNErrorHandle {
    <#
    .SYNOPSIS
        Blah
    .DESCRIPTION
        TEMPLATE
    .PARAMETER 1
        Blah
    .PARAMETER 2
        Blah
    .EXAMPLE
        ApprovedVerb-TWPATTERNErrorHandle -site https://speval -lib "customLib"
    #>
    [CmdletBinding()] 
    param (
        [string]$site, 
        [string]$lib
    )
      
    BEGIN {

        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $web = Get-SPWeb $site
            $list = $web.Lists[$lib]

                if($list) {
                    try {
                        Start-Sleep -s 15
                        $list.delete()
                        $web.Update()
                        Write-Output "list $lib deleted"                    
                    }
        
                    catch {
                        $error = $_
                        Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"                   
                    }
                }

                else {
                    Write-Output "list $lib doesnt exist in $site"                
                }
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }

    END {

        $web.dispose()    
    }
} 

Function Add-Library {
    <#
    .SYNOPSIS
        Creates a SharePoint document library 
    .DESCRIPTION
        Allows various parameters to be passed
    .PARAMETER libraryName
        Display name of library
    .PARAMETER libraryUrl
        URL for library
    .PARAMETER webSite
        URL of site that library should be created in
    .PARAMETER libraryTemplate
        Desired template
    .PARAMETER libraryDescription
        Description for library
    .EXAMPLE
        Add-Library -libraryName richtest -libraryUrl rt -webSite https://speval -libraryTemplate "Document Library" -libraryDescription rich library  
    #>
    param (
        [string]$libraryName,
        [string]$libraryUrl,
        [string]$webSite, 
        [string]$libraryTemplate, 
        [string]$libraryDescription
    )
    
    $site = Get-SPweb -Identity $webSite -ErrorAction SilentlyContinue

        if($site){
             $listTemplate = $site.ListTemplates[$libraryTemplate]
                
                if($listTemplate){
                $exists = $site.Lists[$libraryName]
                    
                    if(!$exists){
                        $site.Lists.Add($libraryUrl,$libraryDescription,$listTemplate)
                        $list = $site.Lists[$libraryUrl]
                        $list.Title = $libraryName
                        $list.Update()
                        Write-Output "Create-Library - Success! $libraryName added to $webSite"
                    }
                    
                    else{
                        Write-Output "Create-Library - $libraryName already exists with url $libraryUrl, please check the name/url and try again"
                    }
                }
                else{
                    Write-Output "Create-Library - $libraryTemplate doesn't exist, please check the name and try again" 
                } 
        }

        else{
            Write-Output "Create-Library - site doesn't exist, please check the url and try again"
        }
}

Function Add-CTLibrary {
    <#
    .SYNOPSIS
        Adds a content type to a library 
    .DESCRIPTION
        Add-CTLibrary
    .PARAMETER libraryName
        Display name of library to add content type to
    .PARAMETER webSite
        URL of site that library lives in
    .PARAMETER contentType
        Desired content type
    .EXAMPLE
        Add-CTLibrary -libraryName richtest -webSite https://speval -contentType "Report" 
    #>
    param (
        [string]$libraryName,
        [string]$webSite,
        [string]$contentType
    )

    $site = Get-SPweb -Identity $webSite -ErrorAction SilentlyContinue

        if($site){
            $ct = $site.Site.RootWeb.ContentTypes[$contentType]

                if($ct){
                    $list = $site.Lists[$libraryName]

                    if($list){
                       $allowCT = $list.ContentTypesEnabled
                       $exits = $list.ContentTypes[$contentType]
                       
                        if(!$exits){ 
                            if($allowCT -eq $false){
                                $list.ContentTypesEnabled = $true
                                $list.Update()
                                Write-Output "Add-CTLibrary - Enabled content types in $libraryName"
                                $list.ContentTypes.Add($ct)
                                Write-Output "Add-CTLibrary - Success! $contentType added to $libraryName in $webSite"          
                            }

                            else{
                                $list.ContentTypes.Add($ct)
                                Write-Output "Add-CTLibrary - Success! $contentType added to $libraryName in $webSite"                            
                            }
                        }
                        
                        else{
                        Write-Output "Add-CTLibrary - content type already exists, please check and try again"
                        }
                    }
                    
                    else{
                        Write-Output "Add-CTLibrary - list doesn't exist, please check and try again"
                    }
                }

                else{
                    Write-Output "Add-CTLibrary - content type doesn't exist, please check and try again" 
                }
        }

        else{
            Write-Output "Add-CTLibrary - site doesn't exist, please check the url and try again" 
        }
}

Function Remove-ContentTypeList {
    <#
    .SYNOPSIS
        Removes a content type from a list
    .DESCRIPTION
        Remove-ContentTypeList
    .PARAMETER site
        URL for site
    .PARAMETER listName
        Display name of library
    .PARAMETER contentType
        Content type name
    .EXAMPLE
        Remove-ContentTypeList -listName richtest -site https://speval -contentType "Report"
    #>
    param (
        [string]$listName,
        [string]$site,
        [string]$contentType
    )

    $web = Get-SPweb -Identity $site -ErrorAction SilentlyContinue

        if($web){
            $list = $web.Lists[$listName]
        
            if($list){
                $exits = $list.ContentTypes[$contentType]

                if($exits){
                    $list.ContentTypes.Delete($exits.id)
                    $list.Update()
                    Write-Output "$contentType deleted"
                }          
                           
                else{
                    Write-Output "$contentType content type type not found"                       
                }
            }

            else{
                Write-Output "list not found"
            }
        }
        else{
            Write-Output "site not found"
        }
}

Function Add-DLasTemplate {
    <#
    .SYNOPSIS
        Script saves a list or library as a template
    .DESCRIPTION
        Add-DLasTemplate
    .PARAMETER libraryName
        Display name of library to remove the content type from
    .PARAMETER webSite
        URL of site that library lives in
    .PARAMETER templateName
        Template file name i.e blah.stp
    .PARAMETER templateTitle
        Template display name
    .PARAMETER templateDescripton
        Template description
    .PARAMETER content
        Include content, options are yes or no
    .EXAMPLE
        Add-DLasTemplate https://speval -templateName "richtest.stp" -templateTitle "rich template" -templateDescripton "Richads template" -content yes/no
    #>   
    param (
        [string]$libraryName, 
        [string]$webSite, 
        [string]$templateName, 
        [string]$templateTitle, 
        [string]$templateDescripton, 
        [string]$content
    )
    
    Start-SPAssignment –Global
    
    $site = Get-SPweb -Identity $webSite -ErrorAction SilentlyContinue

        if($site -ne $null){
            $list = $site.Lists[$libraryName]
        
            if($list -ne $null){

                if($content -eq "yes"){
                    $list.SaveAsTemplate($templateName,$templateTitle,$templateDescripton,1)
                    write-host "Save-DLasTemplate - Success! $templateTitle saved (with content)" -BackgroundColor "Green" -ForegroundColor "White"
                }

                else{
                    $list.SaveAsTemplate($templateName,$templateTitle,$templateDescripton,0)
                    write-host "Save-DLasTemplate - Success! $templateTitle saved (without content)" -BackgroundColor "Green" -ForegroundColor "White"      
                }
            }

            else{
                Write-Host "Save-DLasTemplate - list doesn't exist, please check and try again" -BackgroundColor "Red" -ForegroundColor "White" 
            }
        }

        else{
        Write-Host "Save-DLasTemplate - site doesn't exist, please check the url and try again" -BackgroundColor "Red" -ForegroundColor "White"  
        }

    Stop-SPAssignment –Global

}

Function Get-AllLibraries {
    <#
    .SYNOPSIS
        Script to produce CSV report of all libraries in a site collection
    .DESCRIPTION
        Get-AllLibraries
    .PARAMETER siteCollection
        Desired site collection
    .PARAMETER outPath
        Path that CSV file will be saved to
    .EXAMPLE
        Get-AllLibraries -siteCollection https://speval -outPath "C:\Users\userName\Desktop"
    #>   
    param (
        [string]$siteCollection, 
        [string]$outPath
    )
            $date = Get-Date -Format "ddmmyy"

            $name = "_Get-AllLibraries.csv"
            $pathStatus = Test-Path -Path $outPath -ErrorAction SilentlyContinue
                if($pathStatus -eq $true){
                    set-variable -option constant -name out -value "$outPath\$date$name"
                    "sep=;" | Out-File $out 
                    "library;parent;url;content type;parent url" | Out-File $out -append

                    Start-SPAssignment –Global

                    $site = Get-SPsite -Identity $siteCollection -ErrorAction SilentlyContinue

                        if($site -ne $null){
                            $webs = Get-SPSite -Identity $siteCollection -ErrorAction SilentlyContinue | Get-SPWeb -Limit All -WarningAction SilentlyContinue
                            
                                foreach($web in $webs){
                                    $libraries = $web.Lists

                                    foreach($library in $libraries){

                                        if($library.BaseType -eq "DocumentLibrary"){
                                            $title = $library.title
                                            $parent = $library.parentweb
                                            $url = $library.DefaultViewUrl
                                            $contentType = $library.ContentTypes.name
                                            $contentWeb = $library.ParentWeb.Url
                                            $title + ";" + $parent + ";" + "$url" + ";" + $contentType + ";" + $contentWeb + ";" | Out-File $out -append
                                        }
                                    }
                                }

                    Stop-SPAssignment –Global

                    write-host "Success! get the output here : $out" -BackgroundColor "Green" -ForegroundColor "White" 
                        }
        
                        else{
                            Write-Host "Site collection doesn't exist, please check the url and try again" -BackgroundColor "Red" -ForegroundColor "White"  
                        }

                }

                else{
                    write-host "dodgy path!, please check the path and try again" -BackgroundColor "Red" -ForegroundColor "White"  
                }
}

Function Get-AllLists {
    <#
    .SYNOPSIS
        Script to produce CSV report of all lists in a site collection
    .DESCRIPTION
        Get-AllLists
    .PARAMETER siteCollection
        Desired site collection
    .PARAMETER outPath
        Path that CSV file will be saved to
    .EXAMPLE
        Get-AllLists -siteCollection https://speval -outPath "C:\Users\userName\Desktop"
    #>  
    param (
        [string]$siteCollection, 
        [string]$outPath
    )
            $date = Get-Date -Format "ddmmyy"
            $name = "_Get-AllLists.csv"
            $pathStatus = Test-Path -Path $outPath -ErrorAction SilentlyContinue

                if($pathStatus -eq $true){
                    set-variable -option constant -name out -value "$outPath\$date$name"
                    "sep=;" | Out-File $out 
                    "library;parent;url;content type" | Out-File $out -append

                    Start-SPAssignment –Global

                    $site = Get-SPsite -Identity $siteCollection -ErrorAction SilentlyContinue

                        if($site -ne $null){
                            $webs = Get-SPSite -Identity $siteCollection -ErrorAction SilentlyContinue | Get-SPWeb -Limit All -WarningAction SilentlyContinue
                            
                                foreach($web in $webs){
                                    $lists = $web.Lists

                                    foreach($list in $lists){

                                        if($list.BaseType -eq "GenericList"){
                                            $title = $list.title
                                            $parent = $list.parentweb
                                            $url = $list.DefaultViewUrl
                                            $contentType = $list.ContentTypes.name
                                            $parentWeb = $list.ParentWeb.Url
                                            $title + ";" + $parent + ";" + "$url" + ";" + $contentType + ";" + $parentWeb + ";" | Out-File $out -append
                                        }
                                    }
                                }

                    Stop-SPAssignment –Global

                    write-host "Success! get the output here : $out" -BackgroundColor "Green" -ForegroundColor "White" 
                        }
        
                        else{
                            Write-Host "Site collection doesn't exist, please check the url and try again" -BackgroundColor "Red" -ForegroundColor "White"  
                        }

                }

                else{
                    write-host "dodgy path!, please check the path and try again" -BackgroundColor "Red" -ForegroundColor "White"  
                }
}

Function Remove-CTs {
    <#
    .SYNOPSIS
        Script to remove a specified content type from all libraries in a site collection
    .DESCRIPTION
        Get-AllLists
    .PARAMETER contentType
        Content type that you want to remove
    .PARAMETER siteCollection
        Desired site collection
    .EXAMPLE
        Remove-CTs -contentType "the content type" -siteCollection https://speval
    #>  
    param (
        [string]$contentType, 
        [string]$siteCollection
    )
        $site = Get-SPSite $siteCollection -ErrorAction SilentlyContinue

            if($site -ne $null){
                $web = $site.RootWeb
                $ct = $web.ContentTypes["$contentType"]

                if($ct -ne $null){
                    $ct.Delete()
                    write-host "Success! $contentType deleted from $siteCollection" -BackgroundColor "Green" -ForegroundColor "White"     

                }
            
                else{
                    write-host "$contentType doesn't exist in $siteCollection, please check the name and try again" -BackgroundColor "Red" -ForegroundColor "White"  
                }
            }
        
            else{
            Write-Host "site doesn't exist, please check the url and try again" -BackgroundColor "Red" -ForegroundColor "White"  
            }
}

Function Get-ColStatus {
    <#
    .SYNOPSIS
        Script to produce CSV report of status of paricular column in all libraries in a site collection (option or manditory)
    .DESCRIPTION
        Get-ColStatus
    .PARAMETER siteCollection
        Desired site collection
    .PARAMETER column
        Column to generate report for
    .PARAMETER outPath
        Path that CSV file will be saved to
    .EXAMPLE
        Get-ColStatus -siteCollection https://speval -outPath "C:" -column "Document Date"
    #>  
    param (
        [string]$siteCollection, 
        [string]$outPath, 
        [string]$column
    )
            $date = Get-Date -Format "ddmmyy"
            $name = "_Get-ColStatus.csv"
            $pathStatus = Test-Path -Path $outPath -ErrorAction SilentlyContinue

                if($pathStatus -eq $true){
                    set-variable -option constant -name out -value "$outPath\$date$name"
                    "sep=;" | Out-File $out 
                    "library;parent;url;content type;parent url;required" | Out-File $out -append

                    Start-SPAssignment –Global

                    $site = Get-SPsite -Identity $siteCollection -ErrorAction SilentlyContinue

                        if($site -ne $null){
                            $webs = Get-SPSite -Identity $siteCollection -ErrorAction SilentlyContinue | Get-SPWeb -Limit All -WarningAction SilentlyContinue
                            
                                foreach($web in $webs){
                                    $libraries = $web.Lists

                                    foreach($library in $libraries){

                                        if($library.BaseType -eq "DocumentLibrary"){
                                            $col = $library.fields[$column]
                                            
                                                if($col -ne $null){

                                                    $title = $library.title
                                                    $parent = $library.parentweb
                                                    $url = $library.DefaultViewUrl
                                                    $contentType = $library.ContentTypes.name
                                                    $contentWeb = $library.ParentWeb.Url
                                                    $status = $col.Required
                                                    $title + ";" + $parent + ";" + "$url" + ";" + $contentType + ";" + $contentWeb + ";" + $status + ";" | Out-File $out -append
                                                    write-host "$column found in $library" -BackgroundColor "Green" -ForegroundColor "White" 
                                                }

                                                else{
                                                    Write-Host "Column $column doesn't exist in $library" -BackgroundColor "Red" -ForegroundColor "White"     
                                                }

                                        }
                                    }
                                }

                    Stop-SPAssignment –Global

                    write-host "Success! get the output here : $out" -BackgroundColor "Green" -ForegroundColor "White" 
                        }
        
                        else{
                            Write-Host "Site collection doesn't exist, please check the url and try again" -BackgroundColor "Red" -ForegroundColor "White"  
                        }

                }

                else{
                    write-host "dodgy path!, please check the path and try again" -BackgroundColor "Red" -ForegroundColor "White"  
                }
}

Function Get-ContentTypeColStatusLibrary {
    <#
    .SYNOPSIS
        Script to produce CSV report of status of paricular column in all libraries in a site collection (option or manditory)
    .DESCRIPTION
        Get-ColStatus
    .PARAMETER siteCollection
        Desired site collection
    .PARAMETER column
        Column to generate report for
    .PARAMETER outPath
        Path that CSV file will be saved to
    .EXAMPLE
        Get-ContentTypeColStatusLibrary -siteCollection https://speval -outPath "C:" -column "Document Date"
    #> 
    param (
        [string]$siteCollection, 
        [string]$outPath, 
        [string]$column
    )
            $date = Get-Date -Format "ddmmyy"
            $name = "_Get-ContentTypeColStatusLibrary.csv"
            $pathStatus = Test-Path -Path $outPath -ErrorAction SilentlyContinue
        
            if($pathStatus -eq $true){
            set-variable -option constant -name out -value "$outPath\$date$name"
            "sep=;" | Out-File $out 
            "library;parent;url;content type;parent url;required" | Out-File $out -append

            $webs = Get-SPSite -Identity $siteCollection | Get-SPWeb -Limit All 

            foreach($web in $webs){

            $libs = $web.Lists

                    foreach($lib in $libs){
                        $cts = $lib.ContentTypes
        
                        foreach($ct in $cts){
                        $col = $ct.Fields[$column]
            
                            if($col -ne $null){
                                $title = $lib.title
                                $parent = $lib.parentweb
                                $url = $lib.DefaultViewUrl
                                $contentType = $ct.Name
                                $contentWeb = $lib.ParentWeb.Url
                                $status = $col.Required
                                $title + ";" + $parent + ";" + "$url" + ";" + $contentType + ";" + $contentWeb + ";" + $status + ";" | Out-File $out -append
                            }
                        }
                    }
                    }
            }
}

Function Delete-AuditALL {
    <#
    .SYNOPSIS
        Purges audit log for all site collections in farm
    .DESCRIPTION
        Delete-AuditALL
    .PARAMETER date
        Purges everything before date specified
    .EXAMPLE
        Delete-AuditALL -date "10/09/15" (purges everything before date specified)
    #> 
    param (
        [string]$date
    )
            $SCs = Get-SPsite -Limit all

                foreach($SC in $SCs){
                    $sc.Audit.DeleteEntries($date)
                    $sc.Audit.Update()
                    write-host "Success! $sc purged" -BackgroundColor "Green" -ForegroundColor "White" 
                }
}

Function Delete-SCAudit {
    <#
    .SYNOPSIS
        Purges audit log for a particular site collection
    .DESCRIPTION
        Delete-AuditALL
    .PARAMETER date
        Purges everything before date specified
    .PARAMETER siteCollection
        Site collection to purge audit log
    .EXAMPLE
        Delete-SCAudit -siteCollection https://speval -date "10/09/15" (purges everything before date specified)
    #>
    param (
        [string]$siteCollection,
        [string]$date
    )
            $sc = Get-SPsite -identity $siteCollection
            $sc.Audit.DeleteEntries($date)
            $sc.Audit.Update()
            write-host "Success! $sc purged" -BackgroundColor "Green" -ForegroundColor "White" 
}

Function Disable-AuditAll {
    <#
    .SYNOPSIS
        Switches off auditing in all sites collections in farm
    .DESCRIPTION
        Disable-AuditAll
    .EXAMPLE
        Disable-AuditAll
    #>
    Start-SPAssignment –Global
    $SCs = Get-SPsite -Limit all

        foreach($SC in $SCs){

                if($sc.Audit.AuditFlags -ne "None"){
                    $sc.Audit.AuditFlags = [Microsoft.SharePoint.SPAuditMaskType]::None
                    $sc.Audit.Update()
                    write-host "Success! auditing disabled in $sc" -BackgroundColor "Green" -ForegroundColor "White" 
                }
                
                else{
                    write-host "Auditing already disabled in $sc" -BackgroundColor "Red" -ForegroundColor "White" 
                }
        }
    
    Stop-SPAssignment –Global
}

Function Enable-AuditAll {
    <#
    .SYNOPSIS
        Switches on auditing in all sites collections in farm
    .DESCRIPTION
        Enable-AuditAll
    .EXAMPLE
        Enable-AuditAll
    #>
    Start-SPAssignment –Global
    $SCs = Get-SPsite -Limit all

        foreach($SC in $SCs){

                if($sc.Audit.AuditFlags -ne "All"){
                    $sc.Audit.AuditFlags = [Microsoft.SharePoint.SPAuditMaskType]::All
                    $sc.Audit.Update()
                    write-host "Success! full auditing enabled in $sc" -BackgroundColor "Green" -ForegroundColor "White" 
                }
                
                else{
                    write-host "Full auditing already disabled in $sc" -BackgroundColor "Red" -ForegroundColor "White" 
                }
        }

    Stop-SPAssignment –Global
}

Function Get-ContentTypeColStatusLibraryAllSites {
    <#
    .SYNOPSIS
        Script to produce CSV report of status of paricular column in all libraries in a particular web application (optional or manditory)
    .DESCRIPTION
        Get-ContentTypeColStatusLibraryAllSites
    .PARAMETER webApp
        Desired web application
    .PARAMETER column
        Column to generate report for
    .PARAMETER outPath
        Path that CSV file will be saved to
    .EXAMPLE
        Get-ContentTypeColStatusLibraryAllSites -webApp https://speval -outPath "C:" -column "Document Date"
    #>
    param (
        [string]$webApp, 
        [string]$outPath, 
        [string]$column
    )
            $date = Get-Date -Format "ddmmyy"
            $name = "_Get-ContentTypeColStatusLibrary.csv"
            $pathStatus = Test-Path -Path $outPath -ErrorAction SilentlyContinue
        
            if($pathStatus -eq $true){
            set-variable -option constant -name out -value "$outPath\$date$name"
            "sep=;" | Out-File $out 
            "library;parent;url;content type;parent url;required" | Out-File $out -append

            Start-SPAssignment –Global
            $webs = Get-SPSite -webapplication $webApp -Limit All | Get-SPWeb -Limit All

            foreach($web in $webs){

            $libs = $web.Lists

                    foreach($lib in $libs){
                        $cts = $lib.ContentTypes
        
                        foreach($ct in $cts){
                        $col = $ct.Fields[$column]
            
                            if($col -ne $null){
                                $title = $lib.title
                                $parent = $lib.parentweb
                                $url = $lib.DefaultViewUrl
                                $contentType = $ct.Name
                                $contentWeb = $lib.ParentWeb.Url
                                $status = $col.Required
                                $title + ";" + $parent + ";" + "$url" + ";" + $contentType + ";" + $contentWeb + ";" + $status + ";" | Out-File $out -append
                            }
                        }
                        Stop-SPAssignment –Global
                    }
                    }
            }
}

Function Get-SitesWithNoAssociatedGroups {
    <#
    .SYNOPSIS
        Script to produce CSV report of sites with no associated permission groups
    .DESCRIPTION
        Get-SitesWithNoAssociatedGroups
    .PARAMETER webApp
        Desired web application
    .EXAMPLE
        Get-SitesWithNoAssociatedGroup -webApp https://speval
    #>
    param (
        [string]$webApp
    )
            $date = Get-Date -format "dd-MMM-yyyy"
            $filePath = "sitesWithNoAssociatedGroups_$date.CSV" 
            "sep=;" | Out-File $filePath
            "Title;Path" | Out-File $filePath -append
            Start-SPAssignment –Global
            $SCs = Get-SPSite -WebApplication $webApp -Limit ALL
    
                foreach ($SC in $SCs) {
                    $url = $SC.url
                    $site = Get-SPWeb ($url)
        
                    foreach ($web in $site.Site.AllWebs) { 
                        $member = $web.AssociatedMemberGroup
                        $owner = $web.AssociatedOwnerGroup
                        $visitor = $web.AssociatedVisitorGroup
                                if(!$member -or !$owner -or !$visitor) {
                                    $title = $web.Title
                                    $url = $web.Url
                                        $title + ";" + $url + ";" | Out-File $filePath -append
                                }
                    }
                }
                Stop-SPAssignment –Global
}

Function Get-ListGuids {
    <#
    .SYNOPSIS
        Script to produce CSV report of list GUID's
    .DESCRIPTION
        Get-ListGuids
    .PARAMETER webApp
        Desired web application
    .PARAMETER libraryName
        Display name of library you want GUID's for
    .EXAMPLE
        Get-ListGuids -webApp https://speval -libraryName "access request lists"
    #>
    param (
        [string]$webApp, 
        [string]$libraryName
    )
            $date = Get-Date -format "dd-MMM-yyyy"
            $filePath = "‪AccessRequests_$date.CSV" 
            "sep=;" | Out-File $filePath
            "Title;Path;GUID" | Out-File $filePath -append
            Start-SPAssignment –Global
            $SCs = Get-SPSite -WebApplication $webApp -Limit ALL
    
                foreach ($SC in $SCs) {
                    $url = $SC.url
                    $site = Get-SPWeb ($url)
        
                    foreach ($web in $site.Site.AllWebs) { 
                        $lists = $web.Lists
            
                        foreach ($list in $lists) {
                    
                                if($list.Title -eq $libraryName) {
                                    $title = $list.Title
                                    $listPath = $list.DefaultViewUrl
                                    $ID = $list.ID
                                        $title + ";" + $listPath + ";" + $ID + ";"| Out-File $filePath -append
                                      
                                }
                        }
                    }
                }
                Stop-SPAssignment –Global
}

Function Disable-SandboxedSolution {
    <#
    .SYNOPSIS
        Script to deactivate sandboxed solution
    .DESCRIPTION
        Disable-SandboxedSolution
    .PARAMETER siteCollection
        Site collection that contains the solution
    .PARAMETER solution
        Display name of solution
    .EXAMPLE
        Disable-SandboxedSolution -siteCollection https://speval -solution "A Custom solution"
    #>
    param (
        [string]$siteCollection, 
        [string]$solution
    )
            $sc = get-spsite -identity $siteCollection
            Uninstall-SPUserSolution -Identity $solution -Site $sc -WarningAction SilentlyContinue
            write-host "Success! $solution deactivated from $siteCollection" -BackgroundColor "Green" -ForegroundColor "White" 
}

Function Remove-SandboxedSolution {
    <#
    .SYNOPSIS
        Script to remove a sandboxed solution
    .DESCRIPTION
        Remove-SandboxedSolution
    .PARAMETER siteCollection
        Site collection that contains the solution
    .PARAMETER solution
        Display name of solution
    .EXAMPLE
        Remove-SandboxedSolution -siteCollection https://speval -solution "A Custom solution"
    #>
    param (
        [string]$siteCollection, 
        [string]$solution
    )
            $sc = get-spsite -identity $siteCollection
            Remove-SPUserSolution -Identity $solution -Site $sc -WarningAction SilentlyContinue
            write-host "Success! $solution removed from $siteCollection" -BackgroundColor "Green" -ForegroundColor "White" 
}

Function Enable-SandboxedSolution {
    <#
    .SYNOPSIS
        Script to activate sandboxed solution
    .DESCRIPTION
        Disable-SandboxedSolution
    .PARAMETER siteCollection
        Site collection that contains the solution
    .PARAMETER solution
        Display name of solution
    .EXAMPLE
        Enable-SandboxedSolution -siteCollection https://speval -solution "A Custom solution"
    #>
    param (
        [string]$siteCollection, 
        [string]$solution
    )
            $sc = get-spsite -identity $siteCollection
            Install-SPUserSolution -Identity $solution -Site $sc
            write-host "Success! $solution activated from $siteCollection" -BackgroundColor "Green" -ForegroundColor "White" 
}

Function Remove-AllSandboxedSolutions {
    <#
    .SYNOPSIS
        Script to deactivate and remove all sandboxed solutions in a particular site collection
    .DESCRIPTION
        Remove-AllSandboxedSolutions
    .PARAMETER siteCollection
        Site collection that contains the solutions
    .EXAMPLE
        Remove-AllSandboxedSolutions -siteCollection https://speval
        Has a dependancy on other functions in this module (Deactivate-SandboxedSolution & Remove-SandboxedSolution)
    #>
    param (
        [string]$siteCollection
    )
            $sc = get-spsite -identity $siteCollection
            $solutions = $sc.Solutions

                if($solutions -ne $null) {
                
                    foreach($solution in $solutions){
                        $status = $solution.status

                        if($status -eq "Activated") {
                            Deactivate-SandboxedSolution -siteCollection $siteCollection -solution $solution.name
                            Remove-SandboxedSolution -siteCollection $siteCollection -solution $solution.name   

                        }
                        else {
                            Remove-SandboxedSolution -siteCollection $siteCollection -solution $solution.name           
                        }
                    }

                }
            
                else {
                    Write-Host "No solutions found in $siteCollection" -BackgroundColor "Red" -ForegroundColor "White"  
                }
}

Function Get-AllSandboxedSolutions {
    <#
    .SYNOPSIS
        Script to produce CSV report of sandboxed solutions in a web application (includes stauts)
    .DESCRIPTION
        Get-ListGuids
    .PARAMETER webApp
        Desired web application
    .PARAMETER outPath
        Path that CSV file will be saved to
    .EXAMPLE
        Get-AllSandboxedSolutions -webApp https://speval -outPath "C:"
    #>
    param (
        [string]$webApp, 
        [string]$outPath
    )
            $date = Get-Date -Format "ddmmyy"
            $name = "_Get-AllSandboxedSolutions.csv"
            $pathStatus = Test-Path -Path $outPath -ErrorAction SilentlyContinue
        
                if($pathStatus -eq $true){
                set-variable -option constant -name out -value "$outPath\$date$name"
                "sep=;" | Out-File $out 
                "solution;parent;url" | Out-File $out -append

                    $scs = Get-SPSite -WebApplication $webApp -Limit all
                    Start-SPAssignment –Global

                        foreach($sc in $scs){
                            $solutions = $sc.Solutions

                            foreach($solution in $solutions){

                                if($solution -ne $null) {
                                    $title = $solution.name
                                    $parent = $sc.RootWeb.Title
                                    $url = $sc.RootWeb.Url
                                    $title + ";" + $parent + ";" + "$url" + ";" | Out-File $out -append  
                                }

                            }
                        }
                            Stop-SPAssignment –Global
                            write-host "Success! get the output here : $out" -BackgroundColor "Green" -ForegroundColor "White" 
                }

                else{
                    write-host "dodgy path!, please check the path and try again" -BackgroundColor "Red" -ForegroundColor "White"  
                }
}

Function Set-GroupToViewAll {
    <#
    .SYNOPSIS
        Script to set permission group visabilty to "all"
    .DESCRIPTION
        Set-GroupToViewAll
    .PARAMETER web
        Website that contains permission group
    .PARAMETER containsName
        Name of group to set permission group visabilty
    .EXAMPLE
        $allWebs = get-spsite -WebApplication https://speval -Limit all | Get-SPWeb -Limit all
            
            foreach($item in $allWebs){
                $url = $item.Url
                Set-GroupToViewAll -web $url -containsName "Members"    
            }
    #>
    param (
        [string]$web, 
        [string]$containsName
    )
            $site = get-spweb -identity $web -ErrorAction SilentlyContinue
        
                if($site -ne $null){
                    $uniquePerms = $site.HasUniqueRoleAssignments
                    $groups = $site.Groups
                
                        if($uniquePerms -eq $true){
                        
                            foreach($group in $groups){
                                $name = $group.name
                                $status = $group.OnlyAllowMembersViewMembership

                                    if($name -like "*$containsName*"){
                                        if($status -eq $true){
                                            $group.OnlyAllowMembersViewMembership = $false
                                            $group.update()
                                            write-host "$site - Success! $name membership view set to everyone" -BackgroundColor "Green" -ForegroundColor "White"
                                        }
                                        else{
                                           write-host "$site - $name membership already set to view for everyone" -BackgroundColor "Yello" -ForegroundColor "Black"  
                                        } 
                                    }
                                    else{
                                        write-host "$site - Ignoring $name" -BackgroundColor "Red" -ForegroundColor "White"     
                                    }
                            }
                        }
                        else{
                            write-host "Ignoring $web, doesn't have unique permissions" -BackgroundColor "Red" -ForegroundColor "White"    
                        }
              }
              else{
                write-host "web $web doesn't exist, please check the url and try again" -BackgroundColor "Red" -ForegroundColor "White"  
              }
}

Function Set-GroupToMemberOnly {
    <#
    .SYNOPSIS
        Script to set permission group visabilty to members only
    .DESCRIPTION
        Set-GroupToMemberOnly
    .PARAMETER web
        Website that contains permission group
    .PARAMETER containsName
        Name of group to set permission group visabilty
    .EXAMPLE
        $allWebs = get-spsite -WebApplication https://speval -Limit all | Get-SPWeb -Limit all
            
            foreach($item in $allWebs){
                $url = $item.Url
                Set-GroupToMemberOnly -web $url -containsName "Members"    
            }
    #>
    param (
        [string]$web, 
        [string]$containsName
    )
            $site = get-spweb -identity $web -ErrorAction SilentlyContinue
        
                if($site -ne $null){
                    $uniquePerms = $site.HasUniqueRoleAssignments
                    $groups = $site.Groups
                
                        if($uniquePerms -eq $true){
                        
                            foreach($group in $groups){
                                $name = $group.name
                                $status = $group.OnlyAllowMembersViewMembership

                                    if($name -like "*$containsName*"){
                                        if($status -eq $false){
                                            $group.OnlyAllowMembersViewMembership = $tue
                                            $group.update()
                                            write-host "$site - Success! $name membership view set to group members" -BackgroundColor "Green" -ForegroundColor "White"
                                        }
                                        else{
                                           write-host "$site - $name membership already set to view for group members" -BackgroundColor "Yello" -ForegroundColor "Black"  
                                        } 
                                    }
                                    else{
                                        write-host "$site - Ignoring $name" -BackgroundColor "Red" -ForegroundColor "White"     
                                    }
                            }
                        }
                        else{
                            write-host "Ignoring $web, doesn't have unique permissions" -BackgroundColor "Red" -ForegroundColor "White"    
                        }
              }
              else{
                write-host "web $web doesn't exist, please check the url and try again" -BackgroundColor "Red" -ForegroundColor "White"  
              }
}

Function Remove-ListPermissionsUser {
    <#
    .SYNOPSIS
        Script to delete unique user permissions from a list
    .DESCRIPTION
        Remove-ListPermissionsUser
    .PARAMETER web
        Website that contains the list
    .PARAMETER library
        Display name of library
    .PARAMETER user
        Display name of user
    .EXAMPLE
        Remove-ListPermissionsUser -web https://speval -library "My Library" -user "Matt Warburton"
    #>
    param (
        [string]$web, 
        [string]$library, 
        [string]$user
    )
            $site = get-spweb -Identity $web
            $id = Get-SPUser -web $web -Identity $user
            $lib = $web.Lists["$library"]
            $lib.RoleAssignments.Remove($id)
            $lib.Update()
}

Function Remove-Website {
    <#
    .SYNOPSIS
        Script to delete a website (site collection or web, script detects this)
    .DESCRIPTION
        Remove-Website
    .PARAMETER website
        Website that should be deleted
    .EXAMPLE
        Remove-Website -website https://speval
    #>
    param (
        [string]$website
    )
            $site = get-spweb -Identity $website -ErrorAction SilentlyContinue
            
                if($site -ne $null){
                    $rootStatus = $site.IsRootWeb
            
                    if($rootStatus -eq $true){
                        Remove-SPSite -Identity $site.Url -Confirm:$false 
                        write-host "Success! Site collection $website deleted" -BackgroundColor "Green" -ForegroundColor "White"     
                    }
                    else{
                        Remove-SPWeb -Identity $site.Url -Confirm:$false 
                        write-host "Success! Web $website deleted" -BackgroundColor "Green" -ForegroundColor "White"  
                    }
                }
                else{
                    write-host "$website doesn't exist, please check the url and try again" -BackgroundColor "Red" -ForegroundColor "White" 
                }
}

Function Remove-ListItem {
    <#
    .SYNOPSIS
        Script to delete item/s from a list
    .DESCRIPTION
        Remove-ListItem
    .PARAMETER website
        Website that contains the list items
    .PARAMETER filterField
        Column filter field. i.e. if you want to delete all items that match a value in a partiular column. This is where you select the column
    .PARAMETER filterVal
        Column filter value. i.e. if you want to delete all items that match a value in a partiular column. This is where you select the value
    .PARAMETER listName
        Display name for the list that contains the items
    .EXAMPLE
        Remove-ListItem -website https://speval -listName theList -filterField Title -filterVal phase
    #>
    param (
        [Parameter(Mandatory=$true)][string]$website, 
        [Parameter(Mandatory=$true)][string]$filterField, 
        [Parameter(Mandatory=$true)][string]$filterVal, 
        [Parameter(Mandatory=$true)][string]$listName
    )
        $site = get-spweb -Identity $website -ErrorAction SilentlyContinue
            
            if($site -ne $null){
                $list = $site.Lists["$listName"]

                    if($list -ne $null){

                        $items = $list.Items | ?{$_["$filterField"] -like "*$filterVal*"}

                            if($items -ne $null){

                                foreach($item in $items){
                                    
                                    $itemName = $item.DisplayName
                                    $item.delete()
                                    Write-Host $itemName deleted -BackgroundColor "Green" -ForegroundColor "White" 
                                }
                            }
                            else{
                                write-host "can't find any items that match your filter parameters $filterField & $filterVal, please check the name and try again" -BackgroundColor "Red" -ForegroundColor "White"    
                            }
                    }
                    else{
                        write-host "List $listName doesn't exist, please check the name and try again" -BackgroundColor "Red" -ForegroundColor "White" 
                    }
            
            }
            else{
                write-host "$website doesn't exist, please check the url and try again" -BackgroundColor "Red" -ForegroundColor "White" 
            }
}

Function Add-SiteCollection {
    <#
    .SYNOPSIS
        Adds a new site collection
    .DESCRIPTION
        Add-SiteCollection
    .PARAMETER url
        URL for new site collection
    .PARAMETER contentDatabase
        Name for content database
    .PARAMETER websiteName
        Display name for site collection
    .PARAMETER primaryLogin
        Primary site collection administrator          
    .EXAMPLE
        Add-SiteCollection -url https://speval -ContentDatabase scripttest5 -WebsiteName scripttest -PrimaryLogin "domain\user" -Template TeamSite
    .NOTES
        More site template codes can be found at http://www.funwithsharepoint.com/sharepoint-2013-site-templates-codes-for-powershell/
    #>
    param (
        [Parameter(Mandatory=$true)][string]$url, 
        [Parameter(Mandatory=$true)][string]$contentDatabase, 
        [Parameter(Mandatory=$true)][string]$websiteName, 
        [Parameter(Mandatory=$true)][string]$primaryLogin,
        [Parameter(Mandatory=$true)]
            [ValidateSet("BasicSearchCenter",
                         "BlankSite",
                         "Blog",
                         "BusinessIntelligenceCenter",
                         "CollaborationPortal",
                         "DocumentCenter",
                         "EnterpriseSearchCenter",
                         "EnterpriseWiki",
                         "PerformancePoint",
                         "ProjectSite",
                         "PublishingPortal",
                         "PublishingSite",
                         "RecordsCenter",
                         "TeamSite",
                         "WikiSite"
                         )][string]$template
    )

            $codes = @{
                "BasicSearchCenter"="SRCHCENTERLITE#0";
                "BlankSite"="STS#1";
                "Blog"="BLOG#0";
                "BusinessIntelligenceCenter"="BICenterSite#0";
                "CollaboratioPortal"="SPSPORTAL#0";
                "DocumentCenter"="BDR#0";
                "EnterpriseSearch Center"="SRCHCEN#0";
                "EnterpriseWiki"="ENTERWIKI#0";
                "PerformancePoint"="PPSMASite#0";
                "ProjectSite"="PROJECTSITE#0";
                "PublishingPortal"="BLANKINTERNETCONTAINER#0";
                "PublishingSite"="CMSPUBLISHING#0";
                "RecordsCenter"="OFFILE#1";
                "TeamSite"="STS#0";
                "WikiSite"="WIKI#0";
            }

            $db = Get-SPContentDatabase -Identity $ContentDatabase -ErrorAction SilentlyContinue
            $existingSC = Get-SPSite -Identity $url -ErrorAction SilentlyContinue
            $code = $codes.$template
            
                if($existingSC -eq $null){ 
                
                    if($db -ne $null){
                                        
                        New-SPSite -Url $url –ContentDatabase $contentDatabase -Name $websiteName -Template "$code" -OwnerAlias $primaryLogin
                        write-host "Add-SiteCollection - Success! $url created" -BackgroundColor "Green" -ForegroundColor "White" 
                    }
                    else{
                        Write-Host "Content database $contentDatabase doesn't exit" -BackgroundColor "Red" -ForegroundColor "White"
                    
                            Do {
                            "[1] Add content DB"
                            "[2] Exit"
                            $Selection = Read-Host "Please select an option"

                                if($selection -eq "1"){
                                    $wa = read-host "please enter the web application that the content DB should be associated with"
                                    Add-ContentDatabase -name $contentDatabase -webapp $wa  

                                }
                                elseif($selection -eq "2"){
                                    exit
                                }
                                else{
                                    write-host "invalid selection"
                                }

                            }
                            Until (($Selection -eq 1) -or ($Selection -eq 2))
                                                   
                        New-SPSite -Url $url –ContentDatabase $contentDatabase -Name $websiteName -Template "$code" -OwnerAlias $primaryLogin
                        write-host "Add-SiteCollection - Success! $url created" -BackgroundColor "Green" -ForegroundColor "White"
                    }
                }
                else{
                    Write-Host "A site collection with the same address already exists, please try again" -BackgroundColor "Red" -ForegroundColor "White"    
                }                               
}

Function Add-ContentDatabase {
    <#
    .SYNOPSIS
        Adds a content database
    .DESCRIPTION
        Add-ContentDatabase
    .PARAMETER webapp
        Web application for content database
    .PARAMETER name
        Name for content database
    .EXAMPLE
        Add-ContentDatabase -Name "blah blah" -WebApplication https://speval
    #>
    param (
        [Parameter(Mandatory=$true)][string]$name, 
        [Parameter(Mandatory=$true)][string]$webapp
    )
        $waTest = Get-SPWebApplication -Identity $webapp -ErrorAction SilentlyContinue
        $dbTest = Get-SPContentDatabase -Identity $name -ErrorAction SilentlyContinue

        if($dbtest -eq $null){

            if($waTest -ne $null){
                New-SPContentDatabase -Name $name -WebApplication $webapp -ErrorAction SilentlyContinue
                write-host "Content database $name created" -BackgroundColor "Green" -ForegroundColor "White"
            }
            else{
                Write-Host "Web app $name not found, please try again" -BackgroundColor "Red" -ForegroundColor "White"    
            }
        }
        else{
            Write-Host "$name already exists" -BackgroundColor "Red" -ForegroundColor "White"     
        }
}

Function Replace-StandardTemplate {
    <#
    .SYNOPSIS
        Replaces the standard template.dotx in librares where found in site collection (for all content types also)
    .DESCRIPTION
        Replace-StandardTemplate
    .PARAMETER sc
        URL for site collection
    .PARAMETER templateFolder
        Path to folder with new template.dotx
    .EXAMPLE
        Replace-StandardTemplate -sc https://speval -templateFolder "C:\Template"
    #>
    param (
        [string]$sc,
        [string]$templateFolder
    )

    $sites = get-spsite -Identity $sc
    $webs = $sites.AllWebs
    $sourceFileFolder = $templateFolder
    $sourceFile = ([System.IO.DirectoryInfo] (Get-Item $sourceFilesFolder)).GetFiles()

         foreach($web in $webs){
            $lists = $web.Lists

                foreach($list in $lists){
                    $title = $list.Title
                
                    if($list.BaseType -eq "DocumentLibrary" -and $list.BaseTemplate -eq "DocumentLibrary" -and $title -notlike "*Master*" -and $title -notlike "*Pages*" -and $title -notlike "*Assets*" -and $title -notlike "*wfsvc*" -and $title -notlike "*Images*" -and $title -notlike "*wfpub*" -and $title -notlike "*style*" -and $title -notlike "*solution*"){
                        $rootForms = $list.RootFolder.SubFolders["forms"]
                        $rootTemplateFile = $rootForms.Files["template.dotx"]
                        $subForms = $rootForms.SubFolders                    
                    
                        if($rootTemplateFile -ne $null){
                            $parent = $rootTemplateFile.ParentFolder
                            $fileStream = ([System.IO.FileInfo] (Get-Item $sourceFile.FullName)).OpenRead()
                            $spFile = $parent.Files.Add($parent.Url + "/" + $sourceFile.Name, [System.IO.Stream]$fileStream, $true)
                            $fileStream.Close();
                            write-host "$sourceFile.Name uploaded to $parent in $web.Url (ROOT)" -ForegroundColor White -BackgroundColor DarkBlue
                        }
                    
                        if($subForms -ne $null){
                        
                            foreach($subForm in $subForms){
                                $subTemplateFile = $subForm.Files["template.dotx"]

                                    if($subTemplateFile -ne $null){
                                        $subFileStream = ([System.IO.FileInfo] (Get-Item $sourceFile.FullName)).OpenRead()
                                        $subParent = $subTemplateFile.ParentFolder
                                        $subSPFile = $subParent.Files.Add($subParent.Url + "/" + $sourceFile.Name, [System.IO.Stream]$subFileStream, $true)
                                        $subFileStream.Close();
                                        write-host "$sourceFile.Name uploaded to $subParent in $web.Url (SUB)" -ForegroundColor White -BackgroundColor DarkCyan
                                    }
                            }    
                        }

                        else{
                             write-host "can't find any libraries with template.dotx" -ForegroundColor White -BackgroundColor Red
                        }
                    }
                }
         }
}

Function Check-FeatureActivated {
    <#
    .SYNOPSIS
        Check to see whether feature is activated
    .DESCRIPTION
        Check-FeatureActivated
    .PARAMETER site
        URL for site collection
    .PARAMETER id
        id of feature
    .EXAMPLE
        Check-FeatureActivated -site https://speval -id 7c637b23-06c4-472d-9a9a-7c175762c5c4
    #>
    param (
        [string]$site,
        [string]$id
    )
        $feature = Get-SPFeature -Site $site -Identity $id -ErrorAction SilentlyContinue 

            if($feature -ne $null){
                $name = $feature.DisplayName
                write-host $name is activated in $site -ForegroundColor White -BackgroundColor DarkCyan                
            }
}

Function Check-FeatureActivatedAllSites {
    <#
    .SYNOPSIS
        Check to see whether feature is activated in all site collections (farm)
    .DESCRIPTION
        Check-FeatureActivated
    .PARAMETER id
        id of feature
    .EXAMPLE
        Check-FeatureActivatedAllSites -id 7c637b23-06c4-472d-9a9a-7c175762c5c4
        
    #>
    param (
        [string]$id
    )
        get-spsite -limit all | ForEach-Object{
            
            Check-FeatureActivated -site $_.Url -id $id
        }
}

Function Unseal-ContentTypeList {
    <#
    .SYNOPSIS
        Unseal content type in list
    .DESCRIPTION
        Unseal-ContentTypeList
    .PARAMETER site
        URL for site
    .PARAMETER listName
        Display name of library
    .PARAMETER contentType
        Content type name
    .EXAMPLE
        Unseal-ContentTypeList -site https://speval -listName customLib -contentType "the content type"
    #>
    param (
        [string]$site,
        [string]$listName,
        [string]$contentType
    )
        $web = Get-SPWeb $site -ErrorAction SilentlyContinue

            if($web -ne $null){
                $list = $web.Lists[$listName]
                
                    if($list -ne $null){
                        $ct = $list.ContentTypes[$contentType]

                            if($ct -ne $null){

                                $status = $ct.ReadOnly

                                    if($status -eq $true){
                                        $ct.ReadOnly = $false
                                        $ct.update()
                                        Write-Host $contentType set to read/write -ForegroundColor White -BackgroundColor DarkCyan
                                    }
                                    else{
                                        Write-Host $contentType is already read/write -BackgroundColor "Red" -ForegroundColor "White" 
                                    }

                            }
                            else{
                                Write-Host $contentType not found -BackgroundColor "Red" -ForegroundColor "White" 
                            }
                    }
                    else{
                        Write-Host $listName not found -BackgroundColor "Red" -ForegroundColor "White"   
                    }               
            }
            else{
                Write-Host site $site not found -BackgroundColor "Red" -ForegroundColor "White" 
            }
}

Function Seal-ContentTypeList {
    <#
    .SYNOPSIS
        Seal content type in list
    .DESCRIPTION
        Seal-ContentTypeList
    .PARAMETER site
        URL for site
    .PARAMETER listName
        Display name of library
    .PARAMETER contentType
        Content type name
    .EXAMPLE
        Seal-ContentTypeList -site https://speval -listName customLib -contentType "the content type"
    #>
    param (
        [string]$site,
        [string]$listName,
        [string]$contentType
    )
        $web = Get-SPWeb $site -ErrorAction SilentlyContinue

            if($web -ne $null){
                $list = $web.Lists[$listName]
                
                    if($list -ne $null){
                        $ct = $list.ContentTypes[$contentType]

                            if($ct -ne $null){

                                $status = $ct.ReadOnly

                                    if($status -eq $false){
                                        $ct.ReadOnly = $true
                                        $ct.update()
                                        Write-Host $contentType sealed -ForegroundColor White -BackgroundColor DarkCyan
                                    }
                                    else{
                                        Write-Host $contentType is already sealed -BackgroundColor "Red" -ForegroundColor "White" 
                                    }

                            }
                            else{
                                Write-Host $contentType not found -BackgroundColor "Red" -ForegroundColor "White" 
                            }
                    }
                    else{
                        Write-Host $listName not found -BackgroundColor "Red" -ForegroundColor "White"   
                    }               
            }
            else{
                Write-Host site $site not found -BackgroundColor "Red" -ForegroundColor "White" 
            }
}

Function Unseal-DeleteContentTypeList {
    <#
    .SYNOPSIS
        Unseals and deletes content type in list
    .DESCRIPTION
        Unseal-DeleteContentTypeList
    .PARAMETER site
        URL for site
    .PARAMETER listName
        Display name of library
    .PARAMETER contentType
        Content type name
    .EXAMPLE
        Unseal-DeleteContentTypeList -site https://speval -listName customLib -contentType "the content type
    .Nodes
        Dependancies on a couple of other functions (Unseal-ContentTypeList & Remove-ContentTypeList)
    #>
    param (
        [string]$listName,
        [string]$site,
        [string]$contentType
    )

    Unseal-ContentTypeList -site $site -listName $listName -contentType $contentType
    Remove-ContentTypeList -site $site -listName $listName -contentType $contentType
}

Function Disable-Feature {
    <#
    .SYNOPSIS
        Disable feature
    .DESCRIPTION
        Disable-Feature
    .PARAMETER site
        URL for site collection
    .PARAMETER id
        id of feature
    .EXAMPLE
        Disable-Feature -site https://speval -id b50e3104-6812-424f-a011-cc90e6327318
    #>
    param (
        [string]$site,
        [string]$id
    )
        $feature = Get-SPFeature -Site $site -Identity $id -ErrorAction SilentlyContinue 

            if($feature -ne $null){
                Disable-SPFeature –Identity $id –url $site -Confirm:$False
                 write-host $id disabled in $site             
            }
}

Function Upload-File {
    <#
    .SYNOPSIS
        Uploads file to SharePoint location
    .DESCRIPTION
        Upload-File
    .PARAMETER web
        URL for SharePoint site
    .PARAMETER fileFolder
        Path to folder that contains file
    .PARAMETER library
        Library name
    .EXAMPLE
        Upload-File -web https://speval -library "List Template Gallery" -fileFolder "C:\Testing"
    #>
    param (
        [string]$web,
        [string]$library,
        [string]$fileFolder
    )

    $sourceFile = ([System.IO.DirectoryInfo] (Get-Item $fileFolder)).GetFiles()
    $site = get-spweb -Identity $web
    $list = $site.lists[$library]
    $root = $list.RootFolder
    $fileStream = ([System.IO.FileInfo] (Get-Item $sourceFile.FullName)).OpenRead()
    $spFile = $root.Files.Add($root.Url + "/" + $sourceFile.Name, [System.IO.Stream]$fileStream, $true)
    $fileStream.Close();
}

Function Set-DefaultColValue {
    <#
    .SYNOPSIS
        Sets the default col value for field in library
    .DESCRIPTION
        Set-DefaultColValue
    .PARAMETER web
        URL for SharePoint site
    .PARAMETER library
        Library name
    .PARAMETER field
        Field name (internal
    .PARAMETER defaultVal
        Desired default val
    .EXAMPLE
        Set-DefaultColValue -web https://speval -library "customLib" -field "Category Tag" -defaultVal "this is a new default val"
    #>
    param (
        [string]$web,
        [string]$library,
        [string]$defaultVal,
        [string]$field
    )

    $site = Get-SPWeb -Identity $web
    $list = $site.lists[$library]
    $columnDefault = New-Object Microsoft.Office.DocumentManagement.MetadataDefaults($list)
    $folder = $list.RootFolder.ServerRelativeUrl
    $columnDefault.SetFieldDefault($folder, $field, $defaultVal) | Out-Null
    $columnDefault.Update()
    Write-Host Default value $defaultVal set -ForegroundColor White -BackgroundColor DarkCyan

}

Function Get-termID {
    <#
    .SYNOPSIS
        Gets the term ID for use in setting column default values
    .DESCRIPTION
        Get-termID
    .PARAMETER webApp
        URL for SharePoint site
    .PARAMETER group
        The term group
    .PARAMETER set
        The term set
    .PARAMETER tag
        The term
    .EXAMPLE
        Get-termID -webApp https://speval -group "Project" -set "Group" -tag "theTag"
    #>
    param (
        [string]$webApp,
        [string]$set,
        [string]$group,
        [string]$tag
    )

    $ts = Get-SPTaxonomySession -Site $webApp
    $tstore = $ts.TermStores[0]
    $tgroup = $tstore.Groups["$group"]
    $tset = $tgroup.TermSets["$set"]
    $term = $tset.Terms["$tag"]
    $format = "1033;#" + $term.Name + "|" + $term.Id
    write-host $format
}

Function Force-Checkout {
    <#
    .SYNOPSIS
        Sets require check out to $true 
    .DESCRIPTION
        Force-Checkout
    .PARAMETER web
        URL for SharePoint site
    .PARAMETER library
    .EXAMPLE
        Force-Checkout -web https://speval -library "TestingAgain"
    #>
    param (
        [string]$web,
        [string]$library
    )
    
    $site = Get-SPWeb -Identity $web
    $list = $site.Lists[$library]
    $status = $list.ForceCheckout

        if($status -eq $false){
            $list.ForceCheckout = $true
            $list.Update()
            Write-Output "Checkout enforced"
        }
        else{ 
            Write-Output "Checkout already enforced"
        }
}

Function Optional-Checkout {
    <#
    .SYNOPSIS
        Sets require check out to $false
    .DESCRIPTION
        Optional-Checkout
    .PARAMETER web
        URL for SharePoint site
    .PARAMETER library
    .EXAMPLE
        Optional-Checkout -web https://speval -library "TestingAgain"
    #>
    param (
        [string]$web,
        [string]$library
    )
    
    $site = Get-SPWeb -Identity $web
    $list = $site.Lists[$library]
    $status = $list.ForceCheckout

        if($status -eq $true){
            $list.ForceCheckout = $false
            $list.Update()
            Write-Host checkout optional
        }
        else{
            Write-Host checkout already optional
        }
}

Function Set-MajorVersion {
    <#
    .SYNOPSIS
        Sets library to major versions
    .DESCRIPTION
        Set-MajorVersion
    .PARAMETER web
        URL for SharePoint site
    .PARAMETER library
        Library name
    .EXAMPLE
        Set-MajorVersion -web https://speval -library "TestingAgain"
    #>
    param (
        [string]$web,
        [string]$library
    )
    
    $site = Get-SPWeb -Identity $web
    $list = $site.Lists[$library]
    $status = $list.EnableVersioning

        if($status -eq $false){
            $list.EnableVersioning = $true
            $list.Update()
            Write-Output "Major versioning enabled"
        }
        else{
            Write-Output "Major versioning already enabled"
        }
}

Function Set-MinorVersion {
    <#
    .SYNOPSIS
        Sets library to minor versions
    .DESCRIPTION
        Set-MinorVersion
    .PARAMETER web
        URL for SharePoint site
    .PARAMETER library
        Library name
    .EXAMPLE
        Set-MinorVersion -web https://speval -library "TestingAgain"
    #>
    param (
        [string]$web,
        [string]$library
    )
    
    $site = Get-SPWeb -Identity $web
    $list = $site.Lists[$library]
    $status = $list.EnableMinorVersions

        if($status -eq $false){
            $list.EnableMinorVersions = $true
            $list.Update()
            Write-Host Minor versioning enabled
        }
        else{
            Write-Host Minor versioning already enabled
        }
}

Function Disable-Versioning {
    <#
    .SYNOPSIS
        Disables versioning in library
    .DESCRIPTION
        Disable-Versioning
    .PARAMETER web
        URL for SharePoint site
    .PARAMETER library
        Library name
    .EXAMPLE
        Disable-Versioning -web https://speval -library "TestingAgain"
    #>
    param (
        [string]$web,
        [string]$library
    )
    
    $site = Get-SPWeb -Identity $web
    $list = $site.Lists[$library]
    $status = $list.EnableVersioning

        if($status -eq $true){
            $list.EnableVersioning = $false
            $list.Update()
            Write-Host Versioning disabled
        }
        else{
            Write-Host Versioning already disabled
        }
}

Function Disable-Folders {
    <#
    .SYNOPSIS
        Disabled use of folders in library
    .DESCRIPTION
        Disable-Folders
    .PARAMETER web
        URL for SharePoint site
    .PARAMETER library
        Library name
    .EXAMPLE
        Disable-Folders -web https://speval -library "TestingAgain"
    #>
    param (
        [string]$web,
        [string]$library
    )
    
    $site = Get-SPWeb -Identity $web
    $list = $site.Lists[$library]
    $status = $list.EnableFolderCreation

        if($status -eq $true){
            $list.EnableFolderCreation = $false
            $list.Update()
            Write-Output "Folders disabled"
        }
        else{
            Write-Output "Folders already disabled"
        }
}

Function Enable-Folders {
    <#
    .SYNOPSIS
        Enables use of folders in library
    .DESCRIPTION
        Enable-Folders
    .PARAMETER web
        URL for SharePoint site
    .PARAMETER library
        Library name
    .EXAMPLE
        Enable-Folders -web https://speval -library "TestingAgain"
    #>
    param (
        [string]$web,
        [string]$library
    )
    
    $site = Get-SPWeb -Identity $web
    $list = $site.Lists[$library]
    $status = $list.EnableFolderCreation

        if($status -eq $false){
            $list.EnableFolderCreation = $true
            $list.Update()
            Write-Host Folders enabled
        }
        else{
            Write-Host Folders already enabled
        }
}

Function Clear-MetaNav {
    <#
    .SYNOPSIS
        Clears all metanav Hierarchy Fields and key filters in a library
    .DESCRIPTION
        Clear-MetaNav
    .PARAMETER web
        URL for SharePoint site
    .PARAMETER library
        Library name
    .EXAMPLE
        Clear-MetaNav -web https://speval -library "customLib"
    #>
    param (
        [string]$web,
        [string]$library
    )
    
    $site = Get-SPWeb -Identity $web
    $list = $site.Lists[$library]
    $listNavSettings = [Microsoft.Office.DocumentManagement.MetadataNavigation.MetadataNavigationSettings]::GetMetadataNavigationSettings($list)
    $status = $listNavSettings.IsEnabled

        if($status -eq $true){
            $listNavSettings.ClearConfiguredHierarchies()
            $listNavSettings.ClearConfiguredKeyFilters()
            [Microsoft.Office.DocumentManagement.MetadataNavigation.MetadataNavigationSettings]::SetMetadataNavigationSettings($list, $listNavSettings, $true)
            $list.RootFolder.Update()
            Write-Host Metanav Hierarchy Fields and key filters cleared
        }
        else{
            Write-Host No  Hierarchy Fields and key filters configured
        }
}

Function Clear-MetaNav {
    <#
    .SYNOPSIS
        Clears all metanav Hierarchy Fields and key filters in a library
    .DESCRIPTION
        Clear-MetaNav
    .PARAMETER web
        URL for SharePoint site
    .PARAMETER library
        Library name
    .EXAMPLE
        Clear-MetaNav -web https://speval -library "customLib"
    #>
    param (
        [string]$web,
        [string]$library
    )
    
    $site = Get-SPWeb -Identity $web
    $list = $site.Lists[$library]
    $listNavSettings = [Microsoft.Office.DocumentManagement.MetadataNavigation.MetadataNavigationSettings]::GetMetadataNavigationSettings($list)
    $status = $listNavSettings.IsEnabled

        if($status -eq $true){
            $listNavSettings.ClearConfiguredHierarchies()
            $listNavSettings.ClearConfiguredKeyFilters()
            [Microsoft.Office.DocumentManagement.MetadataNavigation.MetadataNavigationSettings]::SetMetadataNavigationSettings($list, $listNavSettings, $true)
            $list.RootFolder.Update()
            Write-Host Metanav Hierarchy Fields and key filters cleared
        }
        else{
            Write-Host No  Hierarchy Fields and key filters configured
        }
}

Function Add-MetaNavKeyFilter {
    <#
    .SYNOPSIS
        Adds key filter to metadata navigation in library
    .DESCRIPTION
        Add-MetaNavKeyFilter
    .PARAMETER web
        URL for SharePoint site
    .PARAMETER library
        Library name
    .PARAMETER field
        Field name
    .EXAMPLE
        Add-MetaNavKeyFilter -web https://speval -library "customLib" -field "Created By"
    #>
    param (
        [string]$web,
        [string]$field,
        [string]$library
    )
    
    $feature = Get-SPFeature -Web $web -Identity 7201d6a4-a5d3-49a1-8c19-19c4bac6e668 -ErrorAction SilentlyContinue
    $site = Get-SPWeb -Identity $web
    $list = $site.Lists[$library]

        if($feature -eq $null){
            Enable-SPFeature –Identity 7201d6a4-a5d3-49a1-8c19-19c4bac6e668 –url $web  
            $listNavSettings = [Microsoft.Office.DocumentManagement.MetadataNavigation.MetadataNavigationSettings]::GetMetadataNavigationSettings($list)
            $listNavSettings.AddConfiguredKeyFilter($list.Fields[$field])
            [Microsoft.Office.DocumentManagement.MetadataNavigation.MetadataNavigationSettings]::SetMetadataNavigationSettings($list, $listNavSettings, $true)
            $list.RootFolder.Update()
            Write-Host Meta nav feature activated
            Write-Host Key filter $field added
        }
        if($feature -ne $null){
            $listNavSettings = [Microsoft.Office.DocumentManagement.MetadataNavigation.MetadataNavigationSettings]::GetMetadataNavigationSettings($list)
            $listNavSettings.AddConfiguredKeyFilter($list.Fields[$field])
            [Microsoft.Office.DocumentManagement.MetadataNavigation.MetadataNavigationSettings]::SetMetadataNavigationSettings($list, $listNavSettings, $true)
            $list.RootFolder.Update()
            Write-Host Meta nav feature already activated
            Write-Host Key filter $field added
        }
}

Function Add-MetaHierarchyFilter {
    <#
    .SYNOPSIS
        Adds Hierarchy filter to metadata navigation in library
    .DESCRIPTION
        Add-MetaHierarchyFilter
    .PARAMETER web
        URL for SharePoint site
    .PARAMETER library
        Library name
    .PARAMETER field
        Field name
    .EXAMPLE
        Add-MetaHierarchyFilter -web https://speval -library "customLib" -field "Category Tag"
    #>
    param (
        [string]$web,
        [string]$field,
        [string]$library
    )
    
    $feature = Get-SPFeature -Web $web -Identity 7201d6a4-a5d3-49a1-8c19-19c4bac6e668 -ErrorAction SilentlyContinue
    $site = Get-SPWeb -Identity $web
    $list = $site.Lists[$library]

        if($feature -eq $null){
            Enable-SPFeature –Identity 7201d6a4-a5d3-49a1-8c19-19c4bac6e668 –url $web  
            $listNavSettings = [Microsoft.Office.DocumentManagement.MetadataNavigation.MetadataNavigationSettings]::GetMetadataNavigationSettings($list)
            $listNavSettings.AddConfiguredHierarchy($list.Fields[$field])
            [Microsoft.Office.DocumentManagement.MetadataNavigation.MetadataNavigationSettings]::SetMetadataNavigationSettings($list, $listNavSettings, $true)
            $list.RootFolder.Update()
            Write-Host Meta nav feature activated
            Write-Host Hierarchy Filter $field added
        }
        if($feature -ne $null){
            $listNavSettings = [Microsoft.Office.DocumentManagement.MetadataNavigation.MetadataNavigationSettings]::GetMetadataNavigationSettings($list)
            $listNavSettings.AddConfiguredHierarchy($list.Fields[$field])
            [Microsoft.Office.DocumentManagement.MetadataNavigation.MetadataNavigationSettings]::SetMetadataNavigationSettings($list, $listNavSettings, $true)
            $list.RootFolder.Update()
            Write-Host Meta nav feature already activated
            Write-Host Hierarchy Filter $field added
        }
}

Function Add-LibraryQuickLaunch {
    <#
    .SYNOPSIS
        Adds a library to the quick launch nav
    .DESCRIPTION
        Add-LibraryQuickLaunch
    .PARAMETER web
        URL for SharePoint site
    .PARAMETER library
        Library name
    .EXAMPLE
        Add-LibraryQuickLaunch -web https://speval -library "TestingAgain"
    #>
    param (
        [string]$web,
        [string]$library
    )
    
    $site = Get-SPWeb -Identity $web
    $list = $site.Lists[$library]
    $status = $list.OnQuickLaunch

        if($status -eq $false){
            $list.OnQuickLaunch = $true
            $list.Update()
            Write-Output "Library added to quick launch"
        }
        else{
            Write-Output "Library already on quick launch"
        }
}

Function Find-CT {
    <#
    .SYNOPSIS
        Script to find libraries where a particular content type is associated
    .DESCRIPTION
        Find-CT
    .PARAMETER webApp
        Web application to search
    .PARAMETER contentType
        Name of content types to find
    .EXAMPLE
        Find-CT -webApp Find-CT -webApp https://speval -contentType Template
    #>
    param (
        [string]$webApp, 
        [string]$contentType
    )
            $webs = Get-SPSite -WebApplication $webApp -Limit all | Get-SPWeb -Limit All 

                foreach($web in $webs){
                    Start-SPAssignment -Global
                    $url = $web.Url
                    $lists = $web.Lists
                    
                        foreach($list in $lists){
                            $cts = $list.ContentTypes
                            
                                foreach($ct in $cts){

                                    if($ct.name -eq $contentType){
                                        $lurl = $list.RootFolder.ServerRelativeUrl
                                        $link = "$url$lurl"
                                        write-host $contentType found in $list - $link -ForegroundColor White -BackgroundColor DarkCyan
                                    }                                   
                                }    
                        }                
                    Stop-SPAssignment -Global
                }
}

Function Get-RecentLibsCSV {
    <#
    .SYNOPSIS
        Script to produce CSV report of all libraries in a site collection
    .DESCRIPTION
        Get-AllLibraries
    .PARAMETER siteCollection
        Desired site collection
    .PARAMETER outPath
        Path that CSV file will be saved to
    .EXAMPLE
        $todayFormat = Get-Date -format "d MMM yyyy"
        $file = "libreport_$todayFormat.csv"
        $path = "C:\LibraryReport\$file"
        $spLibrary = "https://speval/Templates"
        $destinationPath = "$spLibrary/$file"
        Get-RecentLibsCSV -webApp https://speval -outPath $path -daysOld 7

        Invoke-WebRequest -Uri $destinationPath -InFile $path -Method PUT -UseDefaultCredentials 
    #>   
    param (
        [string]$webApp, 
        [string]$outPath, 
        [string]$daysOld
    )
        $name = "_libteport.csv"
        $today = Get-Date
        set-variable -option constant -name out -value "$outPath"
        "sep=;" | Out-File $out 
        "library;libraryurl;parent;parenturl;created;createdby" | Out-File $out -append
        $webs = Get-SPSite -WebApplication $webApp -Limit all | Get-SPWeb -Limit All
            
            foreach($web in $webs){
                $lists = $web.GetListsOfType("DocumentLibrary")

                    foreach($list in $lists){
                        $gap = $today - $list.Created
                        $author = $list.Author.DisplayName
                        $title = $list.Title

                            if($author -ne "System Account" -and $author -ne "sp_setup" -and $gap.Days -lt $daysOld -and $title -notlike "*Master*" -and $title -notlike "*Pages*" -and $title -notlike "*Assets*" -and $title -notlike "*wfsvc*" -and $title -notlike "*Images*" -and $title -notlike "*wfpub*" -and $title -notlike "*style*" -and $title -notlike "*solution*"){
                                $parent = $list.ParentWeb.Title
                                $parentUrl = $list.ParentWeb.Url
                                $libraryRelUrl = $list.RootFolder.ServerRelativeUrl
                                $libraryUrl = "$parentUrl$libraryRelUrl"
                                $created = $list.Created
                                $createdFormat = $created.ToString("d MMM yyyy")
                                "$title" + ";" + "$libraryUrl" + ";" + "$parent" + ";" + "$parentUrl" + ";" + "$createdFormat" + ";" + "$author" + ";" | Out-File $out -append
                            }
                    }

            }
}

Function Break-InheritanceList {
    <#
    .SYNOPSIS
        Breaks permission inheritance
    .DESCRIPTION
        Break-InheritanceList
    .PARAMETER web
        URL for SharePoint site
    .PARAMETER library
        Library name
    .PARAMETER option
        $true to copy exiting perms, $false to clear exiting perms
    .EXAMPLE
        Break-InheritanceList -web https://speval -library "Test" -option $true
    #>
    param (
        [string]$web,
        [string]$library,
        [bool]$option
    )
    
    $site = Get-SPWeb -Identity $web
    $list = $site.Lists[$library]
    $status = $list.HasUniqueRoleAssignments

        if($status -eq $false){
            $list.BreakRoleInheritance($option)
            $list.Update()
            Write-Output "Permission inheritance broken"
        }
        else{
            Write-Output "Permission inheritance already broken"
        }
}

Function Reinherit-PermissionsList {
    <#
    .SYNOPSIS
        Reinherits permissions
    .DESCRIPTION
        Reinherit-PermissionsList
    .PARAMETER web
        URL for SharePoint site
    .PARAMETER library
        Library name
    .EXAMPLE
        Reinherit-PermissionsList -web https://speval -library "Test"
    #>
    param (
        [string]$web,
        [string]$library
    )
    
    $site = Get-SPWeb -Identity $web
    $list = $site.Lists[$library]
    $status = $list.HasUniqueRoleAssignments

        if($status -eq $true){
            $list.ResetRoleInheritance()
            $list.Update()
            Write-Host Permission reinherited
        }
        else{
            Write-Host Permission already inherited
        }
}

Function Add-RootFolder {
    <#
    .SYNOPSIS
        Reinherits permissions
    .DESCRIPTION
        Reinherit-PermissionsList
    .PARAMETER web
        URL for SharePoint site
    .PARAMETER library
        Library name
    .EXAMPLE
        Add-RootFolder -web https://speval -library "Test" -folderName "Blah Blah"
    #>
    param (
        [string]$web,
        [string]$library,
        [string]$folderName
    )
    
    $site = Get-SPWeb -Identity $web
    $list = $site.Lists[$library]
    $folder = $list.Folders.Add("", [Microsoft.SharePoint.SPFileSystemObjectType]::Folder, $folderName)
    $folder.Update()
    Write-Output "$folderName added"
}

Function Break-InheritanceFolder {
    <#
    .SYNOPSIS
        Breaks permission on folder 
    .DESCRIPTION
        Break-InheritanceFolder
    .PARAMETER web
        URL for SharePoint site
    .PARAMETER library
        Library name
    .PARAMETER folderUrl
        URL for folder. Use the server relative URL
    .PARAMETER option
        $true to copy exiting perms, $false to clear exiting perms
    .EXAMPLE
        Break-InheritanceFolder -web https://speval -library "Test" -folderUrl "Test/Blah Blah" -option $false
    #>
    param (
        [string]$web,
        [string]$library,
        [bool]$option,
        [string]$folderUrl
    )
    
    $site = Get-SPWeb -Identity $web
    $list = $site.Lists[$library]
    $folders = $list.Folders

        foreach($folder in $folders){
            $path = $folder.url
            
                if($path -eq "$folderUrl"){

                    write-host $folder
                    $folder.BreakRoleInheritance($option)
                    $folder.update()
                    Write-Output "Permission inheritance broken"
                }
        }
}

Function Reinherit-PermissionsFolder {
    <#
    .SYNOPSIS
        Reinherits permission on folder 
    .DESCRIPTION
        Reinherit-PermissionsFolder
    .PARAMETER web
        URL for SharePoint site
    .PARAMETER library
        Library name
    .PARAMETER folderUrl
        URL for folder. Use the server relative URL
    .EXAMPLE
        Reinherit-PermissionsFolder -web https://speval -library "Test" -folderUrl "Test/Blah Blah" -option "$false"
    #>
    param (
        [string]$web,
        [string]$library,
        [string]$folderUrl
    )
    
    $site = Get-SPWeb -Identity $web
    $list = $site.Lists[$library]
    $folders = $list.Folders

        foreach($folder in $folders){
            $path = $folder.url
            
                if($path -eq "$folderUrl"){

                    write-host $folder
                    $folder.ResetRoleInheritance()
                    $folder.update()
                    Write-Host Permission inheritance broken -ForegroundColor White -BackgroundColor DarkCyan
                }
        }
}

Function Grant-PermissionsFolder {
    <#
    .SYNOPSIS
        Grants permissions for particular user/ad group 
    .DESCRIPTION
        Grant-PermissionsFolder
    .PARAMETER web
        URL for SharePoint site
    .PARAMETER library
        Library name
    .PARAMETER folderUrl
        URL for folder. Use the server relative URL
    .PARAMETER user
        User to give access to (if security group need the claims ID e.g c:0+.w|1234")
    .PARAMETER role
        Role to give user (read, contribute etc.)
    .EXAMPLE
        Grant-PermissionsFolder -web https://speval -library "Test" -folderUrl "Test/Blah Blah" -user "c:0+.w|1234" -role "Contribute"
    #>
    param (
        [string]$web,
        [string]$library,
        [string]$folderUrl,
        [string]$user,
        [string]$role
    )
    
    $site = Get-SPWeb -Identity $web
    $list = $site.Lists[$library]
    $folders = $list.Folders
    $site.allusers.Add("$user", "", "", "")
    $spUser = $site.AllUsers[$user]
    $roleDefinition = $site.RoleDefinitions[$role]
    $roleAssignment = New-Object Microsoft.SharePoint.SPRoleAssignment($spUser)
    $roleAssignment.RoleDefinitionBindings.Add($roleDefinition)

        foreach($folder in $folders){
            $path = $folder.url
            
                if($path -eq "$folderUrl"){
                    $folder.RoleAssignments.Add($roleAssignment)
                    $folder.Update()
                    write-output "permissions added"
                }
        }
}

Function Grant-PermissionsList {
    <#
    .SYNOPSIS
        Grants permissions for particular user/ad group 
    .DESCRIPTION
        Grant-PermissionsList
    .PARAMETER web
        URL for SharePoint site
    .PARAMETER library
        Library name
    .PARAMETER user
        User to give access to (if security group need the claims ID e.g c:0+.w|1234")
    .PARAMETER role
        Role to give user (read, contribute etc.)
    .EXAMPLE
        Grant-PermissionsList -web https://speval/sites/mer/PRN6321 -library "Testing" -user "c:0+.w|1234" -role "Contribute"
    #>
    param (
        [string]$web,
        [string]$library,
        [string]$user,
        [string]$role
    )
    
    $site = Get-SPWeb -Identity $web
    $list = $site.Lists[$library]
    $site.allusers.Add("$user", "", "", "")
    $spUser = $site.AllUsers[$user]
    $roleDefinition = $site.RoleDefinitions[$role]
    $roleAssignment = New-Object Microsoft.SharePoint.SPRoleAssignment($spUser)
    $roleAssignment.RoleDefinitionBindings.Add($roleDefinition)
    $list.RoleAssignments.Add($roleAssignment)
    $list.Update()
    write-output "permissions added"   
}

Function Activate-Feature {
    <#
    .SYNOPSIS
        CActivates feature
    .DESCRIPTION
        Activate-Feature
    .PARAMETER id
        id of feature
    .PARAMETER site
        site to activate feature in
    .EXAMPLE
        Activate-Feature -site https://speval -id b50e3104-6812-424f-a011-cc90e6327318
        
    #>
    param (
        [string]$id,
        [string]$site
    )
        $feature = Get-SPFeature -Site $site -Identity $id -ErrorAction SilentlyContinue 

            if($feature -eq $null){
                Enable-SPFeature –Identity $id –url $site
                write-host $id activated in $site -ForegroundColor White -BackgroundColor DarkCyan                
            }
            else{
                write-host $id already activated in $site -ForegroundColor White -BackgroundColor DarkCyan   
            }
}

Function Disable-Feature {
    <#
    .SYNOPSIS
        Disables feature
    .DESCRIPTION
        Disable-Feature
    .PARAMETER id
        id of feature
    .PARAMETER site
        site to disable feature in
    .EXAMPLE
        Disable-Feature -site https://speval -id b50e3104-6812-424f-a011-cc90e6327318
        
    #>
    param (
        [string]$id,
        [string]$site
    )
        $feature = Get-SPFeature -Site $site -Identity $id -ErrorAction SilentlyContinue 

            if($feature -ne $null){
                Disable-SPFeature –Identity $id –url $site -Confirm:$False
                write-host $id disabled in $site -ForegroundColor White -BackgroundColor DarkCyan                
            }
            else{
                write-host $id not activated in $site -ForegroundColor White -BackgroundColor DarkCyan   
            }
}

Function Activate-FeatureBulk {
    <#
    .SYNOPSIS
        CActivates feature
    .DESCRIPTION
        Activate-FeatureBulk
    .PARAMETER id
        id of feature
    .PARAMETER site
        site to activate feature in
    .PARAMETER path
        path to source csv
    .EXAMPLE
        Activate-FeatureBulk -path "C:\sites.csv" -id b50e3104-6812-424f-a011-cc90e6327318
        
    #>
    param (
        [string]$id,
        [string]$path
    )
        $sites = Import-Csv -Path $path

            foreach($site in $sites){
                $feature = Get-SPFeature -Site $site.url -Identity $id -ErrorAction SilentlyContinue 

                    if($feature -eq $null){
                        Enable-SPFeature –Identity $id –url $site.url
                        write-host $id activated in $site -ForegroundColor White -BackgroundColor DarkCyan                
                    }
                    else{
                        write-host $id already activated in $site.url -ForegroundColor White -BackgroundColor DarkCyan   
                    }
            }
}

Function Disable-FeatureBulk {
    <#
    .SYNOPSIS
        Disabled feature
    .DESCRIPTION
        Disable-FeatureBulk
    .PARAMETER id
        id of feature
    .PARAMETER site
        site to activate feature in
    .PARAMETER path
        path to source csv
    .EXAMPLE
        Disable-FeatureBulk -path "C:\sites.csv" -id b50e3104-6812-424f-a011-cc90e6327318
        
    #>
    param (
        [string]$id,
        [string]$path
    )
        $sites = Import-Csv -Path $path

            foreach($site in $sites){
                $feature = Get-SPFeature -Site $site.url -Identity $id -ErrorAction SilentlyContinue 

                    if($feature -ne $null){
                        Disable-SPFeature –Identity $id –url $site.url -Confirm:$False
                        write-host $id disabled in $site -ForegroundColor White -BackgroundColor DarkCyan              
                    }
                    else{
                        write-host $id not activated in $site -ForegroundColor White -BackgroundColor DarkCyan    
                    }
            }
}

Function Start-TimerJob {
    <#
    .SYNOPSIS
        Script to start a timer job
    .DESCRIPTION
        Find-CT
    .PARAMETER jobName
        Name of timer job
    .EXAMPLE
        Start-TimerJob -jobName DocIdAssignment
    #>
    param (
        [string]$jobName
    )
    $date = Get-Date
    Start-SPTimerJob $jobName

        do{
            $job = Get-SPTimerJob $jobName
            $lastRunTime = $job.LastRunTime
            $rdate = Get-Date
            write-host $jobName timer running $rdate -ForegroundColor White -BackgroundColor DarkCyan
            Start-Sleep -s 2 

        }
            until ($lastRunTime -ge $date)
            write-host $jobName timer finished running $rdate -ForegroundColor White -BackgroundColor Green
}

Function Assign-DocIDPrefix {
    <#
    .SYNOPSIS
        Assign a custom docID prefix
    .DESCRIPTION
        Assign-DocIDPrefix
    .PARAMETER website
        Site collection to set prefix
    .PARAMETER prefix
        Desired prefix
    .EXAMPLE
        Assign-DocIDPrefix -website https://speval -prefix "50000_"
    #>
    param (
        [string]$website, 
        [string]$prefix
    )

    $site = Get-SPSite -Identity $website
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.DocumentManagement")
    [Microsoft.Office.DocumentManagement.DocumentId]::EnableAssignment($site,$prefix,$true,$true,$true,$false)
    write-host $prefix added to $website -BackgroundColor "Green" -ForegroundColor "White"
}

Function Refreash-ContentTypes {
    <#
    .SYNOPSIS
        Refreash all content type in a site collection 
    .DESCRIPTION
        Refreash-ContentTypes
    .PARAMETER website
        Site collection to refreash content types in
    .EXAMPLE
        Refreash-ContentTypes -website "https://speval"
    #>
    param (
        [string]$website
    )

    $site = Get-SPSite -Identity $website
    $root = $site.RootWeb
    $root.Properties["MetadataTimeStamp"] = [string]::Empty
    $root.Properties.Update()
    write-host Content types refreashed in $website -BackgroundColor "Green" -ForegroundColor "White"
}

Function Find-CTsWithField {
    <#
    .SYNOPSIS
        Find all content types in an SC with a particular field 
    .DESCRIPTION
        Find-CTsWithField
    .PARAMETER website
        Site collection
    .PARAMETER fieldName
        Field name you're looking for (use internal)
    .EXAMPLE
        Find-CTsWithField -website "https://speval" -fieldName "Author"
    #>
    param (
        [string]$website,
        [string]$fieldName
    )

    $web = Get-SPWeb $website
    $cts = $web.ContentTypes

        foreach($ct in $cts){
            $fields = $ct.fields
            $ctName = $ct.Name

                foreach($field in $fields){
                    $internal = $field.InternalName
        
                    if($internal -eq $fieldName){
                        write-output "$internal found in $ctName"
                    }
                }
        }
}

Function Find-EmailLibraries {
    <#
    .SYNOPSIS
        Find all emailed enabled libraries
    .DESCRIPTION
        Find-EmailLibraries
    .PARAMETER website
        website
    .EXAMPLE
        Find-EmailLibraries -website "https://speval"
    #>
    param (
        [string]$website
    )

    $web = Get-SPWeb $website
    $lists = $web.Lists

        foreach($list in $lists){
            $status = $list.EmailAlias

                if($status -ne $null){
                    Write-Output "$list email enabled in site $website"
                }  
        }
}

Function Find-EmailLibrariesinWA {
    <#
    .SYNOPSIS
        Find all emailed enabled libraries in WA
    .DESCRIPTION
        Find-EmailLibraries
    .PARAMETER webApp
        website
    .EXAMPLE
        Find-EmailLibrariesinWA -webApp "https://speval"
    #>
    param (
        [string]$webApp
    )

    $webs = Get-SPSite -WebApplication $webApp -limit all | Get-SPWeb -Limit all

        foreach($web in $webs){
            Find-EmailLibraries -website $web.Url
        }
}

Function Find-SCAs {
    <#
    .SYNOPSIS
        Find users in groups that have been assigned a certain permission level
    .DESCRIPTION
        Find-SCAs
    .PARAMETER webApp
        Web application
    .PARAMETER outPath
        Path that CSV file will be saved to
    .EXAMPLE
        Find-SCAs -webApp "https://speval" -outPath "D:\testSCA.csv"
    #>
    param (
        [string]$webApp,
        [string]$outPath
    )
    set-variable -option constant -name out -value "$outPath"
    "sep=;" | Out-File $out 
    "User;Site;SiteUrl" | Out-File $out -append
    $webs = Get-spsite -WebApplication $webApp -Limit all | Get-SPWeb -Limit all

        foreach($web in $webs){
            $webUrl = $web.url
            $statusPerm = $web.HasUniqueRoleAssignments

                if($statusPerm -eq $true){
                    $groups = $web.Groups

                        foreach($group in $groups){
                            $statusGroupPerms = $group.roles
                    
                                if($statusGroupPerms -like "*Owner*" -or $statusGroupPerms -like "*Full Control*"){
                                    $users = $group.Users.displayname  
                            
                                        foreach($user in $users){
                                            Write-Output "$user - $group - $web"
                                            "$user" + ";" + "$web" + ";" + "$webUrl" + ";" | Out-File $out -append    
                                        }
                                }    
                        }
                }
        }
}

Function Undeclare-RecordAllItems {
    <#
    .SYNOPSIS
        Un-declare all items in a list
    .DESCRIPTION
        Get-ColStatus
    .PARAMETER webSite
        Website that contains the list
    .PARAMETER listName
        Name of list
    .EXAMPLE
        Undeclare-RecordAllItems -webSite https://speval -listName "Example List"
    #> 
    param (
        [string]$webSite, 
        [string]$listName
    )
    $web = Get-SPWeb $webSite
    $list = $web.lists[$listName].items
 
        foreach ($item in $list){
             $IsRecord = [Microsoft.Office.RecordsManagement.RecordsRepository.Records]::IsRecord($Item)
             if ($IsRecord -eq $true){
                     Write-Host "Undeclared $($item.Name)"
                     [Microsoft.Office.RecordsManagement.RecordsRepository.Records]::UndeclareItemAsRecord($Item)
             }
        }
}

Function Find-SCAsHTML {
    <#
    .SYNOPSIS
        Blah
    .DESCRIPTION
        TEMPLATE
    .PARAMETER $webApp
        Web application
    .EXAMPLE
        $sourcePath = "D:\Scheduled scripts\siteAdminReport.txt"
        $spLibrary = "https://speval/AdminReports"
        $destinationFile = "siteAdminReport.txt"
        $destinationPath = "$spLibrary/$destinationFile"

        Find-SCAsHTML -webApp https://speval  | Out-File $sourcePath

        Invoke-WebRequest -Uri $destinationPath -InFile $sourcePath -Method PUT -UseDefaultCredentials 
    #>
    [CmdletBinding()] 
    param (
        [string]$webApp
    )
    process{
        $today = Get-Date -format "d MMM yyyy"
        $webs = Get-spsite -WebApplication $webApp -Limit all -ErrorAction SilentlyContinue | Get-SPWeb -Limit all | Where-Object {$_.HasUniqueRoleAssignments -eq $true}
    
            if($webs){
                
                $webs | ForEach-Object{
                    $url = $_.url
                    $webName = $_.title
                    $parentUrl = $_.Site.RootWeb.Url
                    $parentName = $_.Site.RootWeb.title
                    $groups = $_.Groups | Where-Object {$_.roles -like "*permlevel*" -or $_.roles -like "*otherpermlevel*"}

                        $groups | ForEach-Object{
                            $users = $_.Users

                                if($users -ne $null){
                                    $userLI = @()
                                    $groupName = $_.name
                            
                                        $users | ForEach-Object{
                                            $userName = $_.displayName
                                            $userLI += "<li>$userName</li>"
                                        }
                                    $groupTR += "
                                                <tr>
                                                    <td>
                                                        <a href='$url'>$webName</a>    
                                                    </td>
                                                    <td>
                                                        <a href='$parentUrl'>$parentName</a>    
                                                    </td>
                                                    <td>
                                                        $groupName       
                                                    </td>
                                                    <td>
                                                        <ul>
                                                           $userLI
                                                        </ul>
                                                    </td>
                                        </tr>"
                                }
                        }
                }
            }

            else{
                Write-Output "web app $webApp doesn't exist, please try again"
            }
    
        $body = "
                <div>
                    <h4>Last updated - $today</h4>
                </div>
                <div>
                    <table id='cdReportTab' class='display'>
                        <thead>
                            <tr>
                                <td>
                                    Site
                                </td>
                                <td>
                                    Site collection
                                </td>
                                <td>
                                    Permission group name
                                </td>
                                <td>
                                    Site administrators
                                </td>
                            </tr>
                        </thead>
                        <tbody>
                            $groupTR
                        </tbody>
                    </table>
                </div>"
        $body   
    }
} 

Function Get-Item {
    <#
    .SYNOPSIS
        Gets a SharePoint item based on it's ID
    .DESCRIPTION
        TEMPLATE
    .PARAMETER site
        Website
    .PARAMETER list
        List name
    .PARAMETER ID
        Item ID
    .EXAMPLE
        Get-Item -site https://speval -list "content" -ID "1"
    #>
    [CmdletBinding()] 
    param (
        [string]$site, 
        [string]$ID,
        [string]$list
    )
      
    BEGIN {
        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $web = Get-SPWeb $site
            $lst = $web.Lists[$list]

                if($lst) {
                    try {
                        $global:item = $lst.GetItemById("$ID")
                        Write-Output "Got item - $item"                    
                    }
        
                    catch {
                        $error = $_
                        Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"                   
                    }
                }

                else {
                    Write-Output "list $list doesnt exist in $site"                
                }
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }
} 

Function Set-BoolFieldValue {
    <#
    .SYNOPSIS
        Sets a bool field to true or false
    .DESCRIPTION
        Set-BoolFieldValue
    .PARAMETER site
        Website
    .PARAMETER list
        List name
    .PARAMETER ID
        Item ID
    .PARAMETER fieldName
        Field name
    .PARAMETER set
        Specify either $true or $false
    .EXAMPLE
        Set-BoolFieldValue -site https://speval -list "BoolTest" -ID "1" -fieldName "BoolField" -set $true
    .NOTES
        Uses the get-item function
    #>
    [CmdletBinding()] 
    param (
        [string]$site, 
        [string]$ID,
        [string]$list,
        [string]$fieldName, 
        [bool]$set
    )
      
    BEGIN {
        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            Get-Item -site $site -list $list -ID $ID
            $field = $item[$fieldName]

                if($field -ne $null){
                    $item[$fieldName] = $set
                    $item.Update()
                    Write-Output "$item set to $set"  
                }

                else{
                    Write-Output "Can't find field - $fieldName"
                } 
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }
}

Function Get-SCSize {
    <#
    .SYNOPSIS
        Gets size of site collection in MB
    .DESCRIPTION
        Get-SCSize
    .PARAMETER site
        Site collection
    .EXAMPLE
        Get-SCSize -site https://speval
    #>
    [CmdletBinding()] 
    param (
        [string]$site
    )
      
    BEGIN {
        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $sc = Get-SPSite -Identity $site
            $global:usage = $sc.usage.storage/1MB
            Write-Output "$usage"
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }
}

Function Get-SCSizeHTML {
    <#
    .SYNOPSIS
        Gets size of all site collections in a particular web application created HTML table
    .DESCRIPTION
        Get-SCSizeHTML
    .PARAMETER webApp
        Web application
    .EXAMPLE
        $sourcePath = "D:\scSizeReport.txt"
        $spLibrary = "https://speval/AdminReports"
        $destinationFile = "scSizeReport.txt"
        $destinationPath = "$spLibrary/$destinationFile"

        Get-SCSizeHTML -webApp https://speval | Out-File $sourcePath

        Invoke-WebRequest -Uri $destinationPath -InFile $sourcePath -Method PUT -UseDefaultCredentials
    .NOTES
        Uses the Get-SCSize & the Get-SCItemCount functions
    #>
    [CmdletBinding()] 
    param (
        [string]$webApp
    )
      
    BEGIN {
        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {
        $today = Get-Date -format "d MMM yyyy"
        $sites = get-spsite -WebApplication $webApp -Limit all  | Sort-Object -Property "url"

            $sites | ForEach-Object{
                $count = @()
                $size = Get-SCSize -site $_.Url
                $count = Get-SCItemCount -site $_.Url
                $mb = [Math]::Ceiling([decimal]($size))
                $url = $_.Url
                $title = $_.RootWeb.Title
                # Write-Output "$title $url $mb"
                $tr += "<tr>
                            <td>
                                <a href='$url'>$title</a>    
                            </td>
                            <td>
                                $mb 
                            </td>
                            <td>
                                $count 
                            </td>
                        </tr>"
            }

        $body = "
        <div>
            <h4>Last updated - $today</h4>
        </div>
        <div>
            <table id='cdReportTab' class='display'>
                <thead>
                    <tr>
                        <td>
                            Site Collection
                        </td>
                        <td>
                            Size (in MB)
                        </td>
                        <td>
                            Item count
                        </td>
                    </tr>
                </thead>
                <tbody>
                    $tr
                </tbody>
            </table>
        </div>"
        $body           
    }
}

Function Get-SCItemCount {
    <#
    .SYNOPSIS
        Gets total number of list items in a site collection
    .DESCRIPTION
        Get-SCItemCount
    .PARAMETER site
        Site collection
    .EXAMPLE
        Get-SCItemCount -site https://speval
    #>
    [CmdletBinding()] 
    param (
        [string]$site
    )
      
    BEGIN {
        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $webs = Get-spsite -Identity $site | Get-SPWeb -Limit all

                $webs | ForEach-Object{
                    $lists = $_.lists

                        foreach($list in $lists){
                            $totalCount += $list.ItemCount
                        }
                }
            Write-Output $totalCount
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }
} 

Function Get-List {
    <#
    .SYNOPSIS
        Gets a list
    .DESCRIPTION
        Get-List
    .PARAMETER webUrl
        Url of web
    .PARAMETER listName
        Name of list
    .EXAMPLE
        Get-List -webUrl https://speval -listName "BoolTest"
    #>
    [CmdletBinding()] 
    param (
        [string]$webUrl, 
        [string]$listName
    )
      
    BEGIN {
        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $web = Get-SPWeb $webUrl
            $list = $web.Lists[$listName]

                if($list) {
                    Write-Output $list
                }

                else {
                    Write-Output "list '$listName' doesn't exist in web '$webUrl'"                
                }
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }
} 

Function Get-Group {
    <#
    .SYNOPSIS
        Gets a SharePoint permission group
    .DESCRIPTION
        Get-Group
    .PARAMETER webUrl
        Url of web
    .PARAMETER groupName
        Name of group
    .EXAMPLE
        Get-Group -webUrl https://speval -groupName "SharePointEval Home Owners"
    #>
    [CmdletBinding()] 
    param (
        [string]$webUrl, 
        [string]$groupName
    )
      
    BEGIN {
        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $web = Get-SPWeb $webUrl
            $group = $web.Groups[$groupName]

                if($group) {
                    Write-Output $group
                }

                else {
                    Write-Output "group '$groupName' doesn't exist in web '$webUrl'"                
                }
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }
} 

Function Get-PermissionLevel {
    <#
    .SYNOPSIS
        Gets a permission level
    .DESCRIPTION
        Get-PermissionLevel
    .PARAMETER siteUrl
        Url of site collection
    .PARAMETER permissionLevelName
        Name of permission level
    .EXAMPLE
        Get-PermissionLevel -siteUrl https://speval -permissionLevelName "Full Control"
    #>
    [CmdletBinding()] 
    param (
        [string]$siteUrl, 
        [string]$permissionLevelName
    )
      
    BEGIN {
        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $site = Get-SPSite -Identity $siteUrl
            $web = $site.RootWeb
            $role = $web.RoleDefinitions | Where-Object {$_.name-like "*$permissionLevelName*"}

                if($role) {
                    Write-Output $role
                }

                else {
                    Write-Output "Role '$permissionLevelName' doesn't exist in web '$siteUrl'"                
                }
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }
} 

Function Get-InheritanceStatusWeb {
    <#
    .SYNOPSIS
        Gets the inheritance status of a web
    .DESCRIPTION
        Get-InheritanceStatusWeb
    .PARAMETER webUrl
        Url of web
    .EXAMPLE
        Get-InheritanceStatusWeb -webUrl https://speval
    #>
    [CmdletBinding()] 
    param (
        [string]$webUrl
    )
      
    BEGIN {
        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $web = Get-SPWeb $webUrl 
            $inheritanceStatusWeb = $web.HasUniqueRoleAssignments
            Write-Output $inheritanceStatusWeb           
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }
} 

Function Get-InheritanceStatusList {
    <#
    .SYNOPSIS
        Gets the inheritance status of a list
    .DESCRIPTION
        Get-InheritanceStatusList
    .PARAMETER listName
        Name of list
    .PARAMETER webUrl
        Url of web
    .EXAMPLE
        Get-InheritanceStatusList -webUrl "https://speval" -listName "BoolTest"
    #>
    [CmdletBinding()] 
    param (
        [string]$listName, 
        [string]$webUrl
    )
      
    BEGIN {
        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $list = Get-List -webUrl $webUrl -listName $listName
            $inheritanceStatusList = $List.HasUniqueRoleAssignments
            Write-Output $inheritanceStatusList          
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }
} 

Function Set-GroupPermissionsWeb {
    <#
    .SYNOPSIS
        Sets permissions for a SharePoint permission group
    .DESCRIPTION
        Set-GroupPermissionsWeb
    .PARAMETER webUrl
        Url of web
    .PARAMETER groupName
        Name of group
    .PARAMETER permissionLevelName
        Name of group
    .EXAMPLE
        Set-GroupPermissionsWeb -webUrl https://speval -groupName "SharePointEval Home Owners" -permissionLevelName "Full Control"
        
        Dependance on:
            Get-InheritanceStatusWeb
            Get-PermissionLevel
            Get-Group
    #>
    [CmdletBinding()] 
    param (
        [string]$webUrl, 
        [string]$groupName,
        [string]$permissionLevelName
    )
      
    BEGIN {
        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $inheritanceStatus = Get-InheritanceStatusWeb -webUrl $webUrl

                if($inheritanceStatus -eq $true){
                    $role = Get-PermissionLevel -siteUrl $webUrl -permissionLevelName $permissionLevelName
                    $group = Get-Group -webUrl $webUrl -groupName $groupName
                    $assignment = $group.ParentWeb.RoleAssignments.GetAssignmentByPrincipal($group)
                    $assignment.RoleDefinitionBindings.Add($role)
                    $assignment.Update()
                    $group.Update()
                    Write-Output "Role '$permissionLevelName' added to group '$groupName' in web '$webUrl'"     
                }

                elseif($inheritanceStatus -eq $fasle){
                    Write-Output "Web '$webUrl' doesn't have unique permissions"
                }        
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }
} 

Function Remove-GroupPermissionsWeb {
    <#
    .SYNOPSIS
        Removed permissions for a SharePoint permission group
    .DESCRIPTION
        Remove-GroupPermissionsWeb
    .PARAMETER webUrl
        Url of web
    .PARAMETER groupName
        Name of group
    .PARAMETER permissionLevelName
        Name of group
    .EXAMPLE
        Remove-GroupPermissionsWeb -webUrl https://speval -groupName "SharePointEval Home Owners" -permissionLevelName "Full Control"
        
        Dependance on:
            Get-InheritanceStatusWeb
            Get-PermissionLevel
            Get-Group
    #>
    [CmdletBinding()] 
    param (
        [string]$webUrl, 
        [string]$groupName,
        [string]$permissionLevelName
    )
      
    BEGIN {
        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $inheritanceStatus = Get-InheritanceStatusWeb -webUrl $webUrl

                if($inheritanceStatus -eq $true){
                    $role = Get-PermissionLevel -siteUrl $webUrl -permissionLevelName $permissionLevelName
                    $group = Get-Group -webUrl $webUrl -groupName $groupName
                    $assignment = $group.ParentWeb.RoleAssignments.GetAssignmentByPrincipal($group)
                    $assignment.RoleDefinitionBindings.Remove($role)
                    $assignment.Update()
                    $group.Update()
                    Write-Output "Role '$permissionLevelName' removed from group '$groupName' in web '$webUrl'"     
                }

                elseif($inheritanceStatus -eq $fasle){
                    Write-Output "Web '$webUrl' doesn't have unique permissions"
                }        
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }
} 

Function Set-GroupPermissionsList {
    <#
    .SYNOPSIS
        Sets permissions for a SharePoint permission group in list
    .DESCRIPTION
        Set-GroupPermissionsList
    .PARAMETER webUrl
        Url of web
    .PARAMETER listName
        Name of list
    .PARAMETER groupName
        Name of group
    .PARAMETER permissionLevelName
        Name of group
    .EXAMPLE
        Set-GroupPermissionsList -webUrl https://speval -listName "BoolTest" -groupName "SharePointEval Home Owners" -permissionLevelName "Contribute"
        
        Dependance on:
            Get-InheritanceStatusList
            Get-PermissionLevel
            Get-Group
            Get-List
    #>
    [CmdletBinding()] 
    param (
        [string]$webUrl, 
        [string]$listName, 
        [string]$groupName,
        [string]$permissionLevelName
    )
      
    BEGIN {
        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $inheritanceStatus = Get-InheritanceStatusList -webUrl $webUrl -listName $listName

                if($inheritanceStatus -eq $true){
                    $role = Get-PermissionLevel -siteUrl $webUrl -permissionLevelName $permissionLevelName
                    $group = Get-Group -webUrl $webUrl -groupName $groupName
                    $list = Get-List -webUrl $webUrl -listName $listName

                    $assignment = $list.RoleAssignments.GetAssignmentByPrincipal($group)
                    $assignment.RoleDefinitionBindings.Add($role)
                    $assignment.Update()
                    $group.Update()
                    Write-Output "Role '$permissionLevelName' added to group '$groupName' in list '$listName'"     
                }

                elseif($inheritanceStatus -eq $fasle){
                    Write-Output "List '$listName' doesn't have unique permissions"
                }        
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }
} 

Function Remove-GroupPermissionsList {
    <#
    .SYNOPSIS
        Removes permissions for a SharePoint permission group in list
    .DESCRIPTION
        Remove-GroupPermissionsList
    .PARAMETER webUrl
        Url of web
    .PARAMETER listName
        Name of list
    .PARAMETER groupName
        Name of group
    .PARAMETER permissionLevelName
        Name of group
    .EXAMPLE
        Remove-GroupPermissionsList -webUrl https://speval -listName "BoolTest" -groupName "SharePointEval Home Owners" -permissionLevelName "Contribute"
        
        Dependance on:
            Get-InheritanceStatusList
            Get-PermissionLevel
            Get-Group
            Get-List
    #>
    [CmdletBinding()] 
    param (
        [string]$webUrl, 
        [string]$listName, 
        [string]$groupName,
        [string]$permissionLevelName
    )
      
    BEGIN {
        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $inheritanceStatus = Get-InheritanceStatusList -webUrl $webUrl -listName $listName

                if($inheritanceStatus -eq $true){
                    $role = Get-PermissionLevel -siteUrl $webUrl -permissionLevelName $permissionLevelName
                    $group = Get-Group -webUrl $webUrl -groupName $groupName
                    $list = Get-List -webUrl $webUrl -listName $listName

                    $assignment = $list.RoleAssignments.GetAssignmentByPrincipal($group)
                    $assignment.RoleDefinitionBindings.Remove($role)
                    $assignment.Update()
                    $group.Update()
                    Write-Output "Role '$permissionLevelName' removed from group '$groupName' in list '$listName'"     
                }

                elseif($inheritanceStatus -eq $fasle){
                    Write-Output "List '$listName' doesn't have unique permissions"
                }        
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }
} 

Function Get-WebsWithNonStandardTheme {
    <#
    .SYNOPSIS
        Script to identify webs with a non standard theme
    .DESCRIPTION
        Get-WebsWithNonStandardTheme
    .PARAMETER webApp
        Web application
    .EXAMPLE
        Get-WebsWithNonStandardTheme -webApp https://speval
    #>
    param (
        [string]$webApp
    )
      
    BEGIN {
        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $webs = Get-SPSite -WebApplication $webApp -Limit all | Get-SPWeb -Limit all | Where-Object {$_.ThemedCssFolderUrl}
            
                if($webs){
                    Write-Output $webs
                }
                
                else{
                    Write-Output "no webs found with non-standard theme" 
                }          
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }
} 

Function Set-PermsLevel {
    <#
    .SYNOPSIS
        Script to set permissions in permission level - all SC's in a WA
    .DESCRIPTION
        Set-PermsLevel
    .PARAMETER webApp
        Web application
    .PARAMETER level
        Permission level
    .EXAMPLE
        Set-PermsLevel -webApp https://awebapp -level "a perm level"
    .NOTES
        Dependancy on Get-PermissionLevel function
    #>
    [CmdletBinding()] 
    param (
        [string]$webApp, 
        [string]$level
    )
      
    BEGIN {
        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{

            $sites = get-spsite -WebApplication $webApp -Limit all

                $sites | ForEach-Object {
                    $url = $_.url
                    $perm = Get-PermissionLevel -siteUrl $url -permissionLevelName $level
        
                        if($perm -NotLike "*doesn't exist in web*"){
                            $base = "ViewListItems,
                                    AddListItems,
                                    EditListItems,
                                    DeleteListItems,
                                    ApproveItems,
                                    OpenItems,
                                    ViewVersions,
                                    DeleteVersions,
                                    CancelCheckout,
                                    ManagePersonalViews,
                                    ManageLists,
                                    ViewFormPages,
                                    Open,
                                    ViewPages,
                                    AddAndCustomizePages,
                                    ApplyStyleSheets,
                                    ViewUsageData,
                                    CreateSSCSite,
                                    CreateGroups,
                                    ManagePermissions,
                                    BrowseDirectories,
                                    BrowseUserInfo,
                                    AddDelPrivateWebParts,
                                    UpdatePersonalWebParts,
                                    ManageWeb,
                                    UseClientIntegration,
                                    UseRemoteAPIs,
                                    ManageAlerts,
                                    CreateAlerts,
                                    EditMyUserInfo,
                                    EnumeratePermissions" 
        
                            $perm.BasePermissions = $base 
                            $perm.Update()
                            Write-Output "$url updated - $level"                
                        }
                }
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }
}

Function Get-PermissionReport {
    <#
    .SYNOPSIS
        Generates a CSV permission report for all webs in a particular web application
    .DESCRIPTION
        Get-PermissionReport
    .PARAMETER webApp
        The web application you want to generate the report for
    .PARAMETER outPath
        The path for the CSV
    .EXAMPLE
        Get-PermissionReport -webApp https://awebapp -outPath "D:\testPermReport.csv"
    #>
    [CmdletBinding()] 
    param (
        [string]$webApp, 
        [string]$outPath
    )
      
    BEGIN {
        Start-SPAssignment -Global
        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $webs = Get-SPSite -WebApplication $webApp -Limit all | Get-SPWeb -Limit all
            set-variable -option constant -name out -value $outPath
            "sep=;" | Out-File $out 
            "WebName;WebUrl;Inherited;User;Group;GivenDirect" | Out-File $out -append

                $webs | ForEach-Object {
                    $webUrl = $_.Url
                    $webName = $_.Title

                        if($_.HasUniquePerm -eq $false){
                            $groups = $_.groups
                            $perms =  $_.RoleAssignments | Where-Object {$_.Member -like "*i:0#.w*" -and $_.RoleDefinitionBindings.name -notlike "*Limited*"}

                                $perms | ForEach-Object{
                                    $member = $_.Member
                                    $person = $member.DisplayName
                                    $defs = $_.RoleDefinitionBindings

                                        $defs | ForEach-Object {
                                        $defName = $_.name

                                            if($defName -ne "Limited Access"){
                                                "$webName" + ";" + "$webUrl" + ";" + "Yes" + ";" + "$person" + ";" + "N/A" + ";" + "$defName" + ";" | Out-File $out -append     
                                            }
                                        }
                                }

                                $groups | ForEach-Object{
                                    $groupName = $_.Name
                                    $users = $_.Users

                                        if($users){

                                            $users | ForEach-Object{
                                                $userName = $_.DisplayName
                                                "$webName" + ";" + "$webUrl" + ";" + "Yes" + ";" + "$userName" + ";" + "$groupName" + ";" + "N/A" + ";" | Out-File $out -append 
                                            }                                  
                                        }
                                }
                        }

                        if($_.HasUniquePerm -eq $true){
                            $groups = $_.groups
                            $perms =  $_.RoleAssignments | Where-Object {$_.Member -like "*i:0#.w*" -and $_.RoleDefinitionBindings.name -notlike "*Limited*"}

                                $perms | ForEach-Object{
                                    $member = $_.Member
                                    $person = $member.DisplayName
                                    $defs = $_.RoleDefinitionBindings

                                        $defs | ForEach-Object {
                                        $defName = $_.name

                                            if($defName -ne "Limited Access"){
                                                "$webName" + ";" + "$webUrl" + ";" + "No" + ";" + "$person" + ";" + "N/A" + ";" + "$defName" + ";" | Out-File $out -append     
                                            }
                                        }
                                }

                                $groups | ForEach-Object{
                                    $groupName = $_.Name
                                    $users = $_.Users

                                        if($users){

                                            $users | ForEach-Object{
                                                $userName = $_.DisplayName
                                                "$webName" + ";" + "$webUrl" + ";" + "No" + ";" + "$userName" + ";" + "$groupName" + ";" + "N/A" + ";" | Out-File $out -append 
                                            }                                  
                                        }
                                }
                        }        
                }

        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }

    END {
        Stop-SPAssignment -Global    
    }
}

Function Get-VersionStatusList {
    <#
    .SYNOPSIS
        Returns the versioning stauts for a list
    .DESCRIPTION
        Get-VersionStatusList
    .PARAMETER site
        SharePoint site Url
    .PARAMETER listName
        Name of list
    .EXAMPLE
        Get-VersionStatusList -site https://asite -listName "A List"
    #>
    [CmdletBinding()] 
    param (
        [string]$site, 
        [string]$listName
    )
      
    BEGIN {
        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $web = Get-SPWeb $site
            $list = $web.Lists[$listName]

                if($list) {
                    $versioning = $list.EnableVersioning 
                    
                        if($versioning -eq $true){
                            $minor = $list.EnableMinorVersions

                                if($minor -eq $true){
                                    Write-Output "Minor"
                                }
                                else{
                                    Write-Output "Major"            
                                }

                        }
                        else{
                            Write-Output "None"    
                        }   
                }

                else {
                    Write-Output "list $lib doesnt exist in $site"                
                }
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }

    END {
    }
}  

Function Get-CheckOutStatusList {
    <#
    .SYNOPSIS
        Returns the versioning stauts for a list
    .DESCRIPTION
        Get-VersionStatusList
    .PARAMETER site
        SharePoint site Url
    .PARAMETER listName
        Name of list
    .EXAMPLE
        Get-CheckOutStatusList -site https://pp-edrm.cma.gov.uk/sites/ict/43 -listName "Analysis"
    #>
    [CmdletBinding()] 
    param (
        [string]$site, 
        [string]$listName
    )
      
    BEGIN {
        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $web = Get-SPWeb $site
            $list = $web.Lists[$listName]

                if($list) {
                    $checkOut = $list.ForceCheckout 
                    
                        if($checkOut -eq $true){                            
                            Write-Output "Enforced" 
                        }
                        else{
                            Write-Output "Optional"    
                        }   
                }

                else {
                    Write-Output "list $lib doesnt exist in $site"                
                }
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }

    END {
    }
} 

Function Get-FolderStatusList {
    <#
    .SYNOPSIS
        Returns the versioning stauts for a list
    .DESCRIPTION
        Get-VersionStatusList
    .PARAMETER site
        SharePoint site Url
    .PARAMETER listName
        Name of list
    .EXAMPLE
        Get-FolderStatusList -site https://pp-edrm.cma.gov.uk/sites/ict/43 -listName "Analysis"
    #>
    [CmdletBinding()] 
    param (
        [string]$site, 
        [string]$listName
    )
      
    BEGIN {
        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $web = Get-SPWeb $site
            $list = $web.Lists[$listName]

                if($list) {
                    $folders = $list.EnableFolderCreation
                    
                        if($folders -eq $false){                            
                            Write-Output "Disabled" 
                        }
                        else{
                            Write-Output "Allowed"    
                        }   
                }

                else {
                    Write-Output "list $lib doesnt exist in $site"                
                }
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }

    END {
    }
} 

Function Get-TermsWithKeyword {
    <#
    .SYNOPSIS
        Finds terms with particular keywords
    .DESCRIPTION
        Get-TermsWithKeyword
    .PARAMETER web
        Web site address
    .PARAMETER filter
        Keyword that you want to look for
    .PARAMETER group
        Taxonomy group
    .EXAMPLE
        Get-TermsWithKeyword -web https://asite -group "Taxonomy" -filter "A Keyword"
    #>
    param (
        [string]$web,
        [string]$filter,
        [string]$group
    )

    $site = Get-SPWeb $web
    $ts = Get-SPTaxonomySession -Site $site.site
    $tstore = $ts.TermStores[0]
    $tgroup = $tstore.Groups[$group]
    $tsets = $tgroup.TermSets
        
        $tsets | ForEach-Object{
            $terms = $_.terms | Where-Object {$_.name -like "*$filter*"}

                $terms | ForEach-Object{
                    $term = $_.name
                    $parent = $_.TermSets.name
                    Write-Output "$term  - $parent"
                }                
        }
}

Function Set-DefaultColValueMM {
    <#
    .SYNOPSIS
        Sets library managed metadata column defaults  
    .DESCRIPTION
        Set-DefaultColValueMM 
    .PARAMETER web
        Web site address
    .PARAMETER library
        Library name
    .PARAMETER field
        Internal field name for column you're wanting to set
    .PARAMETER group
        Term group label
    .PARAMETER set
        Term set label
    .PARAMETER tag
        Term label
    .EXAMPLE
        Set-DefaultColValueMM -web "http://aweb -library "a library" -field "a field" -group "a term group"-set "a term set" -tag "a tag"
    #>
    param (
        [string]$web,
        [string]$library,
        [string]$set,
        [string]$tag,
        [string]$field,
        [string]$group
    )

    $site = Get-SPWeb -Identity $web
    $ts = Get-SPTaxonomySession -Site $site.site.Url
    $tstore = $ts.TermStores[0]
    $tgroup = $tstore.Groups[$group]
    $tset = $tgroup.TermSets[$set]
    $term = $tset.Terms["$tag"]
    $format = "1033;#" + $term.Name + "|" + $term.Id
    $list = $site.lists[$library]
    $columnDefault = New-Object Microsoft.Office.DocumentManagement.MetadataDefaults($list)
    $folder = $list.RootFolder.ServerRelativeUrl
    $columnDefault.SetFieldDefault($folder, $field, $format)
    $columnDefault.Update()
    Write-Host Default value $tag set -ForegroundColor White -BackgroundColor DarkCyan

}

Function Get-IMPolicyStatusLibrary {
    <#
    .SYNOPSIS
        Returns IM policy status of a library
    .DESCRIPTION
        Get-IMPolicyLibrary 
    .PARAMETER site
        The website
    .PARAMETER lib
        The Library
    .EXAMPLE
        Get-IMPolicyStatusLibrary -site https://speval -lib "customLib"
    #>
    [CmdletBinding()] 
    param (
        [string]$site, 
        [string]$lib
    )
      
    BEGIN {
        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $web = Get-SPWeb $site
            $list = $web.Lists[$lib]
            $type = $list.BaseType

                if($list) {
                    
                    if ($type -eq "DocumentLibrary"){ 
                        try {
                            $policy = [Microsoft.Office.RecordsManagement.InformationPolicy.ListPolicySettings]($list)
                            $hasPolicy = $policy.ListHasPolicy
                            Write-Output $hasPolicy                 
                        }
        
                        catch {
                            $error = $_
                            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"                   
                        }
                    }
                    else{
                        Write-Output "list $lib isn't a document library"
                    }
                }

                else {
                    Write-Output "list $lib doesnt exist in $site"                
                }
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }
} 

Function Disable-ThrottlingList {
    <#
    .SYNOPSIS
        Disables throttling in a list
    .DESCRIPTION
        Disable-ThrottlingList 
    .PARAMETER site
        The website
    .PARAMETER list
        The title title
    .EXAMPLE
        Disable-ThrottlingList -site https://speval -list "customLib"
    #>
    [CmdletBinding()] 
    param (
        [string]$site, 
        [string]$list
    )
      
    BEGIN {
        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $web = Get-SPWeb $site
            $title = $web.Title
            $lst = $web.Lists[$list]

                if($lst) {
                    try {
                        $lst.EnableThrottling = $false
                        Write-Output "Throttling disabled in $list | $title"                    
                    }
        
                    catch {
                        $error = $_
                        Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"                   
                    }
                }

                else {
                    Write-Output "list $list doesnt exist in $site"                
                }
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }
} 

Function Enable-ThrottlingList {
    <#
    .SYNOPSIS
        Enables throttling in a list
    .DESCRIPTION
        Enable-ThrottlingList
    .PARAMETER site
        The website
    .PARAMETER list
        The title title
    .EXAMPLE
        Enable-ThrottlingList -site https://speval -list "customLib"
    #>
    [CmdletBinding()] 
    param (
        [string]$site, 
        [string]$list
    )
      
    BEGIN {
        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $web = Get-SPWeb $site
            $title = $web.Title
            $lst = $web.Lists[$list]

                if($lst) {
                    try {
                        $lst.EnableThrottling = $true
                        Write-Output "Throttling enabled in $list | $title"                    
                    }
        
                    catch {
                        $error = $_
                        Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"                   
                    }
                }

                else {
                    Write-Output "list $list doesnt exist in $site"                
                }
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }
} 

Function Set-LibRetention {
    <#
    .SYNOPSIS
        Use this to set a retention period policy in a library
    .DESCRIPTION
        Set-LibRetention
    .PARAMETER site
        The website
    .PARAMETER lib
        Library name
    .PARAMETER field
        Field name. Has to be of type datetime (created, modified, document date etc)
    .PARAMETER Period
        Choose a period of time (days, years, or months)
    .PARAMETER Number
        Choose a unit of time to retain content (has to be an integer) 
    .PARAMETER Action
        Choose either MoveToRecycleBin, delete, DeletePreviousDrafts, or DeletePreviousVersions
    .EXAMPL
        Set-LibRetention -site https://asite -lib "alib" -field "adatetypefield" -period "days" -number "2" -action "delete DeletePreviousVersions
    #>
    [CmdletBinding()] 
    param (
        [string]$site, 
        [string]$lib,
        [string]$description,
        [string]$field,
        [int]$number,
        [ValidateSet("days", "years", "months")][string]$period,
        [ValidateSet("MoveToRecycleBin", "delete", "DeletePreviousDrafts", "DeletePreviousVersions")][string]$action
    )
      
    BEGIN {
        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $web = Get-SPWeb $site 
            $list = $web.Lists[$lib]

                if($list) {
                    $col = $list.Fields | Where-Object {$_.title -eq $field -and $_.type -eq "DateTime"}

                        if($col){
                            
                            try {
                                $fieldID = $col.id
                                $id = "Microsoft.Office.RecordsManagement.PolicyFeatures.Expiration.Action.$action"
                                $xml = "<Schedules nextStageId=`"2`">"+
                                            "<Schedule type=`"Default`">"+
                                            "<stages>"+
                                            "<data stageId=`"1`">"+
                                            "<formula id=`"Microsoft.Office.RecordsManagement.PolicyFeatures.Expiration.Formula.BuiltIn`">"+
                                            "<number>"+$number+"</number>"+
                                            "<property>"+$field+"</property>"+
                                            "<propertyId>"+$fieldID+"</propertyId>"+
                                            "<period>"+$period+"</period>"+
                                            "</formula>"+
                                            "<action type=`"action`" id=`""+$id+"`" />"+
                                            "</data>"+
                                            "</stages>"+
                                            "</Schedule></Schedules>"

                                $title = $web.title
                                $root = $list.RootFolder.ServerRelativeUrl
                                $policy = [Microsoft.Office.RecordsManagement.InformationPolicy.ListPolicySettings]($list)
                                $policy.SetRetentionSchedule($root, $xml, $description)
                                $policy.UseListPolicy = $true
                                $policy.Update()
                                Write-Output "Policy set in $lib | $title"                    
                            }
        
                            catch {
                                $error = $_
                                Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"                   
                            }
                        }

                        else{
                            Write-Output "field $field doesnt exist or is not of type datetime" 
                        }
                }

                else {
                    Write-Output "list $lib doesnt exist in $title"                
                }
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }

    END {
        $web.Dispose()   
    }
} 

Function Set-DefaultView {
    <#
    .SYNOPSIS
        Sets columns in default library view 
    .DESCRIPTION
        Set-DefaultView
    .PARAMETER libraryName
        Name of library
    .PARAMETER webSite
        The website
    .EXAMPLE
        Enable-ThrottlingList -site https://speval -list "customLib"
    #>
    param (
        [string]$libraryName,
        [string]$webSite
    )

    $items = "DocIcon","LinkFilename","Modified","Editor","CheckoutUser"
    $web = Get-SPWeb -Identity $webSite
    $list = $web.Lists[$libraryName]
    $view = $list.Views["All Documents"]
    $view.ViewFields.DeleteAll()
    $view.Update()

        foreach($item in $items){
            $view.ViewFields.Add("$item")
            Write-Output "$item column added"
        }
    $view.Update()
}

Function Set-CrawlsContentSource {
    <#
    .SYNOPSIS
        Start, stop, pause etc. crawls for content source
    .DESCRIPTION
        Set-CrawlsContentSource
    .PARAMETER contentSource
        The content source name
    .PARAMETER action
        PauseCrawl, StopCrawl, StartFullCrawl, StartIncrementalCrawl etc.  
    .EXAMPLE
        Set-CrawlsContentSource -contentSource "SharePoint" -action PauseCrawl
    #>
    [CmdletBinding()] 
    param (
        [string]$contentSource, 
        [ValidateSet("PauseCrawl", "StopCrawl", "StartFullCrawl", "StartIncrementalCrawl")][string]$action
    )
      
    BEGIN {

        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $sa = Get-SPEnterpriseSearchServiceApplication -Identity "Search Service Application"
            $cs = $sa | Get-SPEnterpriseSearchCrawlContentSource -Identity $contentSource

                try {
                    $command = "$" + "cs." + "$action" + "()"
                    invoke-expression $command
                    Write-Output "Action $action initiated"                    
                }
        
                catch {
                    $error = $_
                    Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"                   
                }
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }
} 

Function Remove-SPUserWrap {
    <#
    .SYNOPSIS
        Removes a user from a site collection 
    .DESCRIPTION
        Remove-SPUserWrap
    .PARAMETER site
        Site Url
    .PARAMETER user
        DOMAIN\username
    .EXAMPLE
        Remove-SPUserWrap -site https://asite -user "DOMAIN\username"

            ##Bulk from CSV##

                $accounts = Import-Csv -Path source.csv
                $sites = Get-SPSite -WebApplication "https://awebapp" -Limit all

                    foreach ($account in $accounts){
                        $id = $account.account

                            $sites | ForEach-Object{
                                $url = $_.Url
                                Remove-SPUserWrap -site $url -user $id
                            }
                        Pause
                    }
    #>
    [CmdletBinding()] 
    param (
        [string]$site, 
        [string]$user
    )
      
    BEGIN {

        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $full = "i:0#.w|$user"
            Remove-SPUser -Identity $full -Web $site -confirm:$false
            Write-Output "$user deleted from $site"

        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }
}

Function Clear-FileLock {
    <#
    .EXAMPLE
        Clear-FileLock -web https://aweb -list "alist" -doc "adoc.docx"
    #>
    [CmdletBinding()] 
    param (
        [string]$web, 
        [string]$list,
        [string]$doc
    )
      
    BEGIN {
        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $wb = Get-SPWeb $web
            $lst = $wb.Lists[$list]
            $item = $lst.Items | Where-Object {$_.name -eq $doc}
            $item.File.ReleaseLock($item.File.LockId)
            Write-Output "removed lock"   
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }

    END {
        $wb.dispose()    
    }
}

Function Clear-FileLockImpersonated {
    <#
    .EXAMPLE
        Clear-FileLockImpersonated -web https://aweb -list "alist" -doc "adoc.docx"
    #>
    [CmdletBinding()] 
    param (
        [string]$web, 
        [string]$list,
        [string]$doc
    )
      
    BEGIN {
        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $wb = Get-SPWeb $web
            $lst = $wb.Lists[$list]
            $item = $lst.Items | Where-Object {$_.name -eq $doc}
            $userId = $item.File.LockedByUser.ID
            $user = $wb.AllUsers.GetById($userId)
            $impSite = New-Object Microsoft.SharePoint.SPSite($wb.Url, $user.UserToken)
            $impWeb = $impSite.OpenWeb()
            $impList = $impWeb.Lists[$list]
            $impItem = $impList.Items | Where-Object {$_.name -eq $doc}
            $impItem.File.ReleaseLock($impItem.File.LockId)
            Write-Output "removed lock for $user"   
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }

    END {
        $impWeb.Dispose()
        $impSite.Dispose()
        $wb.dispose()    
    }
}

Function New-KerbSSLWA {
    <#
    .SYNOPSIS
        Creates a SSL web application with Kerberos Auth
    .DESCRIPTION
        New-KerbWA
    .PARAMETER name
        Web application name
    .PARAMETER url
        Url
    .PARAMETER hostHeader
        Url
    .PARAMETER applicationPoolAccount
        Application pool account
    .PARAMETER databaseName
        Content database name
    .EXAMPLE
        New-KerbSSLWA -name "POC" -url "https://poc.cma.gov.uk" -hostHeader "poc.cma.gov.uk" -applicationPoolAccount "CMA\spevaladmin" -databaseName "SP_Content_POC"
    #>
    [CmdletBinding()]
    param (
        [string]$name,
        [string]$url,
        [string]$hostHeader,
        [string]$applicationPoolAccount,
        [string]$databaseName
    )
      
    BEGIN {

        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $auth = New-SPAuthenticationProvider -DisableKerberos:$false
            New-SPWebApplication -Name $name -HostHeader $hostHeader -Url $url -Port 443 -ApplicationPoolAccount (Get-SPManagedAccount $applicationPoolAccount) -ApplicationPool $name -AuthenticationProvider $auth -SecureSocketsLayer -DatabaseName $databaseName
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }
}

Function Set-WABrowserFileHandlingPermissive {
    <#
    .SYNOPSIS
        Sets browser file handling for web application to permissive
    .DESCRIPTION
        Set-WABrowserFileHandlingPermissive
    .PARAMETER webApplicationUrl
        Web application url
    .EXAMPLE
        Set-WABrowserFileHandlingPermissive -webApplicationUrl "https://poc.cma.gov.uk"
    #>
    [CmdletBinding()]
    param (
        [string]$webApplicationUrl
    )
      
    BEGIN {

        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $wa = Get-SPWebApplication -Identity $webApplicationUrl
            $wa.BrowserFileHandling = "permissive"
            $wa.Update()    
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }
}

Function Set-WATimeZoneUK {
    <#
    .SYNOPSIS
        Sets time zone for web application to UK
    .DESCRIPTION
        Set-BrowserFileHandlingPermissive
    .PARAMETER webApplicationUrl
        Web application url
    .EXAMPLE
        Set-WATimeZoneUK -webApplicationUrl "https://poc.cma.gov.uk"
    #>
    [CmdletBinding()]
    param (
        [string]$webApplicationUrl
    )
      
    BEGIN {

        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $wa = Get-SPWebApplication -Identity $webApplicationUrl
            $wa.DefaultTimeZone = "2"
            $wa.Update()    
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }
}

Function New-WAUserPolicy {
    <#
    .SYNOPSIS
        Adds a user to policy for web application
    .DESCRIPTION
        New-WAUserPolicy
    .PARAMETER webApplicationUrl
        Web application url
    .PARAMETER user
        User account
    .PARAMETER userDisplayName
        User displayname
    .PARAMETER permissionLevel
        Permission level (FullRead, FullControl etc.)
    .EXAMPLE
        New-WAUserPolicy -webApplicationUrl "https://poc.cma.gov.uk" -user "cma\matt.warburton" -userDisplayName "Matt Warburton" -permissionLevel "FullControl"
    #>
    [CmdletBinding()]
    param (
        [string]$webApplicationUrl,
        [string]$user,
        [string]$userDisplayName,
        [string]$permissionLevel
    )
      
    BEGIN {

        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $wa = Get-SPWebApplication -Identity $webApplicationUrl
            $policy = $wa.Policies.Add($user, $userDisplayName)   
            $policyRole = $wa.PolicyRoles.GetSpecialRole([Microsoft.SharePoint.Administration.SPPolicyRoleType]::$permissionLevel)   
            $policy.PolicyRoleBindings.Add($policyRole)
            $wa.Update()    
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }
}

Function New-SCPub {
    <#
    .SYNOPSIS
        Adds a new publishing site collection
    .DESCRIPTION
        New-SCPub
    .PARAMETER url
        URL for new site collection
    .PARAMETER contentDatabase
        Name for content database
    .PARAMETER websiteName
        Display name for site collection
    .PARAMETER primaryLogin
        Primary site collection administrator          
    .EXAMPLE
        New-SCPub -url https://speval/sites/pubstest -ContentDatabase pubstest -WebsiteName "Pub test" -PrimaryLogin "cma\spevaladmin"
    #>
    param (
        [Parameter(Mandatory=$true)][string]$url, 
        [Parameter(Mandatory=$true)][string]$contentDatabase, 
        [Parameter(Mandatory=$true)][string]$websiteName, 
        [Parameter(Mandatory=$true)][string]$primaryLogin
    )

            $db = Get-SPContentDatabase -Identity $ContentDatabase -ErrorAction SilentlyContinue
            $existingSC = Get-SPSite -Identity $url -ErrorAction SilentlyContinue
            $code = $codes.$template
            
                if($existingSC -eq $null){ 
                
                    if($db -ne $null){
                                        
                        New-SPSite -Url $url –ContentDatabase $contentDatabase -Name $websiteName -Template "BLANKINTERNETCONTAINER#0" -OwnerAlias $primaryLogin
                        write-host "Add-SiteCollection - Success! $url created" -BackgroundColor "Green" -ForegroundColor "White" 
                    }
                    else{
                        Write-Host "Content database $contentDatabase doesn't exit" -BackgroundColor "Red" -ForegroundColor "White"
                    
                            Do {
                            "[1] Add content DB"
                            "[2] Exit"
                            $Selection = Read-Host "Please select an option"

                                if($selection -eq "1"){
                                    $wa = read-host "please enter the web application that the content DB should be associated with"
                                    New-CDB -name $contentDatabase -webapp $wa  

                                }
                                elseif($selection -eq "2"){
                                    exit
                                }
                                else{
                                    write-host "invalid selection"
                                }

                            }
                            Until (($Selection -eq 1) -or ($Selection -eq 2))
                                                   
                        New-SPSite -Url $url –ContentDatabase $contentDatabase -Name $websiteName -Template "$code" -OwnerAlias $primaryLogin
                        write-host "Add-SiteCollection - Success! $url created" -BackgroundColor "Green" -ForegroundColor "White"
                    }
                }
                else{
                    Write-Host "A site collection with the same address already exists, please try again" -BackgroundColor "Red" -ForegroundColor "White"    
                }                               
}

Function New-CDB {
    <#
    .SYNOPSIS
        Adds a content database
    .DESCRIPTION
        New-CDB
    .PARAMETER webapp
        Web application for content database
    .PARAMETER name
        Name for content database
    .EXAMPLE
        New-CDB -Name "blah blah" -WebApplication https://speval
    #>
    param (
        [Parameter(Mandatory=$true)][string]$name, 
        [Parameter(Mandatory=$true)][string]$webapp
    )
        $waTest = Get-SPWebApplication -Identity $webapp -ErrorAction SilentlyContinue
        $dbTest = Get-SPContentDatabase -Identity $name -ErrorAction SilentlyContinue

        if($dbtest -eq $null){

            if($waTest -ne $null){
                New-SPContentDatabase -Name $name -WebApplication $webapp -ErrorAction SilentlyContinue
                write-host "Content database $name created" -BackgroundColor "Green" -ForegroundColor "White"
            }
            else{
                Write-Host "Web app $name not found, please try again" -BackgroundColor "Red" -ForegroundColor "White"    
            }
        }
        else{
            Write-Host "$name already exists" -BackgroundColor "Red" -ForegroundColor "White"     
        }
}

Function Set-WebTimeZoneUK {
    <#
    .SYNOPSIS
        Sets time zone for a web to UK
    .DESCRIPTION
        Set-WebTimeZoneUK
    .PARAMETER webUrl
        Web url
    .EXAMPLE
        Set-WebTimeZoneUK -webUrl "https://speval/sites/pubstest1"
    #>
    [CmdletBinding()]
    param (
        [string]$webUrl
    )
      
    BEGIN {

        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $web =  get-spweb $webUrl
            $web.RegionalSettings.LocaleId = 2057
            $web.Update()    
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }
}

Function New-PubPage {
    param (
        [string]$web,
        [string]$name,
        [string]$title,
        [string]$layout  
    )

    $wb = Get-SPWeb $web
    $pubWb =[Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($wb)
    $pl = $pubWb.GetAvailablePageLayouts() | Where { $_.Name -eq $layout }
    $page = $pubWb.AddPublishingPage($name, $pl)
    $page.Update()
    $page.Title = $title
    $page.Update()
    $page.CheckIn("")
    $page.ListItem.File.Publish("")
    $wb.Dispose()
}

Function Set-WelcomePage {
    param (
        [string]$web,
        [string]$path
    )

    $wb = Get-SPWeb $web
    $root = $wb.RootFolder
    $root.WelcomePage = $path
    $root.Update()
    $wb.Update()
    $wb.Dispose()
}

Function Delete-AllVersions {
    [CmdletBinding()]
    param (
        [string]$web, 
        [string]$folderPath
    )
      
    BEGIN {

        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $site = Get-SPWeb $web
            $folder = $site.GetFolder($folderPath)
            $files = $folder.files

                $files | ForEach-Object{
                    $_.Versions.DeleteAll()
                }
        }
        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }
}

Function Add-NavNode {
    [CmdletBinding()]
    param (
        [string]$web, 
        [string]$title,
        [string]$url
    )
      
    BEGIN {

        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $site = Get-SPWeb $web
            $link = New-Object Microsoft.SharePoint.Navigation.SPNavigationNode -ArgumentList  @($title, $url); 
            $site.Navigation.TopNavigationBar.AddAsLast($link)
        }
        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }
}

Function New-UserProfileCSV {
    <#
    .SYNOPSIS
        Use to create user profiles & properties from a CSV file datasource. 

        The first column header in the CSV should be labled "account". Ensure other column header labels match the internal
        name of the relevant UPA properties.
    .PARAMETER mysiteHost
        Enter the url for the mysite host site collection.
    .PARAMETER csvPath
        Enter the path to the CSV file. 
    .EXAMPLE
        New-UserProfileCSV -mysiteHost https://mysite -csvPath "C:\source.csv"
    #>
    [CmdletBinding()] 
    param (
        [string]$mysiteHost, 
        [string]$csvPath
    )
      
    BEGIN {

        $ErrorActionPreference = 'stop'   
    }
    
    PROCESS {

        try{
            $source = Import-Csv -Path $csvPath
            $context = get-spservicecontext($mysiteHost)
            $upm = new-object microsoft.office.server.userprofiles.userprofilemanager($context)

                $source | ForEach-Object {
                $dataRow = $_
                $account = $dataRow.account

                    if($upm.UserExists($account)) {
                        Write-Output "$account already exists"
                    }
                        
                    else {
                        $props = $source | Get-member -MemberType 'NoteProperty' | Select-Object -ExpandProperty 'Name'
                        $newprofile = $upm.createuserprofile($account)
                        Write-Output "$account - created"    
                            
                            $props | where-object {$_ -ne "account"} | ForEach-Object {
                                $prop = $_
                                $data = $dataRow."$prop"
                                $newprofile[$prop].add($data)
                                Write-Output "$account - $prop - $data"
                            }
                        
                        $newprofile.commit()
                        Write-Output "$account - finished"
                    }
                }
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)" 
        }
    }
}

Function Remove-UserProfileCSV {
    <#
    .SYNOPSIS
        Use to bulk delete user profiles from a CSV file datasource. 

        The first column header in the CSV should be labled "account".
    .PARAMETER mysiteHost
        Enter the url for the mysite host site collection.
    .PARAMETER csvPath
        Enter the path to the CSV file. 
    .EXAMPLE
        Remove-UserProfileCSV -mysiteHost https://mysite -csvPath "C:\source.csv"
    #>
    [CmdletBinding()] 
    param (
        [string]$mysiteHost, 
        [string]$csvPath
    )
      
    BEGIN {

        $ErrorActionPreference = 'stop'   
    }
    
    PROCESS {

        try{
            $source = Import-Csv -Path $csvPath
            $context = get-spservicecontext($mysiteHost)
            $upm = new-object microsoft.office.server.userprofiles.userprofilemanager($context)

                $source | ForEach-Object {
                $dataRow = $_
                $account = $dataRow.account

                    if($upm.UserExists($account)) {
                        $newprofile = $upm.removeuserprofile($account)
                        Write-Output "$account - deleted"
                    }
                        
                    else {
                        Write-Output "$account doesn't exist"
                    }
                }
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)" 
        }
    }
} 

Function Remove-UserProfile {
    <#
    .EXAMPLE
        Remove-UserProfile -mysiteHost -account
    #>
    [CmdletBinding()] 
    param (
        [string]$mysiteHost, 
        [string]$account
    )
      
    BEGIN {

        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $context = get-spservicecontext($mysiteHost)
            $upm = new-object microsoft.office.server.userprofiles.userprofilemanager($context)
            $newprofile = $upm.removeuserprofile($account)
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }
}

Function Get-MaxedFilesLengthWeb {
    <#
    .EXAMPLE
        Get-MaxedFilesLengthWeb -web "https://aweb" | Out-File C:\Users\gen_admin\Desktop\test.csv
    #>
    [CmdletBinding()] 
    param (
        [string]$web,
        [string]$outPath
    )
      
    BEGIN {

        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $wb = get-spweb $web
            $lists = $wb.lists

                $lists | ForEach-Object{
                    $items = $_.items | Where-Object {($wb.url.Length + $_.url.length) -gt 2}

                        $items | ForEach-Object {
                            $itemUrl = $_.url
                            Write-Output $web/$itemUrl
                        }
                }
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }
} 

Function Delete-UsersSC {
    <#
    .EXAMPLE
        Delete-UsersSC -sc "https://asc"
    #>
    [CmdletBinding()] 
    param (
        [string]$sc
    )
      
    BEGIN {

        $ErrorActionPreference = 'Stop'    
    }
    
    PROCESS {

        try{
            $users = Get-SPUser -Web $sc

                $users | ForEach-Object {
                    $login = $_.userLogin
                    
                    if($_.IsSiteAdmin -or $login -eq "SHAREPOINT\system") {
                        Write-Output "skipping SCA/System - $login"
                    }
                    else {
                        Remove-SPUser -Identity $login -Web $sc -Confirm:$False
                        Write-Output "deleted - $login"
                    }
                }
        }

        catch{
            $error = $_
            Write-Output "$($error.Exception.Message) - Line Number: $($error.InvocationInfo.ScriptLineNumber)"  
        }
    }
} 