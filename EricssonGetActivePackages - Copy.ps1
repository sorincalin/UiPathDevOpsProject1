############# EDIT the following section #############
$TenantConfig = [PSCustomObject]@{
    'url'   = "https://casorin"
    'tenant' = "default"
    'username' = "admin"
    'password' = "1qazXSW@"
    'tokenExpirationInMinutes' =  30
}

$LocalPackageFolder = "$PSScriptRoot\MyPackages\" #local folder where packages will be downloaded, unpacked and analyzed
$LocalPackageFolder = "C:\Users\sorin.calin\Downloads\Test"

# You can Disable Debug logs by commenting the next line
$DebugPreference = 'Continue'

######################################################

Import-Module UiPath.PowerShell

########################################

function GetFilesFromNupkg
{
    Param([PSObject] $nupkgFilePath, [PSObject] $destinationFolder)
    
    $nupkgFile = Get-Item -Path $nupkgFilePath
    $zipFilePath = $nupkgFile.DirectoryName + "\" + $nupkgFile.BaseName + ".zip"

    Rename-Item $nupkgFilePath $zipFilePath

    Expand-Archive $zipFilePath $destinationFolder
    Get-ChildItem -path $destinationFolder -Exclude 'lib' | Remove-Item -Recurse -force
    Rename-Item "$destinationFolder\lib" "$destinationFolder\lib-temp-move"
    dir "$destinationFolder\lib-temp-move\net45" | mv -dest $destinationFolder
    Remove-Item "$destinationFolder\lib-temp-move\" -Recurse -force
    Rename-Item $zipFilePath $nupkgFilePath
}

function GetActivePackages
{
    Param([PSObject] $tenantConfig, [PSObject] $packageFolder)
    
    $downloadedPackageFolders = New-Object System.Collections.ArrayList

    $authResponse = Get-UiPathAuthToken -URL $tenantConfig.url -TenantName $tenantConfig.tenant -Username $tenantConfig.username -Password $tenantConfig.password -Session
    
    $packages = Get-UiPathPackage | Get-UiPathPackageVersion
    $activePackages = $packages | Where-Object {$_.IsActive}
    Write-Debug "Found packages: $($packages.Count) Active packages: $($activePackages.Count)"

    # For a large number of files to be uploaded, re-authentication is required periodically
    $authTime = Get-Date -Year 1970 #seeting this to ensure authentication at first iteration
    foreach($package in $activePackages)
    {
        # If we have less than 5 minutes before token expiration, perform an authentication
        if (((Get-Date) - $authTime).TotalMinutes -gt ($tenantConfig.tokenExpirationInMinutes - 5))
        {
            $authTime = Get-Date
            $authResponse = Get-UiPathAuthToken -URL $tenantConfig.url -TenantName $tenantConfig.tenant -Username $tenantConfig.username -Password $tenantConfig.password -Session
        }
        try
        {
            $endpoint = "$($tenantConfig.url)/odata/Processes/UiPath.Server.Configuration.OData.DownloadPackage(key='$($package.Key)')"
            
            $fileName = "$packageFolder\$($package.Id)_$($package.Version).nupkg"
            $extractionFolder = "$packageFolder\$($package.Id)_$($package.Version)"
            
            Write-Debug "Downloading $fileName and extracting to $extractionFolder"

            $response = Invoke-WebRequest -Uri $endpoint -Method GET -Headers @{"Authorization"="Bearer " + $authResponse.Token; "accept"= "image/file"} -OutFile $fileName
            GetFilesFromNupkg $fileName $extractionFolder
            $downloadedPackageFolders.Add($extractionFolder)
        }
        catch
        {
            Write-Error "ERROR downloading package <$packageName>: $_ at line $($_.InvocationInfo.ScriptLineNumber)"
        }
    }
}

function GetProjectInformation
{
    Param([PSObject] $projectFile)

    $xamlFilesObject = New-Object System.Collections.ArrayList
    $projectJsonContent = Get-Content -Path $file.FullName | ConvertFrom-Json
    
    # Get all XAML files from the Project Folder
    $projectFolder = Split-Path -Path $file.FullName -Parent
    $xamlFiles = (Get-ChildItem -Path $projectFolder -Force -Recurse -Filter "*.xaml")
    
    if(Get-Member -inputObject $projectJsonContent -name "dependencies" -Membertype Properties)
    {
        $dependenciesNames = $projectJsonContent.dependencies.PSObject.Properties.Name
        $projectJsonContent | Add-Member -NotePropertyName dependenciesNames -NotePropertyValue $dependenciesNames
                
        $dependenciesNamesWithVersion = $projectJsonContent.dependencies.PSObject.Properties | Foreach {"$($_.Name)$($_.Value)"}
        $projectJsonContent | Add-Member -NotePropertyName dependenciesNamesWithVersion -NotePropertyValue $dependenciesNamesWithVersion
    }

    foreach($xamlFile in $xamlFiles)
    {
        try
        {
            $xamlFileObject = New-Object PSObject
            $xamlFileObject | Add-Member -NotePropertyName Filepath -NotePropertyValue $xamlFile.FullName
            
            [XML]$xamlFileContent = Get-Content $xamlFile.FullName
      
            # Get all activities from a XAML file
            $ns = @{sap2010="http://schemas.microsoft.com/netfx/2010/xaml/activities/presentation";ui="http://schemas.uipath.com/workflow/activities"}
            
            if ($xamlFileContent | Select-Xml -Xpath '//*[@sap2010:WorkflowViewState.IdRef]' -Namespace $ns | Select-Object Node)
            {            
                $activities = ($xamlFileContent | Select-Xml -Xpath '//*[@sap2010:WorkflowViewState.IdRef]' -Namespace $ns | Select-Object Node).Node.get_Name()
                $activityNodes = ($xamlFileContent | Select-Xml -Xpath '//*[@sap2010:WorkflowViewState.IdRef]' -Namespace $ns).Node
                
                foreach($node in $activityNodes)
                {
                    if ($node.Type -eq $null)
                    {
                        $node | Add-Member -NotePropertyName Type -NotePropertyValue $node.get_name()
                    }
                }
                
                $xamlFileObject | Add-Member -NotePropertyName Activities -NotePropertyValue $activities

                $activitiesWithSelectors = New-Object System.Collections.ArrayList

                $nodesWithSelectors = ($xamlFileContent | Select-Xml -Xpath '//*[@Selector]' -Namespace $ns).Node
                foreach($node in $nodesWithSelectors)
                {
                    $selector = $node.Selector
                    while (-Not ($node.Type) -and -Not ($node.DisplayName) -and $node -ne $null)
                    {
                        $node = $node.ParentNode
                    }

                    $type = $node.Type
                    $displayName = $node.DisplayName

                    $activity = New-Object PSObject
                    $activity | Add-Member -NotePropertyName Selector -NotePropertyValue $selector
                    $activity | Add-Member -NotePropertyName Type -NotePropertyValue $type
                    $activity | Add-Member -NotePropertyName DisplayName -NotePropertyValue $displayName

                    $activitiesWithSelectors.Add($activity) 1>$null

                    if (($activity.Selector -ne "{x:Null}"))
                    {
                        $selector = "<head>" + $activity.Selector + "</head>"
                        if (([xml]$selector).head.wnd -and ([xml]$selector).head.wnd.app)
                        {
                            $activity | Add-Member -NotePropertyName AppName -NotePropertyValue ([xml]$selector).head.wnd.app
                        }
                        elseif (([xml]$selector).head.html -and ([xml]$selector).head.html.title)
                        {
                            $activity | Add-Member -NotePropertyName HtmlTitle -NotePropertyValue ([xml]$selector).head.html.title
                        }
                    }
                }

                $openAppNodes = @(($xamlFileContent | Select-Xml -Xpath '//ui:OpenApplication' -Namespace $ns).Node | Select DisplayName,FileName,Arguments,Selector)
                $openBrowserNodes = @(($xamlFileContent | Select-Xml -Xpath '//ui:OpenBrowser' -Namespace $ns).Node | Select DisplayName,BrowserType,Url)
                

                $xamlFileObject | Add-Member -NotePropertyName ActivitiesWithSelectors -NotePropertyValue $activitiesWithSelectors
                $xamlFileObject | Add-Member -NotePropertyName OpenAppActivities -NotePropertyValue $openAppNodes
                $xamlFileObject | Add-Member -NotePropertyName OpenBrowserActivities -NotePropertyValue $openBrowserNodes
            }
        
            # Check if the project json is of a recent type, that includes the exact Dependencies
            if(-Not (Get-Member -inputobject $projectJsonContent -name "schemaVersion" -Membertype Properties))
            {
                # OLD Project Version - Dependencies are not in the Project Json file
                $projectJsonContent | Add-Member -NotePropertyName schemaVersion -NotePropertyValue "N/A"
                $projectJsonContent | Add-Member -NotePropertyName studioVersion -NotePropertyValue $projectJsonContent.version
            }

            $references = New-Object System.Collections.ArrayList    
            # Include References as there are no exact Dependencies
            foreach($reference in $xamlFileContent.Activity.'TextExpression.ReferencesForImplementation'.Collection.ChildNodes)
            {
                $references.Add($reference.InnerText) 1>$null
            }

            $xamlFileObject | Add-Member -NotePropertyName References -NotePropertyValue $references
        
            $xamlFilesObject.Add($xamlFileObject) 1>$null

            
        }
        catch
        {
            Write-Error "ERROR processing XAML file <$($xamlFile.FullName)>: $($Error[0])"
        }
    }

    $projectJsonContent | Add-Member -NotePropertyName XAMLFiles -NotePropertyValue $xamlFilesObject
    $projectJsonContent | Add-Member -NotePropertyName UniqueActivities -NotePropertyValue $xamlFilesObject.Activities | Select -unique
    $projectJsonContent | Add-Member -NotePropertyName UniqueReferences -NotePropertyValue $xamlFilesObject.References | Select -Unique

    return $projectJsonContent
}

####################### Execution Steps #######################

#$downloadedPackages = GetActivePackages $TenantConfig $LocalPackageFolder
$projectFiles = (Get-ChildItem -Path $LocalPackageFolder -Force -Recurse -Filter "project.json")

$allProjects = New-Object System.Collections.ArrayList

foreach($file in $projectFiles)
{
    $project = GetProjectInformation $file
    $allProjects.Add($project) 1>$null
}

# Output the summary to a Json file
$allProjects | ConvertTo-json -Depth 10 | Out-File "$LocalPackageFolder\output.json"

# Output only the name and Dependencies
# $allProjects | Select-Object -Property Name, Dependencies | ConvertTo-json -Depth 10 | Out-File $PSScriptRoot\"summary.json"

###############################################################


########## Sample commands you can run #########
# $allProjects | Select-Object -Property Name, Dependencies
# $allProjects.Dependencies | Format-List                    # lists all dependencies as objects
# $allProjects.dependenciesNames | select -Unique            # lists all unique dependency names

# $allProjects.UniqueActivities | select -Unique             # lists all unique activities
# $allProjects.XamlFiles.Activities | group | select Name, Count # lists all unique activities and their count
# $allProjects.XamlFiles.ActivitiesWithSelectors.AppName # lista all app names used in selectors

# $allProjects.studioVersion | select -Unique   # lists all unique Studio versions