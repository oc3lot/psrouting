#############################################################################################
#Date			Who					Comment													#
#-------------	------------------	------------------------------------------------------	#
#21Jul2016		Kelvin Hoyle		Created to quickly create document routing rules        #
#									                                    					#
#############################################################################################

#############################################################################################
#Parameters: $ContentType - cmdline parameter to capture Content Type           			#
#Return Value: N/A																			#
#Purpose: Inserts new document routing rules for specified Content Type                    	#
#############################################################################################


#Load Content Types
$configXMLPath = "C:\temp\ContentTypeListing.xml"

#test to ensure the config exists
if (!(Test-Path $configXMLPath))
{    
    Write-Host "Invalid configuration file path. Exiting..." -ForegroundColor Red
    exit
}

function Dispose-All {
    Get-Variable -exclude Runspace |
        Where-Object {
            $_.Value -is [System.IDisposable]
        } |
        Foreach-Object {
            $_.Value.Dispose()
            Remove-Variable $_.Name
        }
}

#store the config & global variables
$ctNameList = [xml](Get-Content $configXMLPath);

#run scripts
Write-Host "Now creating primary sorting rules..." -foregroundcolor White
Start-Sleep -s 5

foreach($xmlContentType in $ctNameList.ContentTypes.ContentType)
{
    .\PrimaryRoutingRules.ps1 $xmlContentType.Name
}

Write-Host "Successfully completed primary sorting rules" -foregroundcolor White
<#Start-Sleep -s 5


Write-Host "Now creating secondary sorting rules for Assessments Record Centre..." -foregroundcolor White
.\AssessmentRoutingRules.ps1
Dispose-All

Write-Host "Now creating secondary sorting rules for Claims Record Centre..." -foregroundcolor White
.\ClaimsRoutingRules.ps1
Dispose-All

Write-Host "Now creating secondary sorting rules for Finance Record Centre..." -foregroundcolor White
.\FinanceRoutingRules.ps1
Dispose-All

Write-Host "Now creating secondary sorting rules for Prevention Record Centre..." -foregroundcolor White
.\PrevRoutingRules.ps1
Dispose-All

Write-Host "Now creating secondary sorting rules for Content Record Centre..." -foregroundcolor White
foreach($xmlContentType in $ctNameList.ContentTypes.ContentType)
{
    Write-Host "     Creating rules for content type: " $xmlContentType.Name -foregroundcolor Green
    .\ContentRoutingRules.ps1 $xmlContentType.Name
}
Dispose-All
Write-Host "All Document Routing Rules have been successfully created." -foregroundcolor White#>