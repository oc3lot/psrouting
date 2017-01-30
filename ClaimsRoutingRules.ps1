#############################################################################################
#Date			Who					Comment													#
#-------------	------------------	------------------------------------------------------	#
#10Aug2016		Kelvin Hoyle		Created to quickly create document routing rules        #
#                                   for 'Claims' Record Library                             #
#									                                    					#
#############################################################################################

#############################################################################################
#Parameters: $ContentType - cmdline parameter to capture Content Type           			#
#Return Value: N/A																			#
#Purpose: Inserts new document routing rules for 'Claims' RC                       	#
#############################################################################################

$configXMLPath = "C:\temp\ContentTypeListing.xml"

#test to ensure the config exists
if (!(Test-Path $configXMLPath))
{    
    Write-Host "Invalid configuration file path. Exiting..." -ForegroundColor Red
    exit
}

#Load SharePoint PS Snap-In if necessary
if ((Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -eq $null) 
{
    Add-PSSnapin "Microsoft.SharePoint.PowerShell"
}

#store the config & global variables
$ctConfig = [xml](Get-Content $configXMLPath);
$RecordCenterURL = "http://infoctr/ic_claims/"
$web = Get-SPWeb $RecordCenterURL
$docLib = $web.Lists["Record Library"]
$conField = $docLib.Fields.GetField("Confidential Document")
$ppmField = $docLib.Fields.GetField("PPM ID")

#loop through each Content Type 
foreach($xmlContentType in $ctConfig.ContentTypes.ContentType)
{
    #Create Non-Confidential Rule
    [Microsoft.SharePoint.SPSite]$site = Get-SPSite $RecordCenterURL
    [Microsoft.SharePoint.SPWeb]$web = Get-SPWeb $RecordCenterURL
    [Microsoft.SharePoint.SPContentType]$ct = $site.RootWeb.ContentTypes[$xmlContentType]
    [Microsoft.Office.RecordsManagement.RecordsRepository.EcmDocumentRouterRule]$rule = New-Object Microsoft.Office.RecordsManagement.RecordsRepository.EcmDocumentRouterRule($web)

    $rule.ConditionsString = "<Conditions><Condition Column='" + $conField.Id + "|ConfDoc|Confidential Document' Operator='IsEqual' Value='False'></Condition></Conditions>" 
    $rule.CustomRouter = ""
    $rule.Name = "Non-Confidential " + $xmlContentType.Name + " Routing Rule"
    $rule.Description = ""
    $rule.ContentTypeString = $xmlContentType.Name
    $rule.RouteToExternalLocation = $false
    $rule.Priority = "5"
    $rule.TargetPath = $web.Lists["Record Library"].RootFolder.ServerRelativeUrl
    $rule.AutoFolderSettings.Enabled = $true
    $rule.AutoFolderSettings.AutoFolderPropertyName = $ppmField.InternalName
    $rule.AutoFolderSettings.AutoFolderPropertyId = $ppmField.Id
    $rule.AutoFolderSettings.AutoFolderFolderNameFormat = "%2"
    $rule.enabled = $true
    $rule.Update()

    #Create Confidential Rule
    [Microsoft.SharePoint.SPSite]$site = Get-SPSite $RecordCenterURL
    [Microsoft.SharePoint.SPWeb]$web = Get-SPWeb $RecordCenterURL
    [Microsoft.SharePoint.SPContentType]$ct = $site.RootWeb.ContentTypes[$xmlContentType]
    [Microsoft.Office.RecordsManagement.RecordsRepository.EcmDocumentRouterRule]$rule = New-Object Microsoft.Office.RecordsManagement.RecordsRepository.EcmDocumentRouterRule($web)

    $rule.ConditionsString = "<Conditions><Condition Column='" + $conField.Id + "|ConfDoc|Confidential Document' Operator='IsEqual' Value='True'></Condition></Conditions>" 
    $rule.CustomRouter = ""
    $rule.Name = "Confidential " + $xmlContentType.Name + " Routing Rule"
    $rule.Description = ""
    $rule.ContentTypeString = $xmlContentType.Name
    $rule.RouteToExternalLocation = $false
    $rule.Priority = "4"
    $rule.TargetPath = $web.Lists["Confidential Records"].RootFolder.ServerRelativeUrl
    $rule.AutoFolderSettings.Enabled = $true
    $rule.AutoFolderSettings.AutoFolderPropertyName = $ppmField.InternalName
    $rule.AutoFolderSettings.AutoFolderPropertyId = $ppmField.Id
    $rule.AutoFolderSettings.AutoFolderFolderNameFormat = "%2"
    $rule.enabled = $true
    $rule.Update()

}

