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

Param(
    [String]$ContentType
)

$configXMLPath = "C:\temp\RoutingRuleConfigTemplate.xml"

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
$mainURL = $ctConfig.ManageRules.WebURL
$RecordCenterURL = $ctConfig.ManageRules.RecordCentreURL

#creation of 'Document Centre' sorting rules
Write-Host "     Creating document routing rules for content type:" $ContentType -foregroundcolor Green

#loop through each Product Family (Non-Confidential Rules)
foreach($xmlProdFam in $ctConfig.ManageRules.ProductFamilies.ProductFamily)
{
    #Write-Host "Creating rules for" $xmlProdFam.Name "..."

    #retrieve Term ID for Product Family
    $PFValue = $xmlProdFam.Name
    $ts = Get-SPTaxonomySession -Site $ctConfig.ManageRules.WebURL
    $tstore = $ts.TermStores[0]
    $tgroup = $tstore.Groups[$ctConfig.ManageRules.TermGroup]
    $tset = $tgroup.TermSets[$ctConfig.ManageRules.TermSet]
    $term = $tset.GetTerms($PFValue, $true)
    $termValueGuid = $term.Id

   #construct full literal value of MMC
    $web = Get-SPWeb $mainURL
    $docLib = $web.Lists["Documents"]
    $appField = [Microsoft.SharePoint.Taxonomy.TaxonomyField]$docLib.Fields["Application"]
    [Microsoft.SharePoint.Taxonomy.TaxonomyFieldValue]$taxonomyFieldValue = New-Object Microsoft.SharePoint.Taxonomy.TaxonomyFieldValue($appField)
    $taxonomyFieldValue.PopulateFromLabelGuidPair([Microsoft.SharePoint.Taxonomy.TermSet]::NormalizeName($Application) + "|" + $termValueGuid)

    #Create Rule
    [Microsoft.SharePoint.SPSite]$site = Get-SPSite $mainURL
    [Microsoft.SharePoint.SPWeb]$web = Get-SPWeb $RecordCenterURL
    [Microsoft.SharePoint.SPContentType]$ct = $site.RootWeb.ContentTypes[$ContentType]
    [Microsoft.Office.RecordsManagement.RecordsRepository.EcmDocumentRouterRule]$rule = New-Object Microsoft.Office.RecordsManagement.RecordsRepository.EcmDocumentRouterRule($web)

    $rule.ConditionsString = "<Conditions><Condition Column='" + $appField.Id + "|Application|Application' Operator='EqualsOrIsAChildOf' Value='" + $taxonomyFieldValue.ValidatedString +  "'></Condition></Conditions>" 
    $rule.CustomRouter = ""
    $rule.Name = $xmlProdFam.Name + " " + $ct.Name + " Routing Rule"
    $rule.Description = "Routes '" + $ct.Name + "' documents from the '" + $xmlProdFam.Name + "' Product Family to the records library"
    $rule.ContentTypeString = $ct.Name
    $rule.RouteToExternalLocation = $true
    $rule.Priority = "5"
    $rule.TargetPath = $xmlProdFam.RoutingLocation
    $rule.enabled = $true
    $rule.Update()

}
