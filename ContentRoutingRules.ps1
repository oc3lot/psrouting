#############################################################################################
#Date			Who					Comment													#
#-------------	------------------	------------------------------------------------------	#
#10Aug2016		Kelvin Hoyle		Created to quickly create document routing rules        #
#                                   for 'Content' Record Library                            #
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
$RecordCenterURL =  "http://infoctr/ic_content/"

#loop through each Product Family (Non-Confidential Rules)
foreach($xmlProdFam in $ctConfig.ManageRules.ProductFamilies.ProductFamily)
{
    #retrieve Term ID for Product Family
    $PFValue = $xmlProdFam.Name
    $ts = Get-SPTaxonomySession -Site $ctConfig.ManageRules.WebURL
    $tstore = $ts.TermStores[0]
    $tgroup = $tstore.Groups[$ctConfig.ManageRules.TermGroup]
    $tset = $tgroup.TermSets[$ctConfig.ManageRules.TermSet]
    $term = $tset.GetTerms($PFValue, $true)
    $termValueGuid = $term.Id

    #construct full literal value of MMC
    $web = Get-SPWeb $RecordCenterURL
    $docLib = $web.Lists["Record Library"]
    $appField = [Microsoft.SharePoint.Taxonomy.TaxonomyField]$docLib.Fields["Application"]
    [Microsoft.SharePoint.Taxonomy.TaxonomyFieldValue]$taxonomyFieldValue = New-Object Microsoft.SharePoint.Taxonomy.TaxonomyFieldValue($appField)
    $taxonomyFieldValue.PopulateFromLabelGuidPair([Microsoft.SharePoint.Taxonomy.TermSet]::NormalizeName($Application) + "|" + $termValueGuid)

    #Create Rule
    [Microsoft.SharePoint.SPSite]$site = Get-SPSite $mainURL
    [Microsoft.SharePoint.SPWeb]$web = Get-SPWeb $RecordCenterURL
    [Microsoft.SharePoint.SPContentType]$ct = $site.RootWeb.ContentTypes[$ContentType]
    $docLib = $web.Lists["Record Library"]
    $conField = $docLib.Fields.GetField("Confidential Document")
    $ppmField = $docLib.Fields.GetField("PPM ID")
    [Microsoft.Office.RecordsManagement.RecordsRepository.EcmDocumentRouterRule]$rule = New-Object Microsoft.Office.RecordsManagement.RecordsRepository.EcmDocumentRouterRule($web)

    if ($PFValue -eq "Business Intelligence") {
        $rule.ConditionsString = "<Conditions><Condition Column='" + $appField.Id + "|Application|Application' Operator='EqualsOrIsAChildOf' Value='" + $taxonomyFieldValue.ValidatedString +  "'></Condition><Condition Column='" + $conField.Id + "|ConfDoc|Confidential Document' Operator='IsEqual' Value='False'></Condition></Conditions>" 
        $rule.CustomRouter = ""
        $rule.Name = $xmlProdFam.Name + " " + $ct.Name + " Non-Confidential Routing Rule"
        $rule.Description = ""
        $rule.ContentTypeString = $ct.name
        $rule.RouteToExternalLocation = $false
        $rule.Priority = "5"
        $rule.TargetPath = $web.Lists["BI Record Centre"].RootFolder.ServerRelativeUrl
        $rule.AutoFolderSettings.Enabled = $true
        $rule.AutoFolderSettings.AutoFolderPropertyName = $ppmField.InternalName
        $rule.AutoFolderSettings.AutoFolderPropertyId = $ppmField.Id
        $rule.AutoFolderSettings.AutoFolderFolderNameFormat = "%2" 
        $rule.enabled = $true
        $rule.Update()
    }
    Elseif ($PFValue -eq "Corporate Services") {
       $rule.ConditionsString = "<Conditions><Condition Column='" + $appField.Id + "|Application|Application' Operator='EqualsOrIsAChildOf' Value='" + $taxonomyFieldValue.ValidatedString +  "'></Condition><Condition Column='" + $conField.Id + "|ConfDoc|Confidential Document' Operator='IsEqual' Value='False'></Condition></Conditions>" 
        $rule.CustomRouter = ""
        $rule.Name = $xmlProdFam.Name + " " + $ct.Name + " Non-Confidential Routing Rule"
        $rule.Description = ""
        $rule.ContentTypeString = $ct.name
        $rule.RouteToExternalLocation = $false
        $rule.Priority = "5"
        $rule.TargetPath = $web.Lists["Corporate Services Record Centre"].RootFolder.ServerRelativeUrl
        $rule.AutoFolderSettings.Enabled = $true
        $rule.AutoFolderSettings.AutoFolderPropertyName = $ppmField.InternalName
        $rule.AutoFolderSettings.AutoFolderPropertyId = $ppmField.Id
        $rule.AutoFolderSettings.AutoFolderFolderNameFormat = "%2" 
        $rule.enabled = $true
        $rule.Update() 
    }
    Elseif ($PFValue -eq "Healthcare") {
       $rule.ConditionsString = "<Conditions><Condition Column='" + $appField.Id + "|Application|Application' Operator='EqualsOrIsAChildOf' Value='" + $taxonomyFieldValue.ValidatedString +  "'></Condition><Condition Column='" + $conField.Id + "|ConfDoc|Confidential Document' Operator='IsEqual' Value='False'></Condition></Conditions>" 
        $rule.CustomRouter = ""
        $rule.Name = $xmlProdFam.Name + " " + $ct.Name + " Non-Confidential Routing Rule"
        $rule.Description = ""
        $rule.ContentTypeString = $ct.name
        $rule.RouteToExternalLocation = $false
        $rule.Priority = "5"
        $rule.TargetPath = $web.Lists["Healthcare Record Centre"].RootFolder.ServerRelativeUrl
        $rule.AutoFolderSettings.Enabled = $true
        $rule.AutoFolderSettings.AutoFolderPropertyName = $ppmField.InternalName
        $rule.AutoFolderSettings.AutoFolderPropertyId = $ppmField.Id
        $rule.AutoFolderSettings.AutoFolderFolderNameFormat = "%2" 
        $rule.enabled = $true
        $rule.Update() 
    }
    Elseif ($PFValue -eq "HR") {
       $rule.ConditionsString = "<Conditions><Condition Column='" + $appField.Id + "|Application|Application' Operator='EqualsOrIsAChildOf' Value='" + $taxonomyFieldValue.ValidatedString +  "'></Condition><Condition Column='" + $conField.Id + "|ConfDoc|Confidential Document' Operator='IsEqual' Value='False'></Condition></Conditions>" 
        $rule.CustomRouter = ""
        $rule.Name = $xmlProdFam.Name + " " + $ct.Name + " Non-Confidential Routing Rule"
        $rule.Description = ""
        $rule.ContentTypeString = $ct.name
        $rule.RouteToExternalLocation = $false
        $rule.Priority = "5"
        $rule.TargetPath = $web.Lists["HR Record Centre"].RootFolder.ServerRelativeUrl
        $rule.AutoFolderSettings.Enabled = $true
        $rule.AutoFolderSettings.AutoFolderPropertyName = $ppmField.InternalName
        $rule.AutoFolderSettings.AutoFolderPropertyId = $ppmField.Id
        $rule.AutoFolderSettings.AutoFolderFolderNameFormat = "%2" 
        $rule.enabled = $true
        $rule.Update() 
    }
    Elseif ($PFValue -eq "IT") {
       $rule.ConditionsString = "<Conditions><Condition Column='" + $appField.Id + "|Application|Application' Operator='EqualsOrIsAChildOf' Value='" + $taxonomyFieldValue.ValidatedString +  "'></Condition><Condition Column='" + $conField.Id + "|ConfDoc|Confidential Document' Operator='IsEqual' Value='False'></Condition></Conditions>" 
        $rule.CustomRouter = ""
        $rule.Name = $xmlProdFam.Name + " " + $ct.Name + " Non-Confidential Routing Rule"
        $rule.Description = ""
        $rule.ContentTypeString = $ct.name
        $rule.RouteToExternalLocation = $false
        $rule.Priority = "5"
        $rule.TargetPath = $web.Lists["IT Record Centre"].RootFolder.ServerRelativeUrl
        $rule.AutoFolderSettings.Enabled = $true
        $rule.AutoFolderSettings.AutoFolderPropertyName = $ppmField.InternalName
        $rule.AutoFolderSettings.AutoFolderPropertyId = $ppmField.Id
        $rule.AutoFolderSettings.AutoFolderFolderNameFormat = "%2" 
        $rule.enabled = $true
        $rule.Update() 
    }
    Elseif ($PFValue -eq "Cross Cutting") {
       $rule.ConditionsString = "<Conditions><Condition Column='" + $appField.Id + "|Application|Application' Operator='EqualsOrIsAChildOf' Value='" + $taxonomyFieldValue.ValidatedString +  "'></Condition><Condition Column='" + $conField.Id + "|ConfDoc|Confidential Document' Operator='IsEqual' Value='False'></Condition></Conditions>" 
        $rule.CustomRouter = ""
        $rule.Name = $xmlProdFam.Name + " " + $ct.Name + " Non-Confidential Routing Rule"
        $rule.Description = ""
        $rule.ContentTypeString = $ct.name
        $rule.RouteToExternalLocation = $false
        $rule.Priority = "5"
        $rule.TargetPath = $web.Lists["Cross Cutting Record Centre"].RootFolder.ServerRelativeUrl
        $rule.AutoFolderSettings.Enabled = $true
        $rule.AutoFolderSettings.AutoFolderPropertyName = $ppmField.InternalName
        $rule.AutoFolderSettings.AutoFolderPropertyId = $ppmField.Id
        $rule.AutoFolderSettings.AutoFolderFolderNameFormat = "%2" 
        $rule.enabled = $true
        $rule.Update() 
    }
    Elseif ($PFValue -eq "Marketing ＆ Communications") {
       $rule.ConditionsString = "<Conditions><Condition Column='" + $appField.Id + "|Application|Application' Operator='EqualsOrIsAChildOf' Value='" + $taxonomyFieldValue.ValidatedString +  "'></Condition><Condition Column='" + $conField.Id + "|ConfDoc|Confidential Document' Operator='IsEqual' Value='False'></Condition></Conditions>" 
        $rule.CustomRouter = ""
        $rule.Name = $xmlProdFam.Name + " " + $ct.Name + " Non-Confidential Routing Rule"
        $rule.Description = ""
        $rule.ContentTypeString = $ct.name
        $rule.RouteToExternalLocation = $false
        $rule.Priority = "5"
        $rule.TargetPath = "/sites/contentrc/mktgcomm"
        $rule.AutoFolderSettings.Enabled = $true
        $rule.AutoFolderSettings.AutoFolderPropertyName = $ppmField.InternalName
        $rule.AutoFolderSettings.AutoFolderPropertyId = $ppmField.Id
        $rule.AutoFolderSettings.AutoFolderFolderNameFormat = "%2" 
        $rule.enabled = $true
        $rule.Update() 
    }
    Elseif ($PFValue -eq "Reviews ＆ Appeals") {
       $rule.ConditionsString = "<Conditions><Condition Column='" + $appField.Id + "|Application|Application' Operator='EqualsOrIsAChildOf' Value='" + $taxonomyFieldValue.ValidatedString +  "'></Condition><Condition Column='" + $conField.Id + "|ConfDoc|Confidential Document' Operator='IsEqual' Value='False'></Condition></Conditions>" 
        $rule.CustomRouter = ""
        $rule.Name = $xmlProdFam.Name + " " + $ct.Name + " Non-Confidential Routing Rule"
        $rule.Description = ""
        $rule.ContentTypeString = $ct.name
        $rule.RouteToExternalLocation = $false
        $rule.Priority = "5"
        $rule.TargetPath = "/sites/contentrc/review"
        $rule.AutoFolderSettings.Enabled = $true
        $rule.AutoFolderSettings.AutoFolderPropertyName = $ppmField.InternalName
        $rule.AutoFolderSettings.AutoFolderPropertyId = $ppmField.Id
        $rule.AutoFolderSettings.AutoFolderFolderNameFormat = "%2" 
        $rule.enabled = $true
        $rule.Update() 
    }
}
    #loop through each Product Family (Confidential Rules)
foreach($xmlProdFam in $ctConfig.ManageRules.ProductFamilies.ProductFamily)
{
    #retrieve Term ID for Product Family
    $PFValue = $xmlProdFam.Name
    $ts = Get-SPTaxonomySession -Site $ctConfig.ManageRules.WebURL
    $tstore = $ts.TermStores[0]
    $tgroup = $tstore.Groups[$ctConfig.ManageRules.TermGroup]
    $tset = $tgroup.TermSets[$ctConfig.ManageRules.TermSet]
    $term = $tset.GetTerms($PFValue, $true)
    $termValueGuid = $term.Id

    #construct full literal value of MMC
    $web = Get-SPWeb $RecordCenterURL
    $docLib = $web.Lists["Record Library"]
    $appField = [Microsoft.SharePoint.Taxonomy.TaxonomyField]$docLib.Fields["Application"]
    [Microsoft.SharePoint.Taxonomy.TaxonomyFieldValue]$taxonomyFieldValue = New-Object Microsoft.SharePoint.Taxonomy.TaxonomyFieldValue($appField)
    $taxonomyFieldValue.PopulateFromLabelGuidPair([Microsoft.SharePoint.Taxonomy.TermSet]::NormalizeName($Application) + "|" + $termValueGuid)

    #Create Rule
    [Microsoft.SharePoint.SPSite]$site = Get-SPSite $mainURL
    [Microsoft.SharePoint.SPWeb]$web = Get-SPWeb $RecordCenterURL
    [Microsoft.SharePoint.SPContentType]$ct = $site.RootWeb.ContentTypes[$ContentType]
    $docLib = $web.Lists["Record Library"]
    $conField = $docLib.Fields.GetField("Confidential Document")
    $ppmField = $docLib.Fields.GetField("PPM ID")
    [Microsoft.Office.RecordsManagement.RecordsRepository.EcmDocumentRouterRule]$rule = New-Object Microsoft.Office.RecordsManagement.RecordsRepository.EcmDocumentRouterRule($web)

    if ($PFValue -eq "Business Intelligence") {
        $rule.ConditionsString = "<Conditions><Condition Column='" + $appField.Id + "|Application|Application' Operator='EqualsOrIsAChildOf' Value='" + $taxonomyFieldValue.ValidatedString +  "'></Condition><Condition Column='" + $conField.Id + "|ConfDoc|Confidential Document' Operator='IsEqual' Value='True'></Condition></Conditions>" 
        $rule.CustomRouter = ""
        $rule.Name = $xmlProdFam.Name + " " + $ct.Name + " Confidential Routing Rule"
        $rule.Description = ""
        $rule.ContentTypeString = $ct.name
        $rule.RouteToExternalLocation = $false
        $rule.Priority = "4"
        $rule.TargetPath = $web.Lists["BI Record Centre - Confidential"].RootFolder.ServerRelativeUrl
        $rule.AutoFolderSettings.Enabled = $true
        $rule.AutoFolderSettings.AutoFolderPropertyName = $ppmField.InternalName
        $rule.AutoFolderSettings.AutoFolderPropertyId = $ppmField.Id
        $rule.AutoFolderSettings.AutoFolderFolderNameFormat = "%2" 
        $rule.enabled = $true
        $rule.Update()
    }
    Elseif ($PFValue -eq "Corporate Services") {
       $rule.ConditionsString = "<Conditions><Condition Column='" + $appField.Id + "|Application|Application' Operator='EqualsOrIsAChildOf' Value='" + $taxonomyFieldValue.ValidatedString +  "'></Condition><Condition Column='" + $conField.Id + "|ConfDoc|Confidential Document' Operator='IsEqual' Value='True'></Condition></Conditions>" 
        $rule.CustomRouter = ""
        $rule.Name = $xmlProdFam.Name + " " + $ct.Name + " Confidential Routing Rule"
        $rule.Description = ""
        $rule.ContentTypeString = $ct.name
        $rule.RouteToExternalLocation = $false
        $rule.Priority = "4"
        $rule.TargetPath = $web.Lists["Corporate Services Record Centre - Confidential"].RootFolder.ServerRelativeUrl
        $rule.AutoFolderSettings.Enabled = $true
        $rule.AutoFolderSettings.AutoFolderPropertyName = $ppmField.InternalName
        $rule.AutoFolderSettings.AutoFolderPropertyId = $ppmField.Id
        $rule.AutoFolderSettings.AutoFolderFolderNameFormat = "%2" 
        $rule.enabled = $true
        $rule.Update() 
    }
    Elseif ($PFValue -eq "Healthcare") {
       $rule.ConditionsString = "<Conditions><Condition Column='" + $appField.Id + "|Application|Application' Operator='EqualsOrIsAChildOf' Value='" + $taxonomyFieldValue.ValidatedString +  "'></Condition><Condition Column='" + $conField.Id + "|ConfDoc|Confidential Document' Operator='IsEqual' Value='True'></Condition></Conditions>" 
        $rule.CustomRouter = ""
        $rule.Name = $xmlProdFam.Name + " " + $ct.Name + " Confidential Routing Rule"
        $rule.Description = ""
        $rule.ContentTypeString = $ct.name
        $rule.RouteToExternalLocation = $false
        $rule.Priority = "4"
        $rule.TargetPath = $web.Lists["Healthcare Record Centre - Confidential"].RootFolder.ServerRelativeUrl
        $rule.AutoFolderSettings.Enabled = $true
        $rule.AutoFolderSettings.AutoFolderPropertyName = $ppmField.InternalName
        $rule.AutoFolderSettings.AutoFolderPropertyId = $ppmField.Id
        $rule.AutoFolderSettings.AutoFolderFolderNameFormat = "%2" 
        $rule.enabled = $true
        $rule.Update() 
    }
    Elseif ($PFValue -eq "HR") {
       $rule.ConditionsString = "<Conditions><Condition Column='" + $appField.Id + "|Application|Application' Operator='EqualsOrIsAChildOf' Value='" + $taxonomyFieldValue.ValidatedString +  "'></Condition><Condition Column='" + $conField.Id + "|ConfDoc|Confidential Document' Operator='IsEqual' Value='True'></Condition></Conditions>" 
        $rule.CustomRouter = ""
        $rule.Name = $xmlProdFam.Name + " " + $ct.Name + " Confidential Routing Rule"
        $rule.Description = ""
        $rule.ContentTypeString = $ct.name
        $rule.RouteToExternalLocation = $false
        $rule.Priority = "4"
        $rule.TargetPath = $web.Lists["HR Record Centre - Confidential"].RootFolder.ServerRelativeUrl
        $rule.AutoFolderSettings.Enabled = $true
        $rule.AutoFolderSettings.AutoFolderPropertyName = $ppmField.InternalName
        $rule.AutoFolderSettings.AutoFolderPropertyId = $ppmField.Id
        $rule.AutoFolderSettings.AutoFolderFolderNameFormat = "%2" 
        $rule.enabled = $true
        $rule.Update() 
    }
    Elseif ($PFValue -eq "IT") {
       $rule.ConditionsString = "<Conditions><Condition Column='" + $appField.Id + "|Application|Application' Operator='EqualsOrIsAChildOf' Value='" + $taxonomyFieldValue.ValidatedString +  "'></Condition><Condition Column='" + $conField.Id + "|ConfDoc|Confidential Document' Operator='IsEqual' Value='True'></Condition></Conditions>" 
        $rule.CustomRouter = ""
        $rule.Name = $xmlProdFam.Name + " " + $ct.Name + " Confidential Routing Rule"
        $rule.Description = ""
        $rule.ContentTypeString = $ct.name
        $rule.RouteToExternalLocation = $false
        $rule.Priority = "4"
        $rule.TargetPath = $web.Lists["IT Record Centre - Confidential"].RootFolder.ServerRelativeUrl
        $rule.AutoFolderSettings.Enabled = $true
        $rule.AutoFolderSettings.AutoFolderPropertyName = $ppmField.InternalName
        $rule.AutoFolderSettings.AutoFolderPropertyId = $ppmField.Id
        $rule.AutoFolderSettings.AutoFolderFolderNameFormat = "%2" 
        $rule.enabled = $true
        $rule.Update() 
    }
    Elseif ($PFValue -eq "Cross Cutting") {
       $rule.ConditionsString = "<Conditions><Condition Column='" + $appField.Id + "|Application|Application' Operator='EqualsOrIsAChildOf' Value='" + $taxonomyFieldValue.ValidatedString +  "'></Condition><Condition Column='" + $conField.Id + "|ConfDoc|Confidential Document' Operator='IsEqual' Value='True'></Condition></Conditions>" 
        $rule.CustomRouter = ""
        $rule.Name = $xmlProdFam.Name + " " + $ct.Name + " Confidential Routing Rule"
        $rule.Description = ""
        $rule.ContentTypeString = $ct.name
        $rule.RouteToExternalLocation = $false
        $rule.Priority = "4"
        $rule.TargetPath = $web.Lists["Cross Cutting Record Centre - Confidential"].RootFolder.ServerRelativeUrl
        $rule.AutoFolderSettings.Enabled = $true
        $rule.AutoFolderSettings.AutoFolderPropertyName = $ppmField.InternalName
        $rule.AutoFolderSettings.AutoFolderPropertyId = $ppmField.Id
        $rule.AutoFolderSettings.AutoFolderFolderNameFormat = "%2" 
        $rule.enabled = $true
        $rule.Update() 
    }
    Elseif ($PFValue -eq "Marketing ＆ Communications") {
       $rule.ConditionsString = "<Conditions><Condition Column='" + $appField.Id + "|Application|Application' Operator='EqualsOrIsAChildOf' Value='" + $taxonomyFieldValue.ValidatedString +  "'></Condition><Condition Column='" + $conField.Id + "|ConfDoc|Confidential Document' Operator='IsEqual' Value='True'></Condition></Conditions>" 
        $rule.CustomRouter = ""
        $rule.Name = $xmlProdFam.Name + " " + $ct.Name + " Confidential Routing Rule"
        $rule.Description = ""
        $rule.ContentTypeString = $ct.name
        $rule.RouteToExternalLocation = $false
        $rule.Priority = "4"
        $rule.TargetPath = "/sites/contentrc/mktgcommConf"
        $rule.AutoFolderSettings.Enabled = $true
        $rule.AutoFolderSettings.AutoFolderPropertyName = $ppmField.InternalName
        $rule.AutoFolderSettings.AutoFolderPropertyId = $ppmField.Id
        $rule.AutoFolderSettings.AutoFolderFolderNameFormat = "%2" 
        $rule.enabled = $true
        $rule.Update() 
    }
    Elseif ($PFValue -eq "Reviews ＆ Appeals") {
       $rule.ConditionsString = "<Conditions><Condition Column='" + $appField.Id + "|Application|Application' Operator='EqualsOrIsAChildOf' Value='" + $taxonomyFieldValue.ValidatedString +  "'></Condition><Condition Column='" + $conField.Id + "|ConfDoc|Confidential Document' Operator='IsEqual' Value='True'></Condition></Conditions>" 
        $rule.CustomRouter = ""
        $rule.Name = $xmlProdFam.Name + " " + $ct.Name + " Confidential Routing Rule"
        $rule.Description = ""
        $rule.ContentTypeString = $ct.name
        $rule.RouteToExternalLocation = $false
        $rule.Priority = "4"
        $rule.TargetPath = "/sites/contentrc/reviewConf"
        $rule.AutoFolderSettings.Enabled = $true
        $rule.AutoFolderSettings.AutoFolderPropertyName = $ppmField.InternalName
        $rule.AutoFolderSettings.AutoFolderPropertyId = $ppmField.Id
        $rule.AutoFolderSettings.AutoFolderFolderNameFormat = "%2" 
        $rule.enabled = $true
        $rule.Update() 
    }


}

