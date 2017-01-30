#############################################################################################
#Date			Who					Comment													#
#-------------	------------------	------------------------------------------------------	#
#25Sep2015		David Drever		Created to quickly create content types and site columns#
#									and to update existing content types					#
#############################################################################################

#############################################################################################
#Parameters: $configPath - Path to the configuration file for the process        			#
#Return Value: N/A																			#
#Purpose: Inserts new content tpes and site columns or updates existing content types    	#
#############################################################################################

Param(
    [String]$configXMLPath = "C:\temp\ContentTypeConfigTemplate.xml"
)

#*==================================================
#* Load Functions (all functions used in this master
#* batch are located in the files called below)
#*==================================================
. .\SPModifyContentFunctions.ps1


#test to ensure the config exists
if (!(Test-Path $configXMLPath))
{    
    Write-Host "Invalid configuration file path. Exiting..." -ForegroundColor Red
    exit
}

#store the config
$ctConfig = [xml](Get-Content $configXMLPath);


#loop through each new content type
foreach($xmlNewCT in $ctConfig.ManageContent.ContentTypes.InsertCTs.InsertCT)
{
    CreateContentType $ctConfig.ManageContent.webURL $xmlNewCT;

    #add new site columns if needed
    if($xmlNewCT.NewSiteColumns.SiteColumn)
    {
        foreach($xmlNewSiteColumn in $xmlNewCT.NewSiteColumns.SiteColumn)
        {
            switch ($xmlNewSiteColumn.Type)
            {
                "Lookup"
                {
                    CreateLookupSiteColumn $ctConfig.ManageContent.WebURL $xmlNewSiteColumn;
                    break;
                }

                "Choice"
                {
                    CreateChoiceColumn $ctConfig.ManageContent.WebURL $xmlNewSiteColumn;
                    break;
                }

                "TaxonomyFieldType"
                {
                    CreateMMSiteColumn $ctConfig.ManageContent.WebURL $xmlNewSiteColumn;
                    break;
                }
                default
                {
                    CreateSiteColumn $ctConfig.ManageContent.WebURL $xmlNewSiteColumn;
                }
            }

            AddSiteColumnToCT $ctConfig.ManageContent.webURL $xmlNewCT.Name $xmlNewSiteColumn.DisplayName;
        }        
    }

    #add any current site columns to CT
    if($xmlNewCT.ExistingSiteColumns.SiteColumn)
    {
        foreach($xmlExistingSiteColumn in $xmlNewCT.ExistingSiteColumns.SiteColumn)
        {
            AddSiteColumnToCT $ctConfig.ManageContent.webURL $xmlNewCT.Name $xmlExistingSiteColumn.DisplayName;
        }
    }


    #check if this content type requires a document templated attached.
    if(![string]::IsNullOrWhiteSpace($xmlNewCT.DocTemplateLocation))
    {
        AddDocTemplateToCT $ctConfig.ManageContent.webURL $xmlNewCT.Name $xmlNewCT.DocTemplateLocation $xmlNewCT.DocTemplateName;
    }

    Write-Host ("Created Content Type: {0} and added site columns." -f $xmlNewCT.Name) -ForegroundColor Green;

}

#loop through any content types to be updated
foreach($xmlUpdateCT in $ctConfig.ManageContent.ContentTypes.UpdateCTs.UpdateCT)
{
    #add new site columns if needed
    if($xmlUpdateCT.NewSiteColumns.SiteColumn)
    {
        foreach($xmlNewSiteColumn in $xmlUpdateCT.NewSiteColumns.SiteColumn)
        {
            switch ($xmlNewSiteColumn.Type)
            {
                "Lookup"
                {
                    CreateLookupSiteColumn $ctConfig.ManageContent.webURL $xmlNewSiteColumn;
                    break;
                }

                "Choice"
                {
                    CreateChoiceColumn $ctConfig.ManageContent.webURL $xmlNewSiteColumn;
                    break;
                }

                "TaxonomyFieldType"
                {
                    CreateMMSiteColumn $ctConfig.ManageContent.webURL $xmlNewSiteColumn;
                    break;
                }
                default
                {
                    CreateSiteColumn $ctConfig.ManageContent.WebURL $xmlNewSiteColumn;
                }
            }

            AddSiteColumnToCT $ctConfig.ManageContent.webURL $xmlUpdateCT.Name $xmlNewSiteColumn.DisplayName;
        }
    }

    #add any current site columns to CT
    if($xmlUpdateCT.ExistingSiteColumns.SiteColumn)
    {
        foreach($xmlExistingSiteColumn in $xmlUpdateCT.ExistingSiteColumns.SiteColumn)
        {
            AddSiteColumnToCT $ctConfig.ManageContent.webURL $xmlUpdateCT.Name $xmlNewSiteColumn.DisplayName;
        }
    }

    Write-Host ("Updated Content Type: {0} and added site columns." -f $xmlUpdateCT.Name) -ForegroundColor Green;

    #publish content type
    Write-Host ("Now publishing Content Types") -ForegroundColor Green;
    $site = Get-SPSite $ctConfig.ManageContent.webURL
    $group = $ctConfig.ManageContent.CTGroup
    $contentTypePublisher = New-Object Microsoft.SharePoint.Taxonomy.ContentTypeSync.ContentTypePublisher ($site)
    $site.RootWeb.ContentTypes | ? {$_.Group -match $Group} | % {
        $contentTypePublisher.Publish($_)
            write-host "Content type" $_.Name "has been republished" -foregroundcolor Green
    }
}

