#############################################################################################
#Date			Who					Comment													#
#-------------	------------------	------------------------------------------------------	#
#29Sep2015		David Drever		Contains the methods and functions used for modifying   #
#									content within SharePoint            					#
#############################################################################################

#############################################################################################
#Function Name: CreateSiteColumn               				     							#
#Parameters: $siteURL - URL where subsite is to be placed 									#
#            $xmlNewSiteColumn - xml Object containing all the data for new site column     #
#                                                                                           #
#Return Value: N/A																			#
#Purpose: Creates a new site column in a site                 								#
#############################################################################################

Function CreateSiteColumn($webURL, $xmlNewSiteColumn)
{
    try
    {
        $spWeb = Get-SPWeb $webURL;

        #check to see if the column exists already
        if(!$spWeb.Fields[$xmlNewSiteColumn.DisplayName])
        {

            [string]$newColumnDef = ("<Field `
                                        Name='{0}' `
                                        DisplayName='{1}' `
                                        Description='{2}' `
                                        Type='{3}' `
                                        Required='{4}' `
                                        Group='{5}' >`
                                      </Field>" -f `
                               $xmlNewSiteColumn.InternalName,$xmlNewSiteColumn.DisplayName,$xmlNewSiteColumn. `
                               Description,$xmlNewSiteColumn.Type,$xmlNewSiteColumn.Required,$xmlNewSiteColumn.SCGroup);
            
            #add the site column
            $spWeb.Fields.AddFieldAsXml($newColumnDef) | Out-Null;

            if(![string]::IsNullOrWhiteSpace($xmlNewSiteColumn.DefaultValue))
            {
                $spCType = $spWeb.Fields.GetField[$xmlNewSiteColumn.InternalName];
                $spCType.DefaultValue = $xmlNewSiteColumn.DefaultValue;
                $spCType.Update()
            }
        }
    }
    catch
    {
        $ErrorMessage = $_.Exception.Message;
        Write-Host ("{0}: An error occurred creating the site column: {1}.  Error received: {2}" -f `
                        (get-date).ToString("yyyy-MM-dd HH:mm:ss"), $xmlNewSiteColumn.DisplayName, $ErrorMessage) -ForegroundColor Red
    }
    finally
    {
        $spWeb.Dispose();
    }
}


#############################################################################################
#Function Name: CreateMMSiteColumn               											#
#Parameters: $siteURL - URL where subsite is to be placed 									#
#            $xmlNewSiteColumn - xml Object containing all the data for new site column     #
#                                                                                           #
#Return Value: N/A																			#
#Purpose: Creates a new Taxonomy field (site column) in a site								#
#############################################################################################

Function CreateMMSiteColumn($webURL, $xmlNewSiteColumn)
{
    try
    {
        $spWeb = Get-SPWeb $webURL;

        #check to see if the column exists already
        if(!$spWeb.Fields[$xmlNewSiteColumn.DisplayName])
        {
            [string]$newColumnDef = ("<Field `
                                        Name='{0}' `
                                        DisplayName='{1}' `
                                        Description='{2}' `
                                        Type='{3}' `
                                        ShowField='Term1033' `
                                        Required='{4}' `
                                        Group='{5}' >`
                                      </Field>" -f `
                               $xmlNewSiteColumn.InternalName,$xmlNewSiteColumn.DisplayName,$xmlNewSiteColumn. `
                               Description,$xmlNewSiteColumn.Type,$xmlNewSiteColumn.Required,$xmlNewSiteColumn.SCGroup);            
            
            #add the site column
            $spWeb.Fields.AddFieldAsXml($newColumnDef) | Out-Null;

            #create a connection to MMS and link column
            $taxonomySession = New-Object Microsoft.SharePoint.Taxonomy.TaxonomySession($spWeb.Site);
            $termStore = $taxonomySession.TermStores[0];
            $termGroup = $termStore.Group[$xmlNewSiteColumn.TermGroup];
            $termSet = $termGroup.TermSets[$xmlNewSiteColumn.TermSet];

            #create new taxonomy field from term store connection
            [Microsoft.SharePoint.Taxonomy.TaxonomyField]$newTaxField = `
                [Microsoft.SharePoint.Taxonomy.TaxonomyField]$ctHubSite.Fields.GetFieldByInternalName($xmlNewSiteColumn.InternalName);

            #configure the new field
            $newTaxField.SspId = $termSet.TermStore.Id;
            $newTaxField.TermSetId = $termSet.Id;
            $newTaxField.TargetTemplate = [System.string]::Empty;
            $newTaxField.AnchorId = [System.GUID]::Empty;

            #update the field
            $newTaxField.Update();
        }
    }
    catch
    {
        $ErrorMessage = $_.Exception.Message;
        Write-Host ("{0}: An error occurred creating the Taxonomy site column: {1}.  Error received: {2}" -f `
                        (get-date).ToString("yyyy-MM-dd HH:mm:ss"), $xmlNewSiteColumn.DisplayName, $ErrorMessage) -ForegroundColor Red
    }
    finally
    {
        $spWeb.Dispose();
    }
}


#############################################################################################
#Function Name: CreateChoiceColumn               											#
#Parameters: $siteURL - URL where subsite is to be placed 									#
#            $xmlNewSiteColumn - xml Object containing all the data for new site column     #
#                                                                                           #
#Return Value: N/A																			#
#Purpose: Creates a new Choice site column in a site        								#
#############################################################################################

Function CreateChoiceColumn($webURL, $xmlNewSiteColumn)
{
    try
    {
        $spWeb = Get-SPWeb $webURL;

        #check to see if the column exists already
        if(!$spWeb.Fields[$xmlNewSiteColumn.DisplayName])
        {
            [string]$newColumnDef = ("<Field `
                                        Name='{0}' `
                                        DisplayName='{1}' `
                                        Description='{2}' `
                                        Type='{3}' `
                                        Required='{4}' `
                                        Group='{5}' >`
                                      </Field>" -f `
                               $xmlNewSiteColumn.InternalName,$xmlNewSiteColumn.DisplayName,$xmlNewSiteColumn. `
                               Description,$xmlNewSiteColumn.Type,$xmlNewSiteColumn.Required,$xmlNewSiteColumn.SCGroup);
            
            #add the site column
            $spWeb.Fields.AddFieldAsXml($newColumnDef) | Out-Null;

            #get the list of choices from the config and add to array (if more than one exist).
            $choiceList = $xmlNewSiteColumn.Choices;

            #check to see if there is more than one choice
            if($choiceList -like "*;*")
            {
                $choiceField = $spWeb.Fields.GetField($xmlNewSiteColumn.InternalName);

                $list = @($listOfChoices);
                
                #place the list of choices into an array and then loop through to add to field
                $listOfChoices = $choiceList.split(";");

                foreach($choice in $listOfChoices)
                {
                    $choiceField.Choices.Add($choice) | Out-Null;
                }
                
                $choiceField.Update();
            }
            else
            {
                $choiceField = $spWeb.Fields.GetField($xmlNewSiteColumn.InternalName);
                $choiceField.Choices.Add($choiceList) | Out-Null;

                $choiceField.Update();
            }

        }
    }
    catch
    {
        $ErrorMessage = $_.Exception.Message;
        Write-Host ("{0}: An error occurred creating the Choice site column: {1}.  Error received: {2}" -f `
                        (get-date).ToString("yyyy-MM-dd HH:mm:ss"), $xmlNewSiteColumn.DisplayName, $ErrorMessage) -ForegroundColor Red
    }
    finally
    {
        $spWeb.Dispose();
    }
}


#############################################################################################
#Function Name: CreateLookupSiteColumn             											#
#Parameters: $siteURL - URL where subsite is to be placed 									#
#            $xmlNewSiteColumn - xml Object containing all the data for new site column     #
#                                                                                           #
#Return Value: N/A																			#
#Purpose: Creates a new Lookup site column in a site                 						#
#############################################################################################

Function CreateLookupSiteColumn($webURL, $xmlNewSiteColumn)
{
    try
    {
        $spWeb = Get-SPWeb $webURL;

        #first check to see if the list exists
        $spList = $spWeb.Lists.TryGetList($xmlNewSiteColumn.LookupListInternal);
        
        if(!$spList)
        {
            CreateLookupList $spWeb $xmlNewSiteColumn.LookupListInternal $xmlNewSiteColumn.LookupList;
        }


        #check to see if the column exists already
        if(!$spWeb.Fields[$xmlNewSiteColumn.DisplayName])
        {

            [string]$newColumnDef = ("<Field `
                                        Name='{0}' `
                                        DisplayName='{1}' `
                                        Description='{2}' `
                                        Type='{3}' `
                                        Required='{4}' `
                                        Group='{5}' `
                                        List='{6}' > `
                                      </Field>" -f `
                               $xmlNewSiteColumn.InternalName,$xmlNewSiteColumn.DisplayName,$xmlNewSiteColumn. `
                               Description,$xmlNewSiteColumn.Type,$xmlNewSiteColumn.Required,$xmlNewSiteColumn.SCGroup, $spWeb.Lists[$xmlNewSiteColumn.LookupList].ID);
               
            #add the site column
            $spWeb.Fields.AddFieldAsXml($newColumnDef) | Out-Null;
        }
    }
    catch
    {
        $ErrorMessage = $_.Exception.Message;
        Write-Host ("{0}: An error occurred creating the Lookup site column: {1}.  Error received: {2}" -f `
                        (get-date).ToString("yyyy-MM-dd HH:mm:ss"), $xmlNewSiteColumn.DisplayName, $ErrorMessage) -ForegroundColor Red
    }
    finally
    {
        $spWeb.Dispose();
    }
}


#############################################################################################
#Function Name: CreateContentType               											#
#Parameters: $siteURL - URL where subsite is to be placed 									#
#            $xmlNewCT         - xml Object containing all the data for new CT              #
#                                                                                           #
#Return Value: N/A																			#
#Purpose: Creates a new content type in a site               								#
#############################################################################################

Function CreateContentType($webURL, $xmlNewCT)
{
    try
    {
        $spWeb = Get-SPWeb $webURL;

        if(!$spWeb.ContentTypes[$xmlNewCT.Name])
        { 
            $parentCT = $spWeb.AvailableContentTypes[$xmlNewCT.ParentCT];
            $newCT = New-Object Microsoft.SharePoint.SPContentType($parentCT,$spWeb.ContentTypes,$xmlNewCT.Name);
            $newCT.Group = $xmlNewCT.CTGroup;            
            $spWeb.ContentTypes.Add($newCT) | Out-Null;            
        }
        else
        {
            Write-Host ("{0}: Content type: {1} already exists." -f `
                        (get-date).ToString("yyyy-MM-dd HH:mm:ss"), $xmlNewCT.Name) -ForegroundColor Yellow
        }
        
    }
    catch
    {
        $ErrorMessage = $_.Exception.Message;
        Write-Host ("{0}: An error occurred creating the content type: {1}.  Error received: {2}" -f `
                        (get-date).ToString("yyyy-MM-dd HH:mm:ss"), $xmlNewCT.Name, $ErrorMessage) -ForegroundColor Red
    }
    finally
    {
        $spWeb.Dispose();
    }
}


#############################################################################################
#Function Name: AddSiteColumnToCT               											#
#Parameters: $siteURL - URL where subsite is to be placed 									#
#            $ctName  - Name of CT to add the site column to                                #
#            $siteColumnName - Name of site column to add to CT                             #
#                                                                                           #
#Return Value: N/A																			#
#Purpose: Adds a site column to a content type in a site    								#
#############################################################################################

Function AddSiteColumnToCT($webURL, $ctName, $siteColumnName)
{
    try
    {
        $spWeb = Get-SPWeb $webURL;

        $stToAdd = $spWeb.AvailableFields[$siteColumnName];
        $stFieldLink = New-Object Microsoft.SharePoint.SPFieldLink($stToAdd);
        
        $ctToUpdate = $spWeb.ContentTypes[$ctName];
        $ctToUpdate.FieldLinks.Add($stFieldLink);
        $ctToUpdate.Update();
        
    }
    catch
    {
        $ErrorMessage = $_.Exception.Message;
        Write-Host ("{0}: An error adding Site Column {1} to the content type: {2}.  Error received: {3}" -f `
                        (get-date).ToString("yyyy-MM-dd HH:mm:ss"), $siteColumnName, $ctName, $ErrorMessage) -ForegroundColor Red
    }
    finally
    {
        $spWeb.Dispose();
    }
}


#############################################################################################
#Function Name: AddDocTemplateToCT               											#
#Parameters: $siteURL - URL where subsite is to be placed 									#
#            $ctName  - Name of CT to add the site column to                                #
#            $templatePath - Path to the stored template to upload and attach               #
#            $templateName - Name of the file to add to the content type                    #
#                                                                                           #
#Return Value: N/A																			#
#Purpose: Adds a document template to a content type            							#
#############################################################################################

Function AddDocTemplateToCT($webURL, $ctName, $templatePath, $templateName)
{
    try
    {
        $spWeb = Get-SPWeb $webURL;
        
        $fileMode = [System.IO.FileMode]::Open
        $templateFile = $templatePath + "/" + $templateName;

        $fileStream = New-Object "System.IO.FileStream" -ArgumentList $templateFile, $fileMode;

        #Add the template to the contenttype's resource folder (holds files and other resources that can be added to a content type)
        $contentType = $spWeb.ContentTypes[$ctName];
        $contentType.ResourceFolder.Files.Add($templateName, $fileStream, $true) | Out-Null

        $fileStream.Close();

        #update the CT
        $contentType.DocumentTemplate = $templateName;
        $contentType.Update();
        
    }
    catch
    {
        $ErrorMessage = $_.Exception.Message;
        Write-Host ("{0}: An error occurred creating the content type: {1}.  Error received: {2}" -f `
                        (get-date).ToString("yyyy-MM-dd HH:mm:ss"), $xmlNewCT.Name, $ErrorMessage) -ForegroundColor Red
    }
    finally
    {
        $spWeb.Dispose();
    }
}


#############################################################################################
#Function Name: CreateLookupList                											#
#Parameters: $spWeb - subsite object to be updated      									#
#            $listIntenalName - The internal list name (URL)                                #
#            $listDisplayName - The list's display name                                     #
#                                                                                           #
#Return Value: N/A																			#
#Purpose: Creates a new Lookup site column in a site                 						#
#############################################################################################

Function CreateLookupList($spWeb, $listIntenalName, $listDisplayName)
{
    try
    {
        $listPath = $spWeb.Url + "/Lists/" + $listIntenalName;

        #get the custom list template
        $spListTemplate = $spWeb.ListTemplates["Custom List"];

        #create a collection of lists (will use this to add the new one)
        $spListCollection = $spWeb.Lists;
        $spNewListID = $spListCollection.Add($listIntenalName, "", $spListTemplate);

        #update the list with the display title
        $spList = $spWeb.Lists[$spNewListID];
        $spList.Title = $listDisplayName;
        $spList.Update();
    }
    catch
    {
        $ErrorMessage = $_.Exception.Message;
        Write-Host ("{0}: An error occurred creating the lookup list: {1}.  Error received: {2}" -f `
                        (get-date).ToString("yyyy-MM-dd HH:mm:ss"), $listDisplayName, $ErrorMessage) -ForegroundColor Red
    }
}