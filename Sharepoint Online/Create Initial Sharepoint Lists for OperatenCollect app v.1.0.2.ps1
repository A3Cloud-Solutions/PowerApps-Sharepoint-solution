$Error.Clear() # clear error stack
Try{
    #Installing and importing requird modules
    if (Get-Module -ListAvailable -Name MSOnline) {
        Write-Host -f Yellow "MSOnline is alredy exist"
    } 
    else {
        Install-Module -Name "MSOnline" -Force -AllowClobber # install MSOnline module 
    }

    if (Get-Module -ListAvailable -Name PnP.PowerShell) {
        Write-Host -f Yellow "PnP.PowerShell is alredy exist"
    } 
    else {
        Install-Module -Name "PnP.PowerShell" -Force -AllowClobber # install PnP module
    }
    #Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
    #Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
    #Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine
    Import-Module -Name "MSOnline" -Force # import MSOnline module
    Import-Module -Name "PnP.PowerShell" -Force # import PnP module
}
Catch{
    Write-Host -f Red "Module error: " $_.Exception.Message
}

If(!$Error){
    #define required variables
    $Error.Clear() # clear error stack

    $MSOLConnection = Connect-MsolService # connect to MSOnline by user credentials which we get earlier
    $Domain = Get-MsolDomain | Where-Object {$_.IsInitial -eq $true} # get initial domain from tenant
    if($MSOLConnection) { Disconnect-MsolService -Connection $MSOLConnection } # disconnect from MSOnline

    $AdminSiteURL = "https://{0}.sharepoint.com/" -f $Domain.Name.Split('.')[0] # url for connecting to sharepoint service
    $SPSiteName = "Operate and Collect Demo102" # Site name
    $AliasName = "OnCd102" # Site alias name - SharepointOnline team site, to store and operate the data
    $ListName1 = "Objects" # The first List name - General information and Object settings Table
    $ListName2 = "Common Data Base" # The second List name - The main operational data store list 
    $ListName3 = "CompanyInfo" # The third List name - Common information about Company incl. company logo 
	$ListName4 = "KPI Set" # The forth List name - Set of KPI to operation level
    $ErrorCount = 0 # Error counting for display success messages

    Connect-PnPOnline -Url $AdminSiteURL -Interactive # Connect to sherepoint service via PnP module to the main page

    # Creating initial Site
    Try{
        $Site = Get-PnPTenantSite | Where {$_.Title -eq $SPSiteName} # Checking on the avaliable site name
        if($Site -eq $null){
            $AliasCheck = Get-PnPIsSiteAliasAvailable -Identity $AliasName | select IsAvailable
            if($AliasCheck.IsAvailable){ #Checking on the avaliable alias name

            Write-Host -f Cyan ("Creating initial Sharepoint site '$($SPSiteName)' ... ")

                #Create initial SharePoint Site 'Operate and Collect'
                $teamSiteUrl = New-PnPSite -Type TeamSite `
                                -Title $SPSiteName `
                                -Lcid 1033 `
                                -Alias $AliasName
                <#
                    Type - site type is Team
                    Title - site name
                    Lcid - language number(1033 - english)
                    Alias - the name by which you can access the site
                #>
            }
            Else{
                Write-Host -f Red "Alias '$($AliasName)' unavaliable"
                $ErrorCount += 1 # if we has an error increment the count variable
            }
        }
        Else{
            Write-Host -f Red "The site '$($SPSiteName)' is already exist"
            $ErrorCount += 1 # if we has an error increment the count variable
        }
    }
    Catch{
        Write-Host -f Red "Site creating error: " $_.Exception.Message
        $ErrorCount += 1 # if we has an error increment the count variable
        $Error.Clear() # clear error stack
    }

    #Creating site Success message
    If($ErrorCount -eq 0){ # if we have no errors display the message
        Write-Host -f Green "Site '$($SPSiteName)' with alias '$($AliasName)' successfully created!"
    }


    If($ErrorCount -eq 0){
        Connect-PnPOnline -Url $teamSiteUrl -Interactive # Connect to created 'Operate and Collect' site page

        #Objects list
        #Creating new inital sharepoint List in 'Operate and collect' site 
        Try{
            
            Write-Host -f Cyan ("Creating Sharepoint Online list '$($ListName1)' ... ")
            
            New-PnPList -Title $ListName1 -Template GenericList | Out-Null # Create new sharepoint list
            <#
                Title - list name which you want to create
                Template - the type which you want in this list
                 Out-Null - don't display any messages after create 
            #>
            Add-PnPField -List $ListName1 -DisplayName "Brand" -InternalName "Brand" -Type Text | Out-Null # Create new field to store Brand in Sharepoint list - 
            <#
                List - list name which you want to update
                DisplayName - The column name which you can see in list 
                InternalName - The column name which you can use to contact from external service
                Type - the type which you want in this field (single-line text, hyperlink or Pictures, Number and etc.)
                Out-Null - don't display any messages after adding field
            #>
            $LogoXML= "<Field Type='URL' Name='LogoURL' ID='$([GUID]::NewGuid())' DisplayName='LogoURL' StaticName='LogoURL' Format='Hyperlink' ></Field>"
            Add-PnPFieldFromXml -List $ListName1 -FieldXml $LogoXML| Out-Null


            $fieldXml = '<Field Type="UserMulti" DisplayName="SharedwithUser" List="UserInfo" Required="FALSE" ID="{bd66d0d0-c441-4ce3-909a-204ff01cdcb7}" ShowField="EMail" UserSelectionMode="PeopleAndGroups" StaticName="SharedwithUser" Name="SharedwithUser" Mult="TRUE" />'
            Add-PnPFieldFromXml -List $ListName1 -FieldXml $fieldXml| Out-Null # Create new field in sharepoint list by xml
            <#
                List - list name which you want to update
                FieldXml - xml properties for the field
                Out-Null - don't display any messages after adding field

                Xml properties:
                    Field Type - the type which you want in this field
                    DisplayName - The column name which you can see in list
                    Required - enable or disable required property
                    ID - Guid id
                    ShowField - what you want to see in this field
                    UserSelectionMode - what you want to choose in this field
                    StaticName - The column name which you can use to contact from external service
                    Name - The column name which you can see in list
                    Mult - Enable or disable multichoice
            #>
            $ReportXML= "<Field Type='URL' Name='ReportLink' ID='$([GUID]::NewGuid())' DisplayName='ReportLink' StaticName='ReportLink' Format='Hyperlink' ></Field>"
            Add-PnPFieldFromXml -List $ListName1 -FieldXml $ReportXML| Out-Null
            Add-PnPField -List $ListName1 -DisplayName "LogoImage1" -InternalName "LogoImage1" -Type Thumbnail | Out-Null
            Add-PnPField -List $ListName1 -DisplayName "VisibleInputFields" -InternalName "VisibleInputFields" -Type MultiChoice -Choices "InputField1", "InputField2", "InputField3", "InputField4", "InputField5", "InputField6", "InputField7", "InputField8", "InputField9" | Out-Null
            Set-PnPList -Identity $ListName1 -EnableAttachments $True | Out-Null # allow to use attachments field
            <#
                Identity - list name which you want to update
                EnableAttachments - enable or disable attachments
                Out-Null - don't display any messages
            #>
            Set-PnPView -List $ListName1 -Identity "All Items" -Fields "Brand","LogoURL","SharedwithUser","ReportLink","LogoImage1","VisibleInputFields" | Out-Null # add fields to default view
        }
        Catch{
            Write-Host -f Red "List '$($ListName1)' creating error: " $_.Exception.Message
            $ErrorCount += 1 # if we has an error increment the count variable
            $Error.Clear() # clear error stack
        }

        #Creating list and Fields Success message
        If($ErrorCount -eq 0){ # if we have no errors display the message
            Write-Host -f Green "Created the list '$($ListName1)' with the following columns: 'Brand', 'LogoURL', 'SharedwithUser', 'ReportLink', 'LogoImage1', 'VisibleInputFields' and enabled attachments"
        }


        #Common Data Base list
        #Creating new inital sharepoint List in 'Operate and collect' site 
        Try{
        
            Write-Host -f Cyan ("Creating Sharepoint Online list '$($ListName2)' ... ")

            New-PnPList -Title $ListName2 -Template GenericList | Out-Null
            <#
                Title - list name which you want to crate
                Template - the type which you want in this list
            #>
            $dateXml = "<Field Type='DateTime' Name='Date' ID='$([GUID]::NewGuid())' DisplayName='Date' Required ='FALSE' Format='DateOnly' ShowField='Date' FriendlyDisplayFormat='Disabled'></Field>"
            Add-PnPFieldFromXml -FieldXml $dateXml -List $ListName2 | Out-Null
            <#
                List - list name which you want to update
                FieldXml - xml properties for the field
                Out-Null - don't display any messages after adding field

                Xml properties:
                    Field Type - the type which you want in this field
                    DisplayName - The column name which you can see in list
                    Required - enable or disable required property
                    ID - Guid id
                    ShowField - what you want to see in this field
                    Name - The column name which you can use to contact from external service
                    Format - display format
                    FriendlyDisplayFormat - enable or disable frendly display format
            #>
            Add-PnPField -List $ListName2 -DisplayName "Brand" -InternalName "Brand" -Type Text | Out-Null
            <#
                List - list name which you want to update
                DisplayName - The column name which you can see in list 
                InternalName - The column name which you can use to contact from external service
                Type - the type which you want in this field (single-line text, hyperlink or Pictures, Number and etc.)
                AddToDefaultView - show this column to the user(in default is hide)
                Out-Null - don't display any messages after adding field
            #>
            Add-PnPField -List $ListName2 -DisplayName "Revenue" -InternalName "InputField1" -Type Number | Out-Null
            Add-PnPField -List $ListName2 -DisplayName "Checks" -InternalName "InputField2" -Type Number | Out-Null
            Add-PnPField -List $ListName2 -DisplayName "Guests" -InternalName "InputField3" -Type Number | Out-Null
            Add-PnPField -List $ListName2 -DisplayName "Delivery Revenue" -InternalName "InputField4" -Type Number | Out-Null
            Add-PnPField -List $ListName2 -DisplayName "Delivery Checks" -InternalName "InputField5" -Type Number | Out-Null
            Add-PnPField -List $ListName2 -DisplayName "Expenses" -InternalName "InputField6" -Type Number | Out-Null
            Add-PnPField -List $ListName2 -DisplayName "InputField7" -InternalName "InputField7" -Type Number | Out-Null
            Add-PnPField -List $ListName2 -DisplayName "InputField8" -InternalName "InputField8" -Type Number | Out-Null
            Add-PnPField -List $ListName2 -DisplayName "InputField9" -InternalName "InputField9" -Type Number | Out-Null
            Add-PnPField -List $ListName2 -DisplayName "Average Check" -InternalName "CalculatedParametr1" -Type Number | Out-Null
            Add-PnPField -List $ListName2 -DisplayName "Average Delivery Checks" -InternalName "CalculatedParametr2" -Type Number | Out-Null
            Add-PnPField -List $ListName2 -DisplayName "Gross profit margin" -InternalName "CalculatedParametr3" -Type Number | Out-Null
            Add-PnPField -List $ListName2 -DisplayName "CalculatedParametr4" -InternalName "CalculatedParametr4" -Type Number | Out-Null
            Set-PnPView -List $ListName2 -Identity "All Items" -Fields "Date","Brand","InputField1","InputField2","InputField3","InputField4","InputField5","InputField6","CalculatedParametr1","CalculatedParametr2","CalculatedParametr3" | Out-Null
        }
        Catch{
            Write-Host -f Red "List '$($ListName2)' creating error: " $_.Exception.Message
            $ErrorCount += 1 # if we has an error increment the count variable
            $Error.Clear() # clear error stack
        }

        #Creating list and Fields Success message
        If($ErrorCount -eq 0){ # if we have no errors display the message
            Write-Host -f Green "Created the list '$($ListName2)' with the following columns: 'Date', 'Brand', 'Revenue', 'Checks', 'Guests', 'Delivery Revenue', 'Delivery Checks', 'Expenses', 'Average Check', 'Average Delivery Checks', 'Gross profit margin'"
        }

        #CompanyInfo list
        #Creating new inital sharepoint List in 'Operate and collect' site 
        Try{
        
            Write-Host -f Cyan ("Creating Sharepoint Online list '$($ListName3)' ... ")


            New-PnPList -Title $ListName3 -Template GenericList | Out-Null
            <#
                Title - list name which you want to crate
                Template - the type which you want in this list
            #>
            $CompanyLogoXML= "<Field Type='URL' Name='CompanyLogoURL' ID='$([GUID]::NewGuid())' DisplayName='CompanyLogoURL' StaticName='CompanyLogoURL' Format='Hyperlink' ></Field>"
            Add-PnPFieldFromXml -List $ListName3 -FieldXml $CompanyLogoXML| Out-Null
            <#
                List - list name which you want to update
                DisplayName - The column name which you can see in list 
                InternalName - The column name which you can use to contact from external service
                Type - the type which you want in this field (single-line text, hyperlink or Pictures, Number and etc.)
                AddToDefaultView - show this column to the user(in default is hide)
                Out-Null - don't display any messages after adding field
            #>
            Add-PnPField -List $ListName3 -DisplayName "Brands" -InternalName "Brands" -Type Text | Out-Null
            Set-PnPView -List $ListName3 -Identity "All Items" -Fields "Title","Modified","Created","CompanyLogoURL","Brands","Created By","Modified By" | Out-Null
        }
        Catch{
            Write-Host -f Red "List '$($ListName3)' creating error: " $_.Exception.Message
            $ErrorCount += 1 # if we has an error increment the count variable
            $Error.Clear() # clear error stack
        }

        #Creating list and Fields Success message. Message with URL
        If($ErrorCount -eq 0){ # if we have no errors display the message
            Write-Host -f Green "Created the list '$($ListName3)' with the following columns: 'CompanyLogoURL', 'Brands'"
        }


        #KPI Set list
        #Creating new inital sharepoint List in 'Operate and collect' site 
        Try{
		    
			Write-Host -f Cyan ("Creating Sharepoint Online list '$($ListName4)'")

		
            New-PnPList -Title $ListName4 -Template GenericList | Out-Null
            <#
                Title - list name which you want to crate
                Template - the type which you want in this list
            #>
            $dateXml = "<Field Type='DateTime' Name='KPI Period' ID='$([GUID]::NewGuid())' DisplayName='Date' Required ='FALSE' Format='DateOnly' ShowField='Date' FriendlyDisplayFormat='Disabled'></Field>"
            Add-PnPFieldFromXml -FieldXml $dateXml -List $ListName4 | Out-Null
            <#
                List - list name which you want to update
                FieldXml - xml properties for the field
                Out-Null - don't display any messages after adding field

                Xml properties:
                    Field Type - the type which you want in this field
                    DisplayName - The column name which you can see in list
                    Required - enable or disable required property
                    ID - Guid id
                    ShowField - what you want to see in this field
                    Name - The column name which you can use to contact from external service
                    Format - display format
                    FriendlyDisplayFormat - enable or disable frendly display format
            #>
            Add-PnPField -List $ListName4 -DisplayName "Brand" -InternalName "Brand" -Type Text | Out-Null
            <#
                List - list name which you want to update
                DisplayName - The column name which you can see in list 
                InternalName - The column name which you can use to contact from external service
                Type - the type which you want in this field (single-line text, hyperlink or Pictures, Number and etc.)
                AddToDefaultView - show this column to the user(in default is hide)
                Out-Null - don't display any messages after adding field
            #>

            Add-PnPField -List $ListName4 -DisplayName "Revenue" -InternalName "InputField1" -Type Number | Out-Null
            Add-PnPField -List $ListName4 -DisplayName "Checks" -InternalName "InputField2" -Type Number | Out-Null
            Add-PnPField -List $ListName4 -DisplayName "Guests" -InternalName "InputField3" -Type Number | Out-Null
            Add-PnPField -List $ListName4 -DisplayName "Delivery Revenue" -InternalName "InputField4" -Type Number | Out-Null
            Add-PnPField -List $ListName4 -DisplayName "Delivery Checks" -InternalName "InputField5" -Type Number | Out-Null
            Add-PnPField -List $ListName4 -DisplayName "Expenses" -InternalName "InputField6" -Type Number | Out-Null
            Add-PnPField -List $ListName4 -DisplayName "InputField7" -InternalName "InputField7" -Type Number | Out-Null
            Add-PnPField -List $ListName4 -DisplayName "InputField8" -InternalName "InputField8" -Type Number | Out-Null
            Add-PnPField -List $ListName4 -DisplayName "InputField9" -InternalName "InputField9" -Type Number | Out-Null
            Add-PnPField -List $ListName4 -DisplayName "Average Check" -InternalName "CalculatedParametr1" -Type Number | Out-Null
            Add-PnPField -List $ListName4 -DisplayName "Average Delivery Checks" -InternalName "CalculatedParametr2" -Type Number | Out-Null
            Add-PnPField -List $ListName4 -DisplayName "Gross profit margin" -InternalName "CalculatedParametr3" -Type Number | Out-Null
            Add-PnPField -List $ListName4 -DisplayName "CalculatedParametr4" -InternalName "CalculatedParametr4" -Type Number | Out-Null
            Set-PnPView -List $ListName4 -Identity "All Items" -Fields "Date","Brand","InputField1","InputField2","InputField3","InputField4","InputField5","InputField6","CalculatedParametr1","CalculatedParametr2","CalculatedParametr3" | Out-Null
        }
        Catch{
            Write-Host -f Red "List '$($ListName4)' creating error: " $_.Exception.Message
            $ErrorCount += 1 # if we has an error increment the count variable
            $Error.Clear() # clear error stack
        }

        #Creating list and Fields Success message
        If($ErrorCount -eq 0){ # if we have no errors display the message
            Write-Host -f Green "Created the list '$($ListName2)' with the following columns: 'Date', 'Brand', 'Revenue', 'Checks', 'Guests', 'Delivery Revenue', 'Delivery Checks', 'Expenses', 'Average Check', 'Average Delivery Checks', 'Gross profit margin'"
            Write-Host -f Green ("You can access the site by following this link - {0}sites/{1}/" -f $SiteURL, $AliasName)
        }

        Disconnect-PnPOnline # disconnect from sharepoint site page
    }
}