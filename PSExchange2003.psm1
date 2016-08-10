<#	
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2015 v4.2.98
	 Created on:   	08.01.2016 10:24
	 Created by:   	Chadima
	 Organization: 	
	 Filename:     	PSExchange2003.psm1
	-------------------------------------------------------------------------
	 Module Name: PSExchange2003
	===========================================================================

TODO: Try to use datatables insead of CustomObjects

#>

function Search-AdUser {
    [CmdletBinding()]
    Param (
        [Parameter(ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True)]
        [Object[]]$Identity
    )
    foreach ($user in $Identity){
        $strFilter = "(&(objectCategory=User)(|(sAMAccountName=" + $user + ")(displayName=" + $user + ")(cn=" + $user + ")))"   
        $objSearcher = [ADSISEARCHER]""
	    $objSearcher.SearchRoot = $objDomain
        $objSearcher.PageSize = 1000
        $objSearcher.SizeLimit = 1000
	    $objSearcher.Filter = $strFilter
	    $objSearcher.SearchScope = "Subtree"
	    Try {
		    $colResults = $objSearcher.FindOne().GetDirectoryEntry()
	    } Catch {
		    #$PSCmdlet.WriteError($_.Exception.Message) 
		    #return
	    }
    }
    return $colResults
}
function Get-2003UserMT {
	[CmdletBinding()]
	param (
        [Parameter(ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True)]
		[Object[]]$Identity,
        [Int]$ResultSize = 1000,
        [String]$Root
		)
    Begin {
        $WarningPreference = 'Continue'
        $ErrorActionPreference = 'Continue'
        $objDomain = [ADSI]''
        $maxThreads = 20
        $elapsed = [System.Diagnostics.Stopwatch]::StartNew()
    }
    Process {
        foreach ($user in $Identity){
            if ($user) {
                $strFilter = "(&(objectCategory=User)(|(sAMAccountName=" + $user + ")(displayName=" + $user + ")(cn=" + $user + ")))"   
            } else {
                $strFilter = "(&(objectCategory=person)(objectClass=organizationalPerson))"
                Write-Warning "Result is limited to $ResultSize. Use -ResultSize option to increase search result size."
                $split = $true
            }
            $objSearcher = [ADSISEARCHER]""
	        $objSearcher.SearchRoot = $objDomain
            $objSearcher.PageSize = 1000
            $objSearcher.SizeLimit = $ResultSize
	        $objSearcher.Filter = $strFilter
	        $objSearcher.SearchScope = "Subtree"
	        Try {
		        $colResults = $objSearcher.FindAll()
	        } Catch {
		        #$PSCmdlet.WriteError($_.Exception.Message) 
		        #return
	        }
            Write-Verbose "colResult Count: $($colResults.count)"
            Write-Verbose 'Foreach Loop Starting'
            $count = 0
            $splitSize = $colResults.Count / 10
            $usersArray = Resize-Array -InputObject $colResults -SplitSize $splitSize
	        $scriptBlock = {
                Param (
                    [Array]$array
                )
                $returnCollection = @()
                foreach ($objResult in $array) {
		            $objItem = $objResult.Properties
                    Write-Progress -Activity "Generating Object No. $count for User: $($objItem.userprincipalname)" -PercentComplete (($count / $colResults.Count) * 100) -Status "Working.."
                    if ($objItem.samaccountname) {
			            $sAMAccountName = $objItem.samaccountname[0]
		            } else {
			            $sAMAccountName = ''
		            }
                    if ($objItem.objectsid) {
                        $stringSID = (New-Object System.Security.Principal.SecurityIdentifier($objItem.objectsid[0],0))
			            $Sid =  $stringSID
		            } else {
			            $Sid = ''
		            }
                    if ($objItem.sidhistory) {
			            $SidHistory = $objItem.sidhistory
		            } else {
			            $SidHistory = @{}
		            }		
                    if ($objItem.userprincipalname) {
			            $UserPrincipalName = $objItem.userprincipalname[0]
		            } else {
			            $UserPrincipalName = ''
		            }
                    if ($objItem.givenname) {
			            $FirstName = $objItem.givenname[0]
		            } else {
			            $FirstName = ''
		            }
                    if ($objItem.sn) {
			            $lastName = $objItem.sn[0]
		            } else {
			            $lastName = ''
		            }
                    if ($objItem.name) {
			            $Name = $objItem.name[0]
		            } else {
			            $Name = ''
		            }
		            if ($objItem.displayname) {
			            $DisplayName = $objItem.displayname[0]
		            } else {
			            $DisplayName = ''
		            }
                    if ($objItem.displaynameprintable) {
			            $SimpleDisplayName = $objItem.displaynameprintable[0]
		            } else {
			            $SimpleDisplayName = ''
		            }
		            if ($objItem.distinguishedname) {
			            $DN = $objItem.distinguishedname[0]
		            } else {
			            $DN = ''
		            }
                    if ($objItem.employeenumber) {
                        $employeeNumber = $objItem.employeenumber[0]
                    } else {
                        $employeeNumber = ''
                    }
                    if ($objItem.useraccountcontrol) {
                        switch ($objItem.useraccountcontrol){
                            '512' {
	                            $userAccountControl = "NormalAccount"                            }
                            '514' {
	                            $userAccountControl = "AccountDisabled"                            }
                            '544' {
	                            $userAccountControl = "NormalAccount, PasswordNotRequired"                            }
                            '546' {
	                            $userAccountControl = "AccountDisabled, PasswordNotRequired"                            }
                            '66048' {
	                            $userAccountControl = "NormalAccount, DoNotExpirePassword"                            }
                            '66050' {
	                            $userAccountControl = "AccountDisabled, DoNotExpirePassword"                            }
                            '66080' {
	                            $userAccountControl = "NormalAccount, DoNotExpirePassword, PasswordNotRequired"                            }
                            '66082' {
	                            $userAccountControl = "AccountDisabled, DoNotExpirePassword, PasswordNotRequired"                            }
                            '262656' {
	                            $userAccountControl = "NormalAccount, SmartCardRequired"                            }
                            '262658' {
	                            $userAccountControl = "AccountDisabled, SmartCardRequired"                            }
                            '262688' {
	                            $userAccountControl = "NormalAccount, SmartCardRequired, PasswordNotRequired"                            }
                            '262690' {
	                            $userAccountControl = "AccountDisabled, SmartCardRequired, PasswordNotRequired"                            }
                            '328192' {
	                            $userAccountControl = "NormalAccount, SmartCardRequired, DoNotExpirePassword"                            }
                            '328194' {
	                            $userAccountControl = "AccountDisabled, SmartCardRequired, DoNotExpirePassword"                            }
                            '328224' {
	                            $userAccountControl = "NormalAccount, SmartCardRequired, DoNotExpirePassword, PasswordNotRequired"                            }
                            '328226' {
	                            $userAccountControl = "AccountDisabled, SmartCardRequired, DoNotExpirePassword, PasswordNotRequired"
                            }
                            Default {
                                $userAccountControl = $objItem.useraccountcontrol
                            }
                        }
			        
		            } else {
			            $userAccountControl = ''
		            }
		            if ($objItem.mail) {
			            $Email = $objItem.mail[0]
		            } else {
			            $Email = ''
		            }
		            if ($objItem.title) {
			            $Title = $objItem.title[0]
		            } else {
			            $Title = ''
		            }
		            if ($objItem.department) {
			            $Department = $objItem.department[0]
		            } else {
			            $Department = ''
		            }
		            if ($objItem.company) {
			            $Company = $objItem.company[0]
		            } else {
			            $Company = ''
		            }
                    if ($objItem.streetaddress) {
			            $StreetAddress = $objItem.streetaddress[0]
		            } else {
			            $StreetAddress = ''
		            }
                    if ($objItem.l) {
			            $City = $objItem.l[0]
		            } else {
			            $City = ''
		            }
                    if ($objItem.postalcode) {
			            $PostalCode = $objItem.postalcode[0]
		            } else {
			            $PostalCode = ''
		            }
                    if ($objItem.co) {
			            $Country = $objItem.co[0]
		            } else {
			            $Country = ''
		            }
		            if ($objItem.physicaldeliveryofficename) {
			            $Office = $objItem.physicaldeliveryofficename[0]
		            } else {
			            $Office = ''
		            }
		            if ($objItem.description) {
			            $Description = $objItem.description[0]
		            } else {
			            $Description = ''
		            }
		            if ($objItem.homemdb) {
			            $homeMDB = $objItem.homemdb[0]
		            } else {
			            $homeMDB = ''
		            }
                    if ($objItem.proxyaddresses) {
			            $proxyAddresses = $objItem.proxyaddresses
		            }
		            if ($objItem.telephonenumber) {
			            $Phone = $objItem.telephonenumber[0]
		            } else {
			            $Phone = ''
		            }	
		            if ($objItem.othertelephone) {
			            $Phone2 = $objItem.othertelephone
		            } else {
			            $Phone2 = @()
		            }
		            if ($objItem.mobile) {
			            $MobilePhone = $objItem.mobile[0]
		            } else {
			            $MobilePhone = ''
		            }
                    if ($objItem.facsimiletelephonenumber) {
			            $fax = $objItem.facsimiletelephonenumber[0]
		            } else {
			            $fax = ''
		            }
                    if ($objItem.homephone) {
			            $homePhone = $objItem.homephone[0]
		            } else {
			            $homePhone = ''
		            }
		            if ($objItem.msexchassistantname) {
			            $msExchAssistantName = $objItem.msexchassistantname[0]
		            } else {
			            $msExchAssistantName = ''
		            }
		            if ($objItem.telephoneassistant) {
			            $telephoneAssistant = $objItem.telephoneassistant[0]
		            } else {
			            $telephoneAssistant = ''
		            }
                    if ($objItem.objectcategory) {
			            $objectCategory = $objItem.objectcategory[0]
		            } else {
			            $objectCategory = ''
		            }
                    if ($objItem.objectclass) {
			            $objectClass = $objItem.objectclass
		            } else {
			            $objectClass = ''
		            }
                    if ($objItem.pwdlastset) {
                        $PasswordLastChanged = [DateTime]::FromFileTime($objItem.pwdlastset[0])
                    } else {
                        $PasswordLastChanged = ''
                    }
                    if ($objItem.whenchanged) {
			            $WhenChanged = Get-LocalTime $objItem.whenchanged[0]
		            } else {
			            $WhenChanged = ''
		            }
                    if ($objItem.whencreated) {
			            $WhenCreated = $objItem.whencreated[0]
		            } else {
			            $WhenCreated = ''
		            }
                
                    foreach ( $item in ($DN.replace('\,','~').split(","))) {
                        switch -regex ($item.TrimStart().Substring(0,3)) {
                            "CN=" {
                                $CN = '/' + $item.replace("CN=","")
                                continue
                            }
                            "OU=" {
                                $ou += ,$item.replace("OU=","");$ou += '/'
                                continue
                            }
                            "DC=" {
                                $DC += $item.replace("DC=","");$DC += '.'
                                continue
                            }
                        }
                    }
                    $CN = $CN.Replace('~', ',')
                    $canoincalOu = $dc.Substring(0,$dc.length - 1)
                    for ($i = $ou.count;$i -ge 0;$i -- ){
                        $canoincalOu += $ou[$i]
                    }
                    $identityCn = $canoincalOu + $CN
		            $obj = New-Object PSCustomObject -Property @{
                        'Name' = $Name
                        'SamAccountName' = $sAMAccountName
                        'Sid' = $Sid
                        'SidHistory' = $SidHistory
                        'UserPrincipalName' = $UserPrincipalName
                        'OrganizationalUnit' = $canoincalOu
                        'Identity' = $identityCn
                        'UserAccountControl' = $userAccountControl
                        'FirstName' = $FirstName
                        'Lastname' = $lastName
                        'DisplayName' = $DisplayName
                        'SimpleDisplayName' = $SimpleDisplayName
                        'DistinguishedName' = $DN
                        'EmployeeNumber' = $employeeNumber
                        'WindowsEmailAddress' = $Email
                        'Title' = $Title
                        'Department' = $Department
                        'Company' = $Company
                        'StreetAddress' = $StreetAddress
                        'City' = $City
                        'PostalCode' = $PostalCode
                        'CountryOrRegion' = $Country
                        'Office' = $Office
                        'Description' = $Description
                        'HomeMDB' = $homeMDB
                        'ProxyAddresses' = $proxyAddresses
                        'Phone' = $Phone
                        'OtherTelephone' = $Phone2
                        'MobilePhone' = $MobilePhone
                        'Fax' = $fax
                        'HomePhone' = $homePhone
                        'AssistantName' = $msExchAssistantName
                        'TelephoneAssistant' = $telephoneAssistant
                        'ObjectCategory' = $objectCategory
                        'ObjectClass' = $objectClass
                        'PasswordLastChanged' = $PasswordLastChanged
                        'WhenChanged' = $WhenChanged
                        'WhenCreated' = $WhenCreated
                    }
                    $obj.psobject.TypeNames.Insert(0,'PSExchange2003.GetUser.TypeName')
                    $returnCollection += $obj
                    $count++
	            }
                return $returnCollection
            }
            $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $maxThreads)
            $RunspacePool.Open()
            $Jobs = @()
            foreach ($array in $usersArray){
                $job = [System.Management.Automation.PowerShell]::Create().AddScript($scriptBlock).AddArgument($array)
                $job.RunspacePool = $RunspacePool
                $Jobs += New-Object psobject -Property @{
                    Pipe = $job
                    Result = $job.BeginInvoke()
                }
            }
            Do{
                $Total = ($jobs).Count
                $Completed = ($jobs | select -ExpandProperty Result | ?{$_.Iscompleted -eq $true}).Count
                Write-Progress `
	            -Activity "Waiting for Jobs - $($MaxThreads - $($RunspacePool.GetAvailableRunspaces())) of $MaxThreads threads running - Time elapsed: $($elapsed.Elapsed.ToString())" `
	            -PercentComplete (($($Total - $Jobs.Count) / $Total) * 100) `
	            -Status "$($Jobs.count) remaining" -Id 1
                Start-Sleep -Seconds 5
            } While (($jobs | select -ExpandProperty Result | ?{$_.IsCompleted -eq $false}).Count -gt 0)
            $results = @()
            Foreach($job in $jobs) {
                $pipe = $Job |Select-Object -ExpandProperty Pipe
                $result = $job | Select-Object -ExpandProperty Result
                $results += $pipe.EndInvoke($result)
            }
            $RunspacePool.Dispose()
            $RunspacePool.Close()
            #Write-Host "Total Elapsed Time: $($elapsed.Elapsed.ToString())"
        }
    }
    End {
	    return $results
    }
}
function Get-2003User {
	[CmdletBinding()]
	param (
        [Parameter(ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True)]
		[Object[]]$Identity,
        [Int]$ResultSize = 1000,
        [String]$Root,
        [String]$Filter,
        [Switch]$Progress
		)
    Begin {
        $WarningPreference = 'Continue'
        $ErrorActionPreference = 'Continue'
        $returnCollection = New-Object System.Collections.ArrayList
        $objDomain = [ADSI]''
        $objSearcher = [ADSISEARCHER]""
        if ($Root){
	        $objSearcher.SearchRoot = [ADSI]("LDAP://$Root")
        } else {
            $objSearcher.SearchRoot = $objDomain
        }
        #$objSearcher.SizeLimit = $ResultSize
        $objSearcher.PageSize = $ResultSize
        $objSearcher.SearchScope = "Subtree"
        #$objSearcher.CacheResults = $false
        $objSearcher.PropertiesToLoad.AddRange(@('samaccountname','sidhistory','userprincipalname','givenname','sn','name','displayname','displaynameprintable','distinguishedname','employeenumber','useraccountcontrol','mail','title','department','company','streetaddress','l','postalcode','co','physicaldeliveryofficename','description','homemdb','proxyaddresses','telephonenumber','othertelephone','mobile','fascimiletelephonenumber','homephone','msexchassistantname','telephoneassistant','objectcategory','objectclass','pwdlastset','whenchanged','whencreated'))
    }
    Process {
        foreach ($user in $Identity){
            if ($user) {
                $strFilter = "(&(objectCategory=User)(|(sAMAccountName=$user)(displayName=$user)(cn=$user)))"   
            } elseif ($Filter) {
                $strFilter = "(&(objectCategory=User)($Filter))"
                Write-Warning "Using filter: $strFilter to narrow results."
            } else {
                $strFilter = '(&(objectCategory=person)(objectClass=organizationalPerson))'
                Write-Warning "Result is limited to $ResultSize. Use -ResultSize option to increase search result size."
            }
	        $objSearcher.Filter = $strFilter
	        Try {
		        $colResults = $objSearcher.FindAll()
	        } Catch {
		        #$PSCmdlet.WriteError($_.Exception.Message) 
		        #return
	        }
            Write-Verbose "colResult Count: $($colResults.count)"
            Write-Verbose 'Foreach Loop Starting'
            $count = 0
	        foreach ($objResult in $colResults) {
		        $objItem = $objResult.Properties
                if ($objItem.objectsid) {
                    $stringSID = (New-Object System.Security.Principal.SecurityIdentifier($objItem.objectsid[0],0))
			        $Sid =  $stringSID
		        } else {
			        $Sid = ''
		        }
                if ($objItem.useraccountcontrol) {
                    switch ($objItem.useraccountcontrol){
                        '512' {
	                        $userAccountControl = "NormalAccount"                        }
                        '514' {
	                        $userAccountControl = "AccountDisabled"                        }
                        '544' {
	                        $userAccountControl = "NormalAccount, PasswordNotRequired"                        }
                        '546' {
	                        $userAccountControl = "AccountDisabled, PasswordNotRequired"                        }
                        '66048' {
	                        $userAccountControl = "NormalAccount, DoNotExpirePassword"                        }
                        '66050' {
	                        $userAccountControl = "AccountDisabled, DoNotExpirePassword"                        }
                        '66080' {
	                        $userAccountControl = "NormalAccount, DoNotExpirePassword, PasswordNotRequired"                        }
                        '66082' {
	                        $userAccountControl = "AccountDisabled, DoNotExpirePassword, PasswordNotRequired"                        }
                        '262656' {
	                        $userAccountControl = "NormalAccount, SmartCardRequired"                        }
                        '262658' {
	                        $userAccountControl = "AccountDisabled, SmartCardRequired"                        }
                        '262688' {
	                        $userAccountControl = "NormalAccount, SmartCardRequired, PasswordNotRequired"                        }
                        '262690' {
	                        $userAccountControl = "AccountDisabled, SmartCardRequired, PasswordNotRequired"                        }
                        '328192' {
	                        $userAccountControl = "NormalAccount, SmartCardRequired, DoNotExpirePassword"                        }
                        '328194' {
	                        $userAccountControl = "AccountDisabled, SmartCardRequired, DoNotExpirePassword"                        }
                        '328224' {
	                        $userAccountControl = "NormalAccount, SmartCardRequired, DoNotExpirePassword, PasswordNotRequired"                        }
                        '328226' {
	                        $userAccountControl = "AccountDisabled, SmartCardRequired, DoNotExpirePassword, PasswordNotRequired"
                        }
                        Default {
                            $userAccountControl = $objItem.useraccountcontrol
                        }
                    }
			        
		        } else {
			        $userAccountControl = ''
		        }
                if ($objItem.objectclass) {
			        $objectClass = $objItem.objectclass
		        } else {
			        $objectClass = ''
		        }
                if ($objItem.pwdlastset) {
                    $PasswordLastChanged = [DateTime]::FromFileTime($objItem.pwdlastset[0])
                } else {
                    $PasswordLastChanged = ''
                }
                
                foreach ( $item in (($objItem.distinguishedname[0]).replace('\,','~').split(","))) {
                    switch -regex ($item.TrimStart().Substring(0,3)) {
                        "CN=" {
                            $CN = '/' + $item.replace("CN=","")
                            continue
                        }
                        "OU=" {
                            $ou += ,$item.replace("OU=","");$ou += '/'
                            continue
                        }
                        "DC=" {
                            $DC += $item.replace("DC=","");$DC += '.'
                            continue
                        }
                    }
                }
                $CN = $CN.Replace('~', ',')
                $canoincalOu = $dc.Substring(0,$dc.length - 1)
                for ($i = $ou.count;$i -ge 0;$i -- ){
                    $canoincalOu += $ou[$i]
                }
                $identityCn = $canoincalOu + $CN
		        $obj = New-Object PSCustomObject -Property @{
                    'Name' = [String]$objItem.name
                    'SamAccountName' = [String]$objItem.samaccountname
                    'Sid' = $Sid
                    'SidHistory' = [String]$objItem.sidhistory
                    'UserPrincipalName' = [String]$objItem.userprincipalname
                    'OrganizationalUnit' = $canoincalOu
                    'Identity' = $identityCn
                    'UserAccountControl' = $userAccountControl
                    'FirstName' = [String]$objItem.givenname
                    'Lastname' = [String]$objItem.sn
                    'DisplayName' = [String]$objItem.displayname
                    'SimpleDisplayName' = [String]$objItem.displaynameprintable
                    'DistinguishedName' = [String]$objItem.distinguishedname
                    'EmployeeNumber' = [String]$objItem.employeenumber
                    'WindowsEmailAddress' = [String]$objItem.mail
                    'Title' = [String]$objItem.title
                    'Department' = [String]$objItem.department
                    'Company' = [String]$objItem.company
                    'StreetAddress' = [String]$objItem.streetaddress
                    'City' = [String]$objItem.l
                    'PostalCode' = [String]$objItem.postalcode
                    'CountryOrRegion' = [String]$objItem.co
                    'Office' = [String]$objDomain.physicaldeliveryofficename
                    'Description' = [String]$objItem.description
                    'HomeMDB' = [String]$objItem.homemdb
                    'ProxyAddresses' = [Array]$objItem.proxyaddresses
                    'Phone' = [String]$objItem.telephonenumber
                    'OtherTelephone' = [String]$objItem.othertelephone
                    'MobilePhone' = [String]$objItem.mobilenumber
                    'Fax' = [String]$objItem.facsimiletelephonenumber
                    'HomePhone' = [String]$objItem.homephone
                    'AssistantName' = [String]$objItem.msexchassistantname
                    'TelephoneAssistant' = [String]$objItem.telephoneassistant
                    'ObjectCategory' = [String]$objItem.objectcategory
                    'ObjectClass' = [String]$objItem.objectclass
                    'PasswordLastChanged' = $PasswordLastChanged
                    'WhenChanged' = [String]$objItem.whenchanged
                    'WhenCreated' = [String]$objItem.whencreated
                }
                $obj.psobject.TypeNames.Insert(0,'PSExchange2003.GetUser.TypeName')
                $returnCollection.Add($obj) > $null
                $count++
                <#
                if (($count % $ResultSize) -eq 0) {
                    [System.GC]::Collect()
                    [System.GC]::WaitForPendingFinalizers()
                }
                #>
                if ($Progress) {
                    Write-Progress -Activity "Generating Object #. $count from $($colResults.Count) for User: $($objItem.userprincipalname)" -PercentComplete (($count / $colResults.Count) * 100) -Status "Working.."
                }
	        }
        }
    }
    End {
        #$colResults.Dispose()
	    return $returnCollection
    }
}
function Set-2003User {
    [CmdletBinding()]
	param (
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True)]
		[Object]$Identity,
		[String]$Office,
		[String]$Department,
		[String]$Company,
		[String]$Street,
		[String]$City,
		[String]$PostalCode,
		[String]$Country,
		[String]$Phone,
		[String[]]$OtherTelephone,
		[String]$MobilePhone,
		[String]$Note,
        [string]$EmployeeNumber,
		[String]$AssistantName,
		[String]$TelephoneAssistant,
        [String]$Password,
        [Bool]$HiddenFromAddressListsEnabled
	)
    Begin { 
        Write-Verbose 'TEST'
    } Process {
        if($Identity.DistinguishedName){
            $dn = $Identity.DistinguishedName
        } else { 
            if ($Identity -eq $null) {
                Write-Verbose 'Searching for all users.. with * wildcard.'
                $strFilter = "(&(objectCategory=User)(sAMAccountName=*))"
            } else {
                $strFilter = "(&(objectCategory=User)(|(sAMAccountName=" + $Identity + ")(displayName=" + $Identity + ")(cn=" + $Identity + ")))"
            }
            $objSearcher = [ADSISEARCHER]""
	        $objSearcher.SearchRoot = $objDomain
	        $objSearcher.PageSize = $ResultSize
	        $objSearcher.Filter = $strFilter
	        $objSearcher.SearchScope = "Subtree"
	        Try {
		        $colResults = $objSearcher.FindOne().GetDirectoryEntry()
	        } Catch {
		        #$PSCmdlet.WriteError($_.Exception.Message) 
		        #return
	        }
            $dn = $colResults.properties.distinguishedname
            if( -not $ResultSize ){
                Write-Warning "WARNING: Result is limited to $ResultSize. Use -ResultSize option to increase search result size."
            }
        }
        Write-Verbose "DN: $dn"
	    $domainPerson = [ADSI]("LDAP://$($dn)")
	    if ($Office) {
            Try {
	            $domainPerson.Put('physicalDeliveryOfficeName', $Office)
            } Catch {
                Write-Error "ERROR: $($_.Error.ErrorMessage)"
            } Finally {
                Write-Verbose "INFO: Set Office Completed."
            }
	    }
	    if ($Title) {
            Try {
	            $domainPerson.Put('title', $Title)
            } Catch {
                Write-Error "ERROR: $($_.Error.ErrorMessage)"
            } Finally {
                Write-Verbose "INFO: Set Title Completed."
            }
	    }
	    if ($Department) {
            Try {
		        $domainPerson.Put('department', $Department)
            } Catch {
                Write-Error "ERROR: $($_.Error.ErrorMessage)"
            } Finally {
                Write-Verbose "INFO: Set Department Completed."
            }
	    }
	    if ($Company) {
		    Try {
                $domainPerson.Put('company', $Company)
            } Catch {
                Write-Error "ERROR: $($_.Error.ErrorMessage)"
            } Finally {
                Write-Verbose "INFO: Set Company Completed."
            }
	    }
	    if ($Street) {
            Try {
		        $domainPerson.Put('streetAddress', $Street)
            } Catch {
                Write-Error "ERROR: $($_.Error.ErrorMessage)"
            } Finally {
                Write-Verbose "INFO: Set StreetAddress Completed."
            }
	    }
	    if ($City) {
            Try {
		        $domainPerson.Put('l', $City)
            } Catch {
                Write-Error "ERROR: $($_.Error.ErrorMessage)"
            } Finally {
                Write-Verbose "INFO: Set City Completed."
            }
	    }
	    if ($PostalCode) {
            Try {
		        $domainPerson.Put('postalCode', $PostalCode)
            } Catch {
                Write-Error "ERROR: $($_.Error.ErrorMessage)"
            } Finally {
                Write-Verbose "INFO: Set PostalCode Completed."
            }
	    }
	    if ($Country) {
            Try {
		        $domainPerson.Put('c', $Country)
            } Catch {
                Write-Error "ERROR: $($_.Error.ErrorMessage)"
            } Finally {
                Write-Verbose "INFO: Set Country Completed."
            }
	    }
	    if ($Phone) {
            Try {
		        $domainPerson.Put('telephoneNumber', $Phone)
            } Catch {
                Write-Error "ERROR: $($_.Error.ErrorMessage)"
            } Finally {
                Write-Verbose "INFO: Set Phone Completed."
            }
	    }
	    if ($OtherTelephone) {
            Try {
		        $domainPerson.PutEx('3','otherTelephone', $OtherTelephone)
            } Catch {
                Write-Error "ERROR: $($_.Error.ErrorMessage)"
            } Finally {
                Write-Verbose "INFO: Set OtherTelephone Completed."
            }
	    }
	    if ($MobilePhone) {
            Try {
		        $domainPerson.Put('mobile', $MobilePhone)
            } Catch {
                Write-Error "ERROR: $($_.Error.ErrorMessage)"
            } Finally {
                Write-Verbose "INFO: Set MobilePhone Completed."
            }
	    }
	    if ($Note) {
            Try {
		        $domainPerson.Put('description', $Note)
            } Catch {
                Write-Error "ERROR: $($_.Error.ErrorMessage)"
            } Finally {
                Write-Verbose "INFO: Set Note Completed."
            }
	    }
        if ($EmployeeNumber) {
            Try {
		        $domainPerson.Put('employeeNumber', $EmployeeNumber)
            } Catch {
                Write-Error "ERROR: $($_.Error.ErrorMessage)"
            } Finally {
                Write-Verbose "INFO: Set EmployeeNumber Completed."
            }
	    }
	    if ($AssistantName) {
            Try {
		        $domainPerson.Put('msExchAssistantName', $AssistantName)
            } Catch {
                Write-Error "ERROR: $($_.Error.ErrorMessage)"
            } Finally {
                Write-Verbose "INFO: Set AssistantName Completed."
            }
	    }
        if ($TelephoneAssistant) {
            Try {
		        $domainPerson.Put('telephoneAssistant', $TelephoneAssistant)
            } Catch {
                Write-Error "ERROR: $($_.Error.ErrorMessage)"
            } Finally {
                Write-Verbose "INFO: Set TelephoneAssistant Completed."
            }
        }
        if ($Password){
            Try {
                $domainPerson.psbase.Invoke('SetPassword', $Password)
            } Catch {
                Write-Error "ERROR: $($_.Error.ErrorMessage)"
            } Finally {
                Write-Verbose "INFO: Set Password Completed."
            }
        }
        if (Get-Variable HiddenFromAddressListsEnabled -ErrorAction SilentlyContinue){
            Try {
                $domainPerson.msExchHideFromAddressLists = $HiddenFromAddressListsEnabled
            } Catch {
                Write-Error "ERROR: $($_.Error.ErrorMessage)"
            } Finally {
                Write-Verbose "INFO: Set Hidden From Address Lists Completed."
            }
        }
        
        Try {
	        $domainPerson.SetInfo()
        } Catch {
            Write-Error "ERROR: $($_.Error.ErrorMessage)"
        } Finally {
            Write-Verbose "INFO: SetInfo Command Completed."
        }
    } End {
        $domainPerson.Dispose()
    }
}
function Get-2003Mailboxes(){
    $mailboxes = @()
	foreach ($mbx in $mbxs) {
        $obj = New-Object System.Management.Automation.PSObject -Property @{
            Server = $mbx.ServerName
            StorageGroupName = $mbx.StorageGroupName
            StoreName = $mbx.StoreName
            DisplayName = $mbx.MailboxDisplayName
            Size = [Math]::Round($mbx.Size / 1KB)
            TotalItems =$mbx.TotalItems
            LastLoggedOnUser = $mbx.LastLoggedOnUserAccount
            LastLogonTime = if($mbx.LastLogonTime -ne $null){ [System.Management.ManagementDateTimeConverter]::ToDateTime($mbx.LastLogonTime) }else{ $null }
            LastLogoffTime = if($mbx.LastLogoffTime -ne $null){ [System.Management.ManagementDateTimeConverter]::ToDateTime($mbx.LastLogoffTime) }else{ $null }
        } 
        $mailboxes += $obj
    }
    return $mailboxes
}
function Get-2003MailboxStatistics {
	param (
		[String]$Identity = "",
		[String]$Server = $env:ComputerName,
		[String]$Database = ""
	)
	
	$Filter = "ServerName='$Server'"
	if ($Database -ne "") {
		$Filter = "$Filter AND StoreName='$Database'"
	}
	if ($Identity -ne "") {
		$Filter = "$Filter AND (MailboxGuid='{$Identity}' OR LegacyDN='$Identity'"
		$Filter = "$Filter OR MailboxDisplayName='$Identity'"
	}
	
	Get-WMIObject Exchange_Mailbox -Namespace "root/MicrosoftExchangeV2" -ComputerName $Server -Filter $Filter |
	Select-Object `
	AssocContentCount,
	@{
		n = 'DateDiscoveredAbsentInDs'; e = {
			if ($_.DateDiscoveredAbsentInDs -ne $null) {
				[Management.ManagementDateTimeConverter]::ToDateTime($_.DateDiscoveredAbsentInDs)
			}
		}
	},
	MailboxDisplayName, TotalItems, LastLoggedOnUserAccount,
	@{
		n = 'LastLogonTime'; e = {
			if ($_.LastLogonTime -ne $null) {
				[Management.ManagementDateTimeConverter]::ToDateTime($_.LastLogonTime)
			}
		}
	},
	@{
		n = 'LastLogoffTime'; e = {
			if ($_.LastLogoffTime -ne $null) {
				[Management.ManagementDateTimeConverter]::ToDateTime($_.LastLogoffTime)
			}
		}
	},
	LegacyDN,
	@{ n = 'MailboxGuid'; e = { ([String]$_.MailboxGuid).ToLower() -replace "{|}" } },
	@{ n = 'ObjectClass'; e = { "Mailbox" } },
	@{
		n = 'StorageLimitStatus'; e = {
			switch ($_.StorageLimitInfo) {
				1  { "BelowLimit" }
				2  { "IssueWarning" }
				4  { "ProhibitSend" }
				8  { "NoChecking" }
				16 { "MailboxDisabled" }
			}
		}
	},
	DeletedMessageSize, Size,
	@{
		n = 'Database'; e = {
			"$($_.ServerName)\$($_.StorageGroupName)\$($_.StoreName)"
		}
	},
	ServerName, StorageGroupName, StoreName,
	@{ n = 'Identity'; e = { ([String]$_.MailboxGuid).ToLower() -replace "{|}" } }
}
function Get-2003MailStorePath {
	param (
		$serverName,
		$storageGroup,
		$mailboxStore
	)
	$wmiServer = Get-WmiObject -Class Exchange_Server -Namespace root\MicrosoftExchangeV2 | Where-Object{ $_.Name -like "*$serverName*" }
	#prepare CDOEXM COM Objects
	$cdoexmIExchangeServer = New-Object -com CDOEXM.ExchangeServer
	$cdoexmIStorageGroup = New-Object -com CDOEXM.StorageGroup
	$cdoexmIMailboxStoreDB = New-Object -com CDOEXM.MailboxStoreDB
	$cdoexmIPublicStoreDB = New-Object -com CDOEXM.PublicStoreDB
	#Initialize ExchangeServer Datasource
	$cdoexmIExchangeServer.DataSource.Open("LDAP://" + $wmiServer.DN)
	foreach ($sg in $cdoexmIExchangeServer.StorageGroups) {
		$cdoexmIStorageGroup.DataSource.Open($sg)
		if ($storageGroup -eq $cdoexmIStorageGroup.Name) {
			foreach ($mbxDB in $cdoexmIStorageGroup.MailboxStoreDBs) {
				$cdoexmIMailboxStoreDB.DataSource.Open($mbxDB)
				if ($mailboxStore -eq $cdoexmIMailboxStoreDB.Name) {
					$found = $true
					$mailStorePath = $mbxDB
				}
			}
		}
	}
	return $mailStorePath
}
function Get-2003Mailbox {
    param (
        [String]$userName, 
        [String]$mailbox, 
        [String]$storageGroup, 
        [String]$exchangeServer = $env:COMPUTERNAME,
        [Int]$ResultSize = 1000,
        [String]$domain = 'mail.cd.cz'
    )
    
    $objDomain = [ADSI]''

	if ($Identity) {
        $strFilter = "(&(objectCategory=User)(|(sAMAccountName=" + $Identity + ")(displayName=" + $Identity + ")(cn=" + $Identity + ")))"   
    } else {
        $strFilter = "(&(objectCategory=User)(sAMAccountName=*))"
        Write-Warning "Result is limited to $ResultSize. Use -ResultSize option to increase search result size."
    }
    $objSearcher = [ADSISEARCHER]""
	$objSearcher.SearchRoot = $objDomain
	$objSearcher.SizeLimit = $ResultSize
	$objSearcher.Filter = $strFilter
	$objSearcher.SearchScope = "Subtree"
	Try {
		$colResults = $objSearcher.FindAll()
	} Catch {
		#$PSCmdlet.WriteError($_.Exception.Message) 
		#return
	}
	$objcollection = @()
    
    $namespace = 'ROOT\MicrosoftExchangeV2'
    if ($mailbox){
        $wmiQuery = "Select * From Exchange_Mailbox Where StoreName='$mailbox'"
    } else {
        $wmiQuery = "Select * From Exchange_Mailbox"
    }
    $mbxs =  Get-WmiObject -Query $wmiQuery -Namespace $namespace -ComputerName $exchangeServer
	
    foreach ($objResult in $colResults) {
        foreach ($mbx in $mbxs){
            if ($mbx.MailboxDisplayName -eq $objResult.displayname) {
		        $objItem = $objResult.Properties
		        $obj = New-Object System.Management.Automation.PSObject
		        if ($objItem.name) {
			        $obj | Add-Member -MemberType NoteProperty -Name Name -Value $objItem.name[0]
		        }
                if ($objItem.displayname) {
			        $obj | Add-Member -MemberType NoteProperty -Name DisplayName -Value $objItem.displayname[0]
		        }
		        if ($objItem.samaccountname) {
			        $obj | Add-Member -MemberType NoteProperty -Name sAMAccountName -Value $objItem.samaccountname[0]
		        }
		        if ($objItem.distinguishedname) {
			        $obj | Add-Member -MemberType NoteProperty -Name distinguishedName -Value $objItem.distinguishedname[0]
		        }
		        if ($objItem.mail) {
			        $obj | Add-Member -MemberType NoteProperty -Name Email -Value $objItem.mail[0]
		        }
		        if ($objItem.title) {
			        $obj | Add-Member -MemberType NoteProperty -Name Title -Value $objItem.title[0]
		        }
		        if ($objItem.department) {
			        $obj | Add-Member -MemberType NoteProperty -Name Department -Value $objItem.department[0]
		        }
		        if ($objItem.company) {
			        $obj | Add-Member -MemberType NoteProperty -Name Company -Value $objItem.company[0]
		        }
		        if ($objItem.office) {
			        $obj | Add-Member -MemberType NoteProperty -Name Office -Value $objItem.office[0]
		        }
		        if ($obgetjItem.description) {
			        $obj | Add-Member -MemberType NoteProperty -Name Description -Value $objItem.description[0]
		        }
		        if ($objItem.homemdb) {
			        $obj | Add-Member -MemberType NoteProperty -Name homeMDB -Value $objItem.homemdb[0]
		        }
                if ($objItem.proxyaddresses) {
			        $obj | Add-Member -MemberType NoteProperty -Name homeMDB -Value $objItem.proxyaddresses
		        }
		        if ($objItem.phone) {
			        $obj | Add-Member -MemberType NoteProperty -Name Phone -Value $objItem.phone[0]
		        } else {
			        $obj | Add-Member -MemberType NoteProperty -Name Phone -Value ''
		        }	
		        if ($objItem.phone2) {
			        $obj | Add-Member -MemberType NoteProperty -Name Phone2 -Value $objItem.phone2[0]
		        }
		        if ($objItem.mobilephone) {
			        $obj | Add-Member -MemberType NoteProperty -Name MobilePhone -Value $objItem.mobilephone[0]
		        }
		        if ($objItem.msexchassistantname) {
			        $obj | Add-Member -MemberType NoteProperty -Name msExchAssistantName -Value $objItem.msexchassistantname[0]
		        }
		        if ($objItem.telephoneassistant) {
			        $obj | Add-Member -MemberType NoteProperty -Name telephoneAssistant -Value $objItem.telephoneassistant[0]
		        }
		        $objCollection += $obj
            }
        }
	}
	return $objCollection
}
function Set-EmployeeNumber {
    [CmdletBinding()]
	param (
        [Parameter(ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True)]
		[Object[]]$Identity,
        [String]$employeeNumber
		)
    Begin {
        $logfile = $env:TEMP + '\SetEmployeeNumber.log'
    }
    Process{
        if (!$Identity){
            $Identity = Read-Host -Prompt 'Please Enter Username or Displayname'
            Write-Output "[$(Get-Date -f s)][INFO] Entered UserName or DisplayName: $($Identity)" | Out-File -FilePath $logfile -Encoding utf8 -Append
        }
	    $strObjResult = Search-ADUser -Identity $Identity
        $user = $strObjResult.Properties	    
        $domainPerson = [ADSI]("LDAP://$($user.distinguishedname)")
        if (!$employeeNumber) {
            $employeeNumber = Read-Host "Enter EmployeeNumber for $($user.displayname) (Current: $($user.employeenumber))"
            Write-Output "[$(Get-Date -f s)][INFO] Entered EmployeeNumber: $($employeeNumber)" | Out-File -FilePath $logfile -Encoding utf8 -Append
        }
        $domainPerson.Put('employeenumber', $employeeNumber)
        Try {
	        $domainPerson.SetInfo()
        } Catch {
            $ErrorMessage = $_.Exception.Message
			Write-Host "[$(Get-Date -f s)][ERROR] $ErrorMessage" -ForegroundColor Red
            Write-Output "[$(Get-Date -f s)][ERROR] $ErrorMessage" | Out-File -FilePath $logfile -Encoding utf8 -Append
            Break
        }
        $ldapString = "LDAP://" + $user.distinguishedname
        $userADSI = [ADSI]$ldapString
    }
    End {
        Write-Host "[INFO] EmployeeNumber Successfully set and is now $($userADSI.employeenumber)." -ForegroundColor Green
        Write-Output "[$(Get-Date -f s)][INFO] EmployeeNumber Successfully set and is now $($userADSI.employeenumber)." | Out-File -FilePath $logfile -Encoding utf8 -Append
    }
}
function Resize-Array ([object[]]$InputObject, [int]$SplitSize = 100) {
	$length = $InputObject.Length
	for ($Index = 0; $Index -lt $length; $Index += $SplitSize){
		, ($InputObject[$index .. ($index + $splitSize - 1)])
	}
}
function Get-LocalTime($UTCTime){
    $strCurrentTimeZone = (Get-WmiObject win32_timezone).StandardName
    $TZ = [System.TimeZoneInfo]::FindSystemTimeZoneById($strCurrentTimeZone)
    $LocalTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($UTCTime, $TZ)
    Return $LocalTime
}
Export-ModuleMember Get-2003UserMT,
                    Get-2003User,
					Set-2003User,
					Get-2003Mailboxes,
					Get-2003MailboxStatistics,
					Get-2003MailStorePath,
                    Get-2003Mailbox,
                    Set-EmployeeNumber