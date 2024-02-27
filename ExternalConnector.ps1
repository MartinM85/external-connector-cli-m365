# create an app

m365 login

$permissions = 'https://graph.microsoft.com/ExternalConnection.ReadWrite.OwnedBy,https://graph.microsoft.com/ExternalItem.ReadWrite.OwnedBy,https://graph.microsoft.com/User.Read.All,https://graph.microsoft.com/MailboxSettings.Read'

$appResponse = m365 entra app add `
--name 'UserDataExternalConnectorApp' `
--withSecret `
--apisApplication $permissions `
--grantAdminConsent | ConvertFrom-Json

m365 logout

$appResponse

# wait for a while until the Entra app is provisioned

Start-Sleep -Seconds 10

# login to a new app 

m365 login --tenant $appResponse.tenantId --appId $appResponse.appId --secret $appResponse.secrets[0].value --authType secret

# connector id
$externalConnectionId = 'UserDataConnector'

# create connector
m365 external connection add --id $externalConnectionId --name 'User data' --description 'User data connection'

# define and create schema

$params = @{
	baseType = "microsoft.graph.externalItem"
	properties = @(
		@{
			name = "id"
			type = "String"
			isRetrievable = "true"
            isQueryable = "false"
            isSearchable = "true"
			labels = @(
				"title"
			)
		}
		@{
			name = "userPrincipalName"
			type = "String"
			isRetrievable = "true"
            isQueryable = "false"
            isSearchable = "true"
            aliases = @(
                "upn"
            )
		}
		@{
			name = "userPurpose"
			type = "String"
			isRetrievable = "true"
            isQueryable = "true"
            isSearchable = "false"
            aliases = @(
                "userType"
            )
		}
        @{
			name = "locale"
			type = "String"
			isRetrievable = "true"
            isQueryable = "true"
            isSearchable = "false"
		}
        @{
			name = "timeZone"
			type = "String"
			isRetrievable = "true"
            isQueryable = "true"
            isSearchable = "false"
		}
        @{
			name = "employeeType"
			type = "String"
			isRetrievable = "true"
            isQueryable = "true"
            isSearchable = "false"
		}
        @{
			name = "mySite"
			type = "String"
			isRetrievable = "true"
            isQueryable = "false"
            isSearchable = "false"
            aliases = @(
                "site"
            )
		}
        @{
			name = "manager"
			type = "String"
			isRetrievable = "true"
            isQueryable = "true"
            isSearchable = "true"
		}
        @{
			name = "country"
			type = "String"
			isRetrievable = "true"
            isQueryable = "true"
            isSearchable = "false"
		}
        @{
			name = "city"
			type = "String"
			isRetrievable = "true"
            isQueryable = "true"
            isSearchable = "false"
		}
        @{
			name = "department"
			type = "String"
			isRetrievable = "true"
            isQueryable = "true"
            isSearchable = "false"
		}
	)
}

$schemaJson = $params | ConvertTo-Json -Compress -Depth 3

m365 external connection schema add -i $externalConnectionId --schema $schemaJson --wait

# read all users ids

$usersWithId = m365 entra user list --properties id | ConvertFrom-Json

Foreach ($userWithId in $usersWithId) {
  # read details of each user
  $user = m365 entra user get --id $userWithId.id --properties 'id,userPrincipalName,mailboxSettings,employeeType,mySite,country,city,department' --withManager | ConvertFrom-Json

  $locale = $user.mailboxSettings.language.locale -eq $null ? '' : $user.mailboxSettings.language.locale
  $timeZone = $user.mailboxSettings.timeZone -eq $null ? '' : $user.mailboxSettings.timeZone
  $employeeType = $user.employeeType -eq $null ? '' : $user.employeeType
  $mySite = $user.mySite -eq $null ? '' : $user.mySite
  $manager = ($user.manager -eq $null -or $user.manager.userPrincipalName -eq $null) ? '' : $user.manager.userPrincipalName
  $country = $user.country -eq $null ? '' : $user.country
  $city = $user.city -eq $null ? '' : $user.city
  $department = $user.department -eq $null ? '' : $user.department
  $content = $user.userPrincipalName -eq $null ? '' : $user.userPrincipalName
  #
  m365 external item add --externalConnectionId $externalConnectionId `
  --id $userWithId.id `
  --userPrincipalName $user.userPrincipalName `
  --userPurpose $user.mailboxSettings.userPurpose `
  --locale $locale `
  --timeZone $timeZone `
  --employeeType $employeeType `
  --mySite $mySite `
  --manager $manager `
  --country $country `
  --city $city `
  --department $department `
  --content $content `
  --acls 'grant,everyone,everyone'
}

m365 logout