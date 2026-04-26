#requires -version 7 -Modules MicrosoftMvp
<#
.SYNOPSIS
Updates your MVP profile with your Github repositories. Supports -WhatIf
#>

[CmdletBinding(SupportsShouldProcess)]
param(
	[Parameter(Mandatory)]
	$User,

	[Parameter(Mandatory)]
	[ArgumentCompleter({
		param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
		$mvpModule = Get-Module MicrosoftMvp -ErrorAction SilentlyContinue
		if ($mvpModule) {
			$techNames = & $mvpModule { (Get-MvpActivityData -ErrorAction SilentlyContinue).technologyArea.technologyName }
			$techNames | Where-Object { $_ -like "*$wordToComplete*" } | ForEach-Object {
				$escapedQuote = $_ -replace "'", "''"
				[System.Management.Automation.CompletionResult]::new("'$escapedQuote'", $_, 'ParameterValue', $_)
			}
		}
	})]
	[string[]]$TechnologyFocusArea,

	#A regex filter for repo titles, can be useful to select certain repos to update for different technology focus areas.
	[string[]]$Filter,

	[ValidateNotNullOrEmpty()]
	[int]$MinimumStars = 5,

	[string]$EarliestActivityDate,

	[int]$ThrottleLimit = 20
)

if ($EarliestActivityDate) {
	if ($EarliestActivityDate -match '^\d{1,2}/\d{1,2}/\d{4}$') {
		Write-Warning "'-EarliestActivityDate $EarliestActivityDate' is ambiguous (could be MM/DD/YYYY or DD/MM/YYYY). Please use YYYY-MM-DD format."
		return
	}
	if ($EarliestActivityDate -notmatch '^\d{4}\D\d{2}\D\d{2}$') {
		Write-Warning "'-EarliestActivityDate $EarliestActivityDate' is not a recognized date format. Please use YYYY-MM-DD."
		return
	}
	$earliestDate = [datetime]::Parse(($EarliestActivityDate -replace '\D', '-'))
} else {
	$earliestDate = $null
}
$apiUrl = "https://api.github.com/users/$User/repos?per_page=100&sort=pushed"

$headers = @{}
# If the caller has $env:GITHUB_TOKEN, use it to avoid rate limits
if ($env:GITHUB_TOKEN) {
    $headers['Authorization'] = "token $env:GITHUB_TOKEN"
}

# Invoke-RestMethod doesn't natively follow Link headers (RFC 5988), 
# so we need to loop until we have all pages
$response = Invoke-RestMethod -Uri $apiUrl -ResponseHeadersVariable rh -Headers $headers
$repos = [System.Collections.ArrayList]::new()
$repos.AddRange($response)

while ($rh.Link -match '<(?<url>[^>]+)>; rel="next"') {
    $nextUrl = $matches['url']
    $response = Invoke-RestMethod -Uri $nextUrl -ResponseHeadersVariable rh -Headers $headers
    $repos.AddRange($response)
}

Connect-Mvp -ErrorAction Stop

$existingActivities = Search-MvpActivitySummary -First 10000
| Where-Object type -Like 'Open Source*'

$filterPattern = if ($Filter) { $Filter -join '|' } else { $null }

$targetRepos = $repos
| Where-Object stargazers_count -GT $MinimumStars
| Where-Object { (-not $filterPattern) -or ($_.name -match $filterPattern) -or ($_.description -match $filterPattern) -or ($_.topics -match $filterPattern) }
| Where-Object { (-not $earliestDate) -or ([datetime]$_.pushed_at -ge $earliestDate) }


#TODO: This "upsert" logic should probably be available in the main module
$mvpModulePath = (Get-Module MicrosoftMvp).path
$InformationPreference = 'Continue'

Write-Verbose 'Starting Parallel Invocation'
$targetRepos | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
	Import-Module -Force $USING:mvpModulePath
	$VerbosePreference = $USING:VerbosePreference
	$WhatIfPreference = $USING:WhatIfPreference
	$DebugPreference = $USING:DebugPreference

	$repo = $PSItem
	$activityTitle = "GitHub: $($repo.name)"
	Write-Verbose "Processing $activityTitle"
	$existingActivity = $USING:existingActivities
	| Where-Object title -EQ $activityTitle
	| ForEach-Object { Get-MvpActivity -Id $_.id }
	| Where-Object url -EQ $repo.html_url

	if ($existingActivity.count -gt 1) {
		Write-Warning "Multiple activities found for '$activityTitle' with url $($repo.html_url). Remove one from your MVP profile to update this."
		continue
	}

	$activity = $existingActivity ? $existingActivity : $(
    $tfa = $USING:TechnologyFocusArea
    $newMvpActivityParams = @{
      Title                     = $activityTitle
      Type                      = 'Open Source/Project/Sample code/Tools'
      TechnologyFocusArea       = $tfa[0]
      AdditionalTechnologyAreas = $tfa.Count -gt 1 ? $tfa[1..($tfa.Count - 1)] : @()
      TargetAudience            = 'Developer', 'IT Pro'
      Description               = $repo.description ?? 'No description provided.'
      Date                      = $repo.created_at
      EndDate                   = $repo.pushed_at
      Quantity                  = 1
      Reach                     = $repo.stargazers_count
    }
    $newActivity = New-MvpActivity @newMvpActivityParams
    $newActivity.url = $repo.html_url
    $newActivity
  )

	if ($existingActivity) {
		#Update the activity with the latest data
		$activity.Date = $repo.created_at
		$activity.DateEnd = $repo.updated_at
		$activity.Reach = $repo.stargazers_count
		$activity.Description = $repo.description ?? $activity.Description
		$activity.url = $repo.html_url
		#Workaround for whatif not working as it is supposed to in parallel
		if ($WhatIfPreference) {
			Write-Host "WhatIf: Updating repository '$activityTitle':"
			$activity | Format-List * | Out-String | Write-Host
			return
		}
		Write-Verbose "Updating repository '$activityTitle' with url $($repo.html_url)"
		Set-MvpActivity $activity -WhatIf:$WhatIfPreference
	} else {
		if ($WhatIfPreference) {
			Write-Host "WhatIf: Adding repository '$activityTitle':"
			$activity | Format-List * | Out-String | Write-Host
			return
		}
		Write-Verbose "Adding new repository '$activityTitle' with url $($repo.html_url)"
		Add-MvpActivity $activity -WhatIf:$WhatIfPreference
	}
}