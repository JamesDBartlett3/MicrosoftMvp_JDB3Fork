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
	if ($EarliestActivityDate -notmatch '^\d{4}-\d{2}-\d{2}$') {
		Write-Warning "'-EarliestActivityDate $EarliestActivityDate' is not a recognized date format. Please use YYYY-MM-DD."
		return
	}
	$earliestDate = [datetime]::ParseExact($EarliestActivityDate, 'yyyy-MM-dd', $null)
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

	# Fetch commit count across all branches, deduplicated by SHA
	$commitHeaders = $USING:headers
	$earliestDate = $USING:earliestDate
	$sinceParam = $earliestDate ? "&since=$($earliestDate.ToString('yyyy-MM-ddTHH:mm:ssZ'))" : ''

	$branches = Invoke-RestMethod -Uri "https://api.github.com/repos/$($repo.full_name)/branches?per_page=100" -Headers $commitHeaders -ErrorAction SilentlyContinue
	$allCommitShas = [System.Collections.Generic.HashSet[string]]::new()
	$activeBranchCount = 0
	$authorParam = "&author=$($USING:User)"
	foreach ($branch in $branches) {
		$encodedBranch = [Uri]::EscapeDataString($branch.name)
		$branchCommitUrl = "https://api.github.com/repos/$($repo.full_name)/commits?per_page=100&sha=$encodedBranch$sinceParam$authorParam"
		$branchHadCommit = $false
		do {
			$branchCommits = Invoke-RestMethod -Uri $branchCommitUrl -ResponseHeadersVariable branchRh -Headers $commitHeaders -ErrorAction SilentlyContinue
			foreach ($c in $branchCommits) {
				$null = $allCommitShas.Add($c.sha)
				$branchHadCommit = $true
			}
			$branchCommitUrl = if ([string]$branchRh.Link -match '<([^>]+)>; rel="next"') { $matches[1] } else { $null }
		} while ($branchCommitUrl)
		if ($branchHadCommit) { $activeBranchCount++ }
	}
	$commitCount = $allCommitShas.Count

	$readmeResponse = Invoke-RestMethod -Uri "https://api.github.com/repos/$($repo.full_name)/readme" -Headers $commitHeaders -ErrorAction SilentlyContinue

	$readmeImageUrl = ''
	$repoPageHtml = Invoke-RestMethod -Uri $repo.html_url -Headers $commitHeaders -ErrorAction SilentlyContinue
	if ($repoPageHtml -match '<meta property="og:image" content="([^"]+)"') {
		$readmeImageUrl = $matches[1]
	}
	if ($readmeResponse.content) {
		$rawBase = "https://raw.githubusercontent.com/$($repo.full_name)/$($repo.default_branch)"
		$readmeTextRaw = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($readmeResponse.content)).Trim()
		$mdMatch   = if ($readmeTextRaw -match '(?s)!\[[^\]]*\]\(([^)\s]+)') { @{ idx = $readmeTextRaw.IndexOf($matches[0]); src = $matches[1] } } else { $null }
		$htmlMatch = if ($readmeTextRaw -match '(?s)<img[^>]+src="([^"]+)"') { @{ idx = $readmeTextRaw.IndexOf($matches[0]); src = $matches[1] } } else { $null }
		$firstMatch = @($mdMatch, $htmlMatch) | Where-Object { $_ } | Sort-Object { $_.idx } | Select-Object -First 1
		if ($firstMatch) {
			$imgSrc = $firstMatch.src
			$readmeImageUrl = if ($imgSrc -match '^https?://') { $imgSrc } else { "$rawBase/" + $imgSrc.TrimStart('/') }
		}
	}

	$baseDescription = if ($readmeResponse.content) {
		$readmeText = $readmeTextRaw
		$readmeText = $readmeText -replace '\[!\[[^\]]*\]\([^)]*\)\]\([^)]*\)', ''  # Linked Markdown images: [![alt](img)](link)
		$readmeText = $readmeText -replace '!\[[^\]]*\]\([^)]*\)', ''               # Standalone Markdown images: ![alt](url)
		$readmeText = $readmeText -replace '(?s)<a[^>]*>\s*<img[^>]+/?>\s*</a>', '' # Linked HTML images: <a><img/></a>
		$readmeText = $readmeText -replace '(?s)<img[^>]+/?>', ''                   # Standalone HTML img tags
		$readmeText = $readmeText.Trim()
		if ($readmeText.Length -gt 750) { $readmeText.Substring(0, 750) + ' [...]' } else { $readmeText }
	} else {
		$repo.description ?? 'No description provided.'
	}
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
      Description               = $baseDescription + $statsSuffix
      Date                      = $repo.created_at
      EndDate                   = $repo.pushed_at
      Quantity                  = 1
      Reach                     = $reach
    }
    $newActivity = New-MvpActivity @newMvpActivityParams
    $newActivity.url = $repo.html_url
    $newActivity.imageUrl = $readmeImageUrl
    $newActivity
  )

	if ($existingActivity) {
		#Update the activity with the latest data
		$activity.Date = $repo.created_at
		$activity.DateEnd = $repo.updated_at
		$activity.Reach = $reach
		$activity.Description = $baseDescription + $statsSuffix
		$activity.url = $repo.html_url
		$activity.imageUrl = $readmeImageUrl
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