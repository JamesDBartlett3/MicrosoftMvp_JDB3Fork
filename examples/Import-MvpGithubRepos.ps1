#requires -version 7 -Modules MicrosoftMvp
<#
	.SYNOPSIS
	Updates your MVP profile with your Github repositories. Supports -WhatIf
	.DESCRIPTION
	Connects to the Github API to retrieve your repositories, then creates or updates corresponding activities on your MVP profile with relevant details like description, stars, forks, and recent activity. You can filter which repos to include based on technology focus area, minimum stars, and recent activity date. The script uses parallel processing for efficiency and supports -WhatIf for safe testing.
	.PARAMETER User
	The Github username whose repositories you want to import.
	.PARAMETER TechnologyFocusArea
	The primary technology focus area(s) to associate with the imported repositories. You can specify one or two focus areas. If more than two are provided, an error is raised. This parameter also supports dynamic tab-completion of available technology areas from the MVP module.
	.PARAMETER Filter
	An optional array of regex patterns to filter repositories by name, description, or topics. If not provided, all repositories are included.
	.PARAMETER MinimumStars
	The minimum number of stars a repository must have to be included. Default is 5.
	.PARAMETER EarliestActivityDate
	An optional date in YYYY-MM-DD format to filter repositories based on recent activity. Only repositories with commits since this date will be included. If not provided, all repositories are included regardless of activity.
	.PARAMETER ThrottleLimit
	The maximum number of repositories to process in parallel. Default is 20. Adjust based on your system's capabilities and the number of repositories being imported.
	.EXAMPLE
	PS> .\Import-MvpGithubRepos.ps1 -User "octocat" -TechnologyFocusArea "Cloud" -MinimumStars 10 -EarliestActivityDate "2023-01-01"
	This command imports repositories from the user "octocat" that have at least 10 stars and have had commits since January 1, 2023. The imported activities will be tagged with the "Cloud" technology focus area.
	.NOTES
	Original author: Justin Grote (@JustinGrote)
	Enhancements by: James Bartlett (@JamesDBartlett3)
#>

[CmdletBinding(SupportsShouldProcess)]
param(
	[Parameter(Mandatory)]
	$User,

	[Parameter(Mandatory)]
	# ValidateCount limits the number of values to a maximum of 3: one primary (mandatory) and up to two additional (optional).
	# If more than 3 are provided, an error is raised.
	[ValidateCount(1, 3)]
	# One or more technology focus areas. The first area will be used as the primary focus area,
	# and any additional areas (up to 2 more) will be passed as secondary/additional focus areas to the MVP activity form.
	[string[]]$TechnologyFocusArea,

	[Parameter(Mandatory)]
	# ArgumentCompleter provides dynamic tab-completion for this parameter.
	# The script block executes when user presses Tab and queries the MVP module for available technology areas.
	# The & operator (call operator) executes a script block within the MicrosoftMvp module's scope to access its internal data.
	[ArgumentCompleter({
		param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
		$mvpModule = Get-Module MicrosoftMvp -ErrorAction SilentlyContinue
		if ($mvpModule) {
			# The & operator runs the script in the module's context to access its data
			$techNames = & $mvpModule { (Get-MvpActivityData -ErrorAction SilentlyContinue).technologyArea.technologyName }
			$techNames | Where-Object { $_ -like "*$wordToComplete*" } | ForEach-Object {
				# Escape single quotes by doubling them so the completion result is properly quoted
				$escapedQuote = $_ -replace "'", "''"
				# CompletionResult creates a formatted completion item for the shell
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
	# -notmatch operator performs regex comparison; \d matches digits, {4} means exactly 4 occurrences
	# This pattern validates YYYY-MM-DD date format in order to eliminate ambiguity across different locales (e.g. MM/DD/YYYY vs DD/MM/YYYY)
	if ($EarliestActivityDate -notmatch '^\d{4}-\d{2}-\d{2}$') {
		Write-Warning "'-EarliestActivityDate $EarliestActivityDate' is not a recognized date format. Please use YYYY-MM-DD."
		return
	}
	# ParseExact converts a string to a specific datetime format. The third parameter is for culture info (null = InvariantCulture)
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
# ArrayList is used here because it's more efficient for repeated additions than regular PowerShell arrays
$repos = [System.Collections.ArrayList]::new()
$repos.AddRange($response)

# This regex uses a named capture group (?<url>...) to extract the next page URL from the Link header
# The -match operator populates the $matches hashtable with captured groups
while ($rh.Link -match '<(?<url>[^>]+)>; rel="next"') {
    # Access the named capture group 'url' from the $matches hashtable
    $nextUrl = $matches['url']
    $response = Invoke-RestMethod -Uri $nextUrl -ResponseHeadersVariable rh -Headers $headers
    $repos.AddRange($response)
}

Connect-Mvp -ErrorAction Stop

$existingActivities = Search-MvpActivitySummary -First 10000
| Where-Object type -Like 'Open Source*'

# Build a regex alternation pattern from the filter array; if no filter provided, pattern is $null
# The -join operator with '|' creates a regex pattern like "pattern1|pattern2|pattern3"
$filterPattern = if ($Filter) { $Filter -join '|' } else { $null }

$targetRepos = $repos
| Where-Object stargazers_count -GT $MinimumStars
# the (-not $filterPattern) condition allows all repos if no filter is provided; otherwise it checks if the repo name, description, or topics match the regex pattern
| Where-Object { (-not $filterPattern) -or ($_.name -match $filterPattern) -or ($_.description -match $filterPattern) -or ($_.topics -match $filterPattern) }
# the (-not $earliestDate) condition allows all repos if no earliest date is provided; otherwise it checks if the pushed_at date is greater than or equal to the earliest date
| Where-Object { (-not $earliestDate) -or ([datetime]$_.pushed_at -ge $earliestDate) }


#TODO: This "upsert" logic should probably be available in the main module
$mvpModulePath = (Get-Module MicrosoftMvp).path
$InformationPreference = 'Continue'

Write-Verbose 'Starting Parallel Invocation'
# -Parallel parameter executes the script block for each item in separate runspaces (PowerShell 7.0+ feature)
# -ThrottleLimit controls how many items are processed simultaneously
$targetRepos | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
	# $USING: is required in parallel script blocks to reference variables from the parent scope
	# Without $USING:, the variable is treated as local to the parallel script block
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
	# This line uses the ternary operator (? :) which is shorthand for if-then-else (PowerShell 7.0+)
	# Syntax: condition ? trueValue : falseValue
	$sinceParam = $earliestDate ? "&since=$($earliestDate.ToString('yyyy-MM-ddTHH:mm:ssZ'))" : ''

	$branches = Invoke-RestMethod -Uri "https://api.github.com/repos/$($repo.full_name)/branches?per_page=100" -Headers $commitHeaders -ErrorAction SilentlyContinue
	# HashSet automatically deduplicates items - using Add() on duplicate SHAs returns false but doesn't raise an error
	# This is more efficient than checking array membership for each commit
	$allCommitShas = [System.Collections.Generic.HashSet[string]]::new()
	$activeBranchCount = 0
	$authorParam = "&author=$($USING:User)"
	foreach ($branch in $branches) {
		# EscapeDataString encodes special characters in the branch name so it's URL-safe
		# This prevents errors with branch names containing special characters
		$encodedBranch = [Uri]::EscapeDataString($branch.name)
		$branchCommitUrl = "https://api.github.com/repos/$($repo.full_name)/commits?per_page=100&sha=$encodedBranch$sinceParam$authorParam"
		$branchHadCommit = $false
		do {
			$branchCommits = Invoke-RestMethod -Uri $branchCommitUrl -ResponseHeadersVariable branchRh -Headers $commitHeaders -ErrorAction SilentlyContinue
			foreach ($c in $branchCommits) {
				$null = $allCommitShas.Add($c.sha)
				$branchHadCommit = $true
			}
			# Extract the next page URL from pagination header. $matches[1] gets the first captured group (the URL inside parentheses)
			$branchCommitUrl = if ([string]$branchRh.Link -match '<([^>]+)>; rel="next"') { $matches[1] } else { $null }
		} while ($branchCommitUrl)
		if ($branchHadCommit) { $activeBranchCount++ }
	}
	$commitCount = $allCommitShas.Count

	$readmeResponse = Invoke-RestMethod -Uri "https://api.github.com/repos/$($repo.full_name)/readme" -Headers $commitHeaders -ErrorAction SilentlyContinue

	$readmeImageUrl = ''
	$repoPageHtml = Invoke-RestMethod -Uri $repo.html_url -Headers $commitHeaders -ErrorAction SilentlyContinue
	# The regex looks for the Open Graph image meta tag which GitHub sets to the first image in the README (if present). This is a simple way to get a representative image for the repo without having to parse the README ourselves. The captured URL is stored in $matches[1].
	if ($repoPageHtml -match '<meta property="og:image" content="([^"]+)"') {
		$readmeImageUrl = $matches[1]
	}
	if ($readmeResponse.content) {
		$rawBase = "https://raw.githubusercontent.com/$($repo.full_name)/$($repo.default_branch)"
		# GitHub API returns README content as base64-encoded. First decode from base64, then convert bytes to UTF8 string
		$readmeTextRaw = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($readmeResponse.content)).Trim()
		# Search for markdown images: ![alt](url). (?s) enables single-line mode (. matches newlines)
		# Store the position (idx) and captured image URL (src) in a hashtable for later comparison
		$mdMatch   = if ($readmeTextRaw -match '(?s)!\[[^\]]*\]\(([^)\s]+)') { @{ idx = $readmeTextRaw.IndexOf($matches[0]); src = $matches[1] } } else { $null }
		# Search for HTML images: <img ... src="url" />
		$htmlMatch = if ($readmeTextRaw -match '(?s)<img[^>]+src="([^"]+)"') { @{ idx = $readmeTextRaw.IndexOf($matches[0]); src = $matches[1] } } else { $null }
		$firstMatch = @($mdMatch, $htmlMatch) | Where-Object { $_ } | Sort-Object { $_.idx } | Select-Object -First 1
		if ($firstMatch) {
			$imgSrc = $firstMatch.src
			$readmeImageUrl = if ($imgSrc -match '^https?://') { $imgSrc } else { "$rawBase/" + $imgSrc.TrimStart('/') }
		}
	}

	$baseDescription = if ($readmeResponse.content) {
		$readmeText = $readmeTextRaw
		# Remove images from README text using regex patterns; -replace operator removes matches
		$readmeText = $readmeText -replace '\[!\[[^\]]*\]\([^)]*\)\]\([^)]*\)', ''  # Linked Markdown images: [![alt](img)](link)
		$readmeText = $readmeText -replace '!\[[^\]]*\]\([^)]*\)', ''               # Standalone Markdown images: ![alt](url)
		$readmeText = $readmeText -replace '(?s)<a[^>]*>\s*<img[^>]+/?>\s*</a>', '' # Linked HTML images: <a><img/></a>
		$readmeText = $readmeText -replace '(?s)<img[^>]+/?>', ''                   # Standalone HTML img tags
		$readmeText = $readmeText.Trim()
		if ($readmeText.Length -gt 750) { $readmeText.Substring(0, 750) + ' [...]' } else { $readmeText }
	} else {
		# The ?? operator (null coalescing) returns the left value unless it's null, then returns the right value (PowerShell 7.0+)
		$repo.description ?? 'No description provided.'
	}

	$reach = $repo.stargazers_count + $repo.forks_count
	$statsSuffix = "`n`n---`nReach: $reach (Stars: $($repo.stargazers_count), Forks: $($repo.forks_count)) | Branches: $activeBranchCount | Commits: $commitCount"

	$existingActivity = $USING:existingActivities
	| Where-Object title -EQ $activityTitle
	| ForEach-Object { Get-MvpActivity -Id $_.id }
	| Where-Object url -EQ $repo.html_url

	if ($existingActivity.count -gt 1) {
		Write-Warning "Multiple activities found for '$activityTitle' with url $($repo.html_url). Remove one from your MVP profile to update this."
		continue
	}

	# Ternary operator: if $existingActivity exists, use it; otherwise create a new activity
	$activity = $existingActivity ? $existingActivity : $(
    $tfa = $USING:TechnologyFocusArea
    # This is a hashtable (key-value pairs) that will be passed to New-MvpActivity using splatting (@newMvpActivityParams)
    $newMvpActivityParams = @{
      Title                     = $activityTitle
      Type                      = 'Open Source/Project/Sample code/Tools'
      TechnologyFocusArea       = $tfa[0]
      # Another ternary operator: if we have multiple technology areas, pass all except the first as additional areas
      # Array slicing syntax: $array[1..($array.Count - 1)] returns elements from index 1 to the end
      AdditionalTechnologyAreas = $tfa.Count -gt 1 ? $tfa[1..($tfa.Count - 1)] : @()
      TargetAudience            = 'Developer', 'IT Pro'
      Description               = $baseDescription + $statsSuffix
      Date                      = $repo.created_at
      EndDate                   = $repo.pushed_at
      Quantity                  = 1
      Reach                     = $reach
    }
    # @ symbol (splatting) expands the hashtable so each key-value pair becomes a separate parameter
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