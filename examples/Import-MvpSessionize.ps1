#requires -module MicrosoftMvp,PowerHTML

using namespace HtmlAgilityPack

<#
.SYNOPSIS
Sessionize doesn't provide a Speaker API, so we must scrape the result of the 'https://sessionize.com/app/speaker/events' page. THIS IS A SCREENSCRAPER AND MAY BREAK IF SESSIONIZE CHANGES THEIR PAGE LAYOUT.
.DESCRIPTION
Connects to Sessionize using cookie-based authentication headers and scrapes your speaker events page to find accepted sessions. Creates or updates corresponding activities on your MVP profile. You can filter which sessions to include by name, date, and technology focus area. The script uses parallel processing for efficiency and supports -WhatIf for safe testing.
.PARAMETER RequestHeaders
The raw HTTP request headers from an authenticated Sessionize session. Log into Sessionize, open DevTools (F12), navigate to your events page, right-click the 'events' request, and choose 'Copy -> Copy request headers'. Then pass the clipboard content: (Get-Clipboard -Raw)
.PARAMETER TechnologyFocusArea
The primary technology focus area(s) to associate with the imported sessions. You can specify one to three focus areas. The first area will be used as the primary focus area, and any additional areas (up to 2 more) will be passed as secondary/additional focus areas to the MVP activity form.
.PARAMETER Filter
An optional array of regex patterns to filter session names. If not provided, all accepted sessions are included.
.PARAMETER EarliestActivityDate
An optional date in YYYY-MM-DD format. Only sessions on or after this date will be included. If not provided, all sessions are included regardless of date.
.PARAMETER ThrottleLimit
The maximum number of sessions to process in parallel. Default is 30.
.EXAMPLE
.\Import-MvpSessionize.ps1 -RequestHeaders (Get-Clipboard -Raw) -TechnologyFocusArea 'Microsoft Fabric' -EarliestActivityDate '2024-01-01' -WhatIf
.NOTES
Original author: Justin Grote (@JustinGrote)
Enhancements by: James Bartlett (@JamesDBartlett3)
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(

	[Parameter(Mandatory)]$RequestHeaders,

	[Parameter(Mandatory)]
	# ValidateCount limits the number of values to a maximum of 3: one primary (mandatory) and up to two additional (optional).
	[ValidateCount(1, 3)]
	# One or more technology focus areas. The first area will be used as the primary focus area,
	# and any additional areas (up to 2 more) will be passed as secondary/additional focus areas to the MVP activity form.
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

	#An optional array of regex patterns to filter session names. If not provided, all accepted sessions are included.
	[string[]]$Filter,

	#An optional date in YYYY-MM-DD format. Only sessions on or after this date will be included.
	[string]$EarliestActivityDate,

	$ThrottleLimit = 30
)

if ($EarliestActivityDate) {
	# Validate YYYY-MM-DD format to avoid locale-dependent date parsing ambiguity
	if ($EarliestActivityDate -notmatch '^\d{4}-\d{2}-\d{2}$') {
		Write-Warning "'-EarliestActivityDate $EarliestActivityDate' is not a recognized date format. Please use YYYY-MM-DD."
		return
	}
	$earliestDate = [datetime]::ParseExact($EarliestActivityDate, 'yyyy-MM-dd', $null)
} else {
	$earliestDate = $null
}

# Build a regex alternation pattern from the filter array; if no filter provided, pattern is $null
$filterPattern = if ($Filter) { $Filter -join '|' } else { $null }

if ($RequestHeaders -notmatch '.AspNet.ApplicationCookie=([^;]*)') {
	throw 'Invalid Request Headers, it should have an aspnet.applicationcookie in the content.'
}

$session = [Microsoft.PowerShell.Commands.WebRequestSession]::new()
$authCookie = [Net.Cookie]::new('.AspNet.ApplicationCookie', $matches[1], '/', 'sessionize.com')
$session.Cookies.Add($authCookie)
$PSDefaultParameterValues['Invoke-RestMethod:WebSession'] = $session
$PSDefaultParameterValues['Invoke-RestMethod:Verbose'] = $false
$PSDefaultParameterValues['Invoke-RestMethod:Debug'] = $false

$ErrorActionPreference = 'stop'
Write-Warning 'This is a screenscraper and may break if Sessionize changes their events page layout.'

#region Utility
filter Assert-ScreenScrapeData {
	param(
		[string]$Target
	)

	if (-not $PSItem) { throw "Invalid Data Detected: $Target. this is either not a valid Sessionize events response, or Sessionize has changed their format and broken this script" }
	return $PSItem
}

filter ConvertFrom-HtmlEncoding {
	[System.Net.WebUtility]::HtmlDecode($PSItem)
}
#endregion

function ConvertFrom-SessionizeSessionTab {
	<#
	.SYNOPSIS
	Screenscrapes the data from a Sessionize event tab.
	#>
	param(
		[Parameter(Mandatory)][HtmlNode]$Tab
	)

	$rows = $Tab.SelectNodes('.//div[@class="table-responsive"]/table/tbody/tr')

	$currentEvent = $null
	$currentEventEndDate = $null
	$sessions = foreach ($row in $rows) {
		#Events are specified in the row with an event class, so determine if this row is an event or a session.
		try {
			if ($row.GetClasses() -contains 'event') {
				$newEvent = $row.SelectNodes('.//h4/a').InnerText.Trim()
				| Assert-ScreenScrapeData 'Table Event Name'
				$currentEvent = $newEvent

				# Extract the event end date from the date range in the <small> tag (e.g. "23 Feb – 1 Mar 2026")
				$currentEventEndDate = $null
				$eventDateText = $row.SelectSingleNode('.//small')?.InnerText.Trim() | ConvertFrom-HtmlEncoding
				if ($eventDateText) {
					# Take the part after the last dash/en-dash/em-dash as the end date
					$endDatePart = ($eventDateText -split '[\u2013\u2014-]')[-1].Trim()
					try { $currentEventEndDate = [datetime]::Parse($endDatePart) }
					catch { Write-Verbose "Could not parse event end date from '$eventDateText'" }
				}

				#The first row should be an event, so if it is not, something is wrong with the format.
				if (-not $currentEvent) { throw 'Event not found in first row of the table, this is either not a valid Sessionize events response, or Sessionize has changed their format and broken this script' }
				continue
			}

			#Otherwise it should be a session entry
			$session = [ordered]@{}

			$session.Event = $currentEvent

			$session.Title = $row.SelectSingleNode('.//td/a[starts-with(@href, "/app/speaker/session")]').InnerText.Trim()
			| ConvertFrom-HtmlEncoding
			| Assert-ScreenScrapeData 'Session Title'

			$session.Name = "$($session.Event): $($session.Title)"
			if ($filterPattern -and ($session.Name -notmatch $filterPattern)) {
				Write-Verbose "Skipping $($session.Name) due to not matching filter"
				continue
			}

			$session.Url = 'https://sessionize.com' + $row.SelectSingleNode('.//td/a[starts-with(@href, "/app/speaker/session")]').GetAttributeValue('href', '')

			# Find the status span - look for a span with a class containing "badge-primary"
			$session.Status = $row.SelectSingleNode('.//span[@class and contains(@class, "badge-primary")]')?.InnerText.Trim()
			
			if (-not $session.Status) {
				# Fallback to any span with a badge class
				$session.Status = $row.SelectSingleNode('.//span[@class and contains(@class, "badge")]')?.InnerText.Trim()
			}
			
			if (-not $session.Status) {
				# Final fallback to any span
				$session.Status = $row.SelectSingleNode('.//span')?.InnerText.Trim()
			}
			
			$session.Status = $session.Status
			| ConvertFrom-HtmlEncoding
			| Assert-ScreenScrapeData 'Event Acceptance Status'

			Write-Verbose "Session '$($session.Title)' - Status: '$($session.Status)'"

			# Match accepted status case-insensitively and handle potential whitespace issues
			if ($session.Status -inotmatch '^\s*accepted\s*$') {
				Write-Verbose "Skipping unaccepted session '$($session.Title)' for '$currentEvent' (Status: '$($session.Status)')"
				#Unaccepted sessions dont have a date which will crash the below line.
				continue
			}

			# Try to get the date, but make it optional (some sessions may not have a scheduled date yet)
			$dateText = $row.SelectSingleNode('.//td[3]').InnerText.Trim()
			if ($dateText) {
				$session.Date = [datetime]($dateText | ConvertFrom-HtmlEncoding | Assert-ScreenScrapeData 'Session Date')
			} elseif ($currentEventEndDate) {
				# Fall back to the last day of the event date range
				$session.Date = $currentEventEndDate
				Write-Verbose "No date for session '$($session.Title)', using event end date: $($session.Date)"
			} else {
				# No date available at all — prompt the user to enter one
				do {
					$userDateInput = Read-Host "No date found for session '$($session.Title)' in event '$currentEvent'. Enter the session date (e.g. 1 Mar 2026)"
					$parsedDate = $null
					$valid = [datetime]::TryParse($userDateInput, [ref]$parsedDate)
					if (-not $valid) { Write-Warning "Could not parse '$userDateInput' as a date. Please try again." }
				} until ($valid)
				$session.Date = $parsedDate
			}

			$session.TechFocusArea = $TechnologyFocusArea

			$eventDetailHtml = ConvertFrom-Html (Invoke-RestMethod $session.Url)
			$descriptionXPath = './/div[@class="ibox-content"]/h4[contains(text(),"Description")]/following-sibling::p'
			$session.Description = $eventDetailHtml.SelectSingleNode($descriptionXPath).InnerText.Trim()
			| ConvertFrom-HtmlEncoding
			| Assert-ScreenScrapeData 'Session Description'

			[PSCustomObject]$Session
		} catch {
			Write-Warning "Invalid session data detected: $PSItem. Skipping this session."
		}
	}

	return $sessions
}

#region Main

#Retrieve the data we care about from the HTML response.
$eventHtml = Invoke-RestMethod 'https://sessionize.com/app/speaker/events'

$page = ConvertFrom-Html $eventHtml

$currentTab = $page.SelectSingleNode('//div[@id="tab-current"]')
| Assert-ScreenScrapeData 'Current Sessions Tab'

$currentSessions = ConvertFrom-SessionizeSessionTab $currentTab

$archiveTab = $page.SelectSingleNode('//div[@id="tab-archived"]')
| Assert-ScreenScrapeData 'Past Sessions Tab'

$archiveSessions = ConvertFrom-SessionizeSessionTab $archiveTab

Write-Verbose "Found $($currentSessions.Count) current Sessions and $($archiveSessions.Count) past Sessions."

$sessionsToUpdate = $currentSessions + $archiveSessions
| Where-Object { (-not $filterPattern) -or ($_.Name -match $filterPattern) }
| Where-Object { (-not $earliestDate) -or ($_.Date -ge $earliestDate) }

#TODO: This "upsert" logic should probably be available in the main module
$mvpModulePath = (Get-Module MicrosoftMvp).path
$InformationPreference = 'Continue'

$existingActivities = Search-MvpActivitySummary -First 10000
| Where-Object type -Like 'Speaker*'

$sessionsToUpdate | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
	Import-Module $USING:mvpModulePath
	$VerbosePreference = $USING:VerbosePreference
	$WhatIfPreference = $USING:WhatIfPreference
	$DebugPreference = $USING:DebugPreference

	$session = $PSItem
	$activityTitle = $session.Name
	$existingActivity = $USING:existingActivities
	| Where-Object title -EQ $activityTitle
	| Get-MvpActivity
	| Where-Object url -EQ $session.Url

	if ($existingActivity.count -gt 1) {
		Write-Warning "Multiple activities found for '$activityTitle' with url $($session.url). Remove one from your MVP profile to update this."
		continue
	}

	$activity = $existingActivity ? ($existingActivity | Get-MvpActivity) : $(
		$tfa = $session.TechFocusArea
		$activityParams = @{
			Title                     = $activityTitle
			Type                      = 'Speaker/Presenter at Third-party event'
			Date                      = $session.date
			Description               = $session.description.Substring(0, [Math]::Min($session.description.Length, 1000))
			TechnologyFocusArea       = $tfa[0]
			AdditionalTechnologyAreas = $tfa.Count -gt 1 ? $tfa[1..($tfa.Count - 1)] : @()
			TargetAudience            = 'Developer', 'IT Pro'
		}
		$newActivity = New-MvpActivity @activityParams
		$newActivity.url = $session.url
		$newActivity
	)

	if ($existingActivity) {
		#Workaround for whatif not working as it is supposed to in parallel
		if ($WhatIfPreference) {
			Write-Host "WhatIf: Updating speaking event '$activityTitle' with url $($session.url)"
			return
		}
		Write-Verbose "Updating speaking event '$activityTitle' with url $($session.url)"
		#Update the activity with the latest data
		$activity.Date = $session.date
		$activity.Description = $session.description
		Set-MvpActivity $activity -WhatIf:$WhatIfPreference
	} else {
		if ($WhatIfPreference) {
			Write-Host "WhatIf: Adding speaking event '$activityTitle' with url $($session.url)"
			return
		}
		Write-Verbose "Adding new speaking event '$activityTitle' with url $($session.url)"
		Add-MvpActivity $activity -WhatIf:$WhatIfPreference
	}
}
#endregion