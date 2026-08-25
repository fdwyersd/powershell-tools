# Amtrak #3 FMD/LAP desktop status panel v1.10
# Windows 7 / Windows PowerShell 2.0 compatible.
# Uses RailRat station pages rather than the ambiguous train-number page.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web

try {
    $currentProtocol = [int][Net.ServicePointManager]::SecurityProtocol
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]($currentProtocol -bor 3072)
} catch { }

$RefreshSeconds = 60
$StationBaseUrl = 'https://railrat.net/stations/'

function New-Record($Properties) {
    return New-Object PSObject -Property $Properties
}

function Get-EpochMilliseconds {
    $epoch = [DateTime]::Parse('1970-01-01T00:00:00Z').ToUniversalTime()
    return [int64]([DateTime]::UtcNow.Subtract($epoch).TotalMilliseconds)
}

function Get-FullExceptionText($ErrorRecord) {
    $parts = @()
    $e = $ErrorRecord.Exception
    while ($e) {
        if ($e.Message -and -not ($parts -contains $e.Message)) { $parts += $e.Message }
        $e = $e.InnerException
    }
    if ($parts.Count -eq 0) { return [string]$ErrorRecord }
    return ($parts -join ' | ')
}

function Get-HtmlWithWebClient([string]$Url) {
    $client = New-Object System.Net.WebClient
    try {
        $client.Headers['User-Agent'] = 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 Chrome/109 Safari/537.36'
        $client.Headers['Cache-Control'] = 'no-cache'
        $client.Headers['Pragma'] = 'no-cache'
        $client.Encoding = [Text.Encoding]::UTF8
        if ([Net.WebRequest]::DefaultWebProxy) {
            [Net.WebRequest]::DefaultWebProxy.Credentials = [Net.CredentialCache]::DefaultCredentials
            $client.Proxy = [Net.WebRequest]::DefaultWebProxy
        }
        return $client.DownloadString($Url)
    } finally {
        $client.Dispose()
    }
}

function Get-HtmlWithWinHttp([string]$Url) {
    $request = New-Object -ComObject 'WinHttp.WinHttpRequest.5.1'
    try { $request.Option(9) = 2048 } catch { }
    try { $request.Option(4) = 13056 } catch { }
    $request.SetTimeouts(10000,10000,15000,20000)
    $request.Open('GET',$Url,$false)
    $request.SetRequestHeader('User-Agent','Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 Chrome/109 Safari/537.36')
    $request.SetRequestHeader('Cache-Control','no-cache')
    $request.Send()
    if ([int]$request.Status -lt 200 -or [int]$request.Status -ge 400) {
        throw ('RailRat returned HTTP ' + [string]$request.Status + ' ' + [string]$request.StatusText)
    }
    return [string]$request.ResponseText
}

function Convert-HtmlToText([string]$Html) {
    $Html = [regex]::Replace($Html, '(?is)<script.*?</script>|<style.*?</style>|<noscript.*?</noscript>', '')
    $Html = [regex]::Replace($Html, '(?i)<br\s*/?>|</p>|</div>|</li>|</tr>|</h[1-6]>|</section>|</article>|</summary>|</details>', "`n")
    $text = [regex]::Replace($Html, '(?s)<[^>]+>', ' ')
    $text = [System.Web.HttpUtility]::HtmlDecode($text)
    $text = $text -replace "`r", "`n"
    $text = $text -replace [char]0xA0, ' '
    $text = $text -replace '[ \t]+', ' '
    $text = $text -replace " *`n *", "`n"
    $text = $text -replace "`n{3,}", "`n`n"
    return $text.Trim()
}

function Get-PageText([string]$Url) {
    $urlWithCacheBust = $Url + '?_=' + (Get-EpochMilliseconds)
    $webClientError = $null
    try {
        $html = Get-HtmlWithWebClient $urlWithCacheBust
    } catch {
        $webClientError = Get-FullExceptionText $_
        try {
            $html = Get-HtmlWithWinHttp $urlWithCacheBust
        } catch {
            $winHttpError = Get-FullExceptionText $_
            throw ('WebClient: ' + $webClientError + ' || WinHTTP: ' + $winHttpError)
        }
    }
    return Convert-HtmlToText $html
}

function Convert-ClockToMinutes([string]$Clock) {
    if ($Clock -notmatch '^(\d{1,2}):(\d{2})$') { return $null }
    return ([int]$Matches[1] * 60) + [int]$Matches[2]
}

function Get-TimeDifferenceMinutes([string]$Scheduled,[string]$Observed) {
    $a = Convert-ClockToMinutes $Scheduled
    $b = Convert-ClockToMinutes $Observed
    if ($a -eq $null -or $b -eq $null) { return 0 }
    $d = $b - $a
    if ($d -gt 720) { $d -= 1440 }
    if ($d -lt -720) { $d += 1440 }
    return [int]$d
}

function Format-Delay([int]$Minutes) {
    if ($Minutes -eq 0) { return 'on time' }
    $n = [Math]::Abs($Minutes)
    if ($n -ge 60) {
        $text = ('{0}h {1}m' -f [Math]::Floor($n / 60), ($n % 60))
    } else {
        $text = ('{0}m' -f $n)
    }
    if ($Minutes -gt 0) { return ($text + ' late') }
    return ($text + ' early')
}

function Get-Section([string]$Text,[string]$StartMarker,[string]$EndMarker) {
    $start = $Text.IndexOf($StartMarker, [StringComparison]::OrdinalIgnoreCase)
    if ($start -lt 0) { return '' }
    $start += $StartMarker.Length
    $end = $Text.IndexOf($EndMarker, $start, [StringComparison]::OrdinalIgnoreCase)
    if ($end -lt 0) { $end = $Text.Length }
    return $Text.Substring($start, $end - $start)
}

function Get-Train3Entries([string]$Section,[string]$SectionName) {
    $entries = @()
    # Do not call this variable $matches. PowerShell variable names are case-insensitive,
    # and every use of the -match operator overwrites the automatic $Matches variable.
    $trainMatches = [regex]::Matches($Section, '(?im)^.*Southwest Chief\s+3\b.*$')
    for ($i = 0; $i -lt $trainMatches.Count; $i++) {
        $begin = $trainMatches[$i].Index
        if ($i + 1 -lt $trainMatches.Count) { $end = $trainMatches[$i + 1].Index }
        else { $end = $Section.Length }
        $length = [Math]::Min(($end - $begin), 1200)
        $chunk = $Section.Substring($begin, $length)

        # Reject an accidental match to a different route or a link elsewhere.
        if ($chunk -notmatch 'Chicago Union\s*\[CHI\]') { continue }
        if ($chunk -notmatch 'Los Angeles Union\s*\[LAX\]') { continue }

        $header = ([string]$trainMatches[$i].Value).Trim()
        $scheduled = $null
        $displayTime = $null
        $kind = $null
        $dateText = $null
        $departure = $null

        if ($chunk -match '(?im)Ar\s+sch\.\s*(\d{1,2}:\d{2})(?:\s*,\s*(est|act)\.\s*(\d{1,2}:\d{2})(?:\s+(\d{1,2}/\d{1,2}))?)?') {
            $scheduled = $Matches[1]
            if ($Matches[2]) { $kind = $Matches[2].ToLowerInvariant() }
            if ($Matches[3]) { $displayTime = $Matches[3] } else { $displayTime = $scheduled }
            if ($Matches[4]) { $dateText = $Matches[4] }
        }
        if ($chunk -match '(?im)Dp\s+sch\.\s*\d{1,2}:\d{2}(?:\s*,\s*(?:est|act)\.\s*(\d{1,2}:\d{2}))?') {
            if ($Matches[1]) { $departure = $Matches[1] }
        }
        if (-not $scheduled) { continue }

        $delay = Get-TimeDifferenceMinutes $scheduled $displayTime
        # Prefer RailRat's explicit status phrase when it survived HTML conversion.
        if ($header -match '(\d+)h\s*(\d+)m\s*(?:lt|late)') {
            $delay = ([int]$Matches[1] * 60) + [int]$Matches[2]
        } elseif ($header -match '(\d+)m\s*(?:lt|late)') {
            $delay = [int]$Matches[1]
        } elseif ($header -match '(\d+)h\s*(\d+)m\s*(?:erly|early)') {
            $delay = -(([int]$Matches[1] * 60) + [int]$Matches[2])
        } elseif ($header -match '(\d+)m\s*(?:erly|early)') {
            $delay = -[int]$Matches[1]
        } elseif ($header -match 'on\s*tm|on\s*time') {
            $delay = 0
        }

        $entries += New-Record @{
            Section=$SectionName; Header=$header; Scheduled=$scheduled;
            Time=$displayTime; Kind=$kind; DateText=$dateText;
            Departure=$departure; Delay=$delay; Raw=$chunk
        }
    }
    return $entries
}

function Get-DateRank([string]$DateText) {
    if ([string]::IsNullOrEmpty($DateText)) { return 0 }
    if ($DateText -notmatch '^(\d{1,2})/(\d{1,2})$') { return 0 }
    $month = [int]$Matches[1]
    $day = [int]$Matches[2]
    $year = (Get-Date).Year
    try {
        $candidate = Get-Date -Year $year -Month $month -Day $day -Hour 12 -Minute 0 -Second 0
        $now = Get-Date
        if ($candidate.AddMonths(6) -lt $now) { $candidate = $candidate.AddYears(1) }
        if ($candidate.AddMonths(-6) -gt $now) { $candidate = $candidate.AddYears(-1) }
        return [int]($candidate - [DateTime]::MinValue).TotalDays
    } catch { return 0 }
}

function Get-MinutesUntilStationTime([string]$Clock,[string]$DateText) {
    if ([string]::IsNullOrEmpty($Clock)) { return $null }
    if ($Clock -notmatch '^(\d{1,2}):(\d{2})$') { return $null }
    $hour = [int]$Matches[1]
    $minute = [int]$Matches[2]

    $nowCentral = $null
    $centralZone = $null
    try {
        $centralZone = [TimeZoneInfo]::FindSystemTimeZoneById('Central Standard Time')
        $nowCentral = [TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow,$centralZone)
    } catch {
        return $null
    }

    $month = $nowCentral.Month
    $day = $nowCentral.Day
    $year = $nowCentral.Year
    if (-not [string]::IsNullOrEmpty($DateText) -and $DateText -match '^(\d{1,2})/(\d{1,2})$') {
        $month = [int]$Matches[1]
        $day = [int]$Matches[2]
    }

    try {
        $arrivalText = ('{0:0000}-{1:00}-{2:00} {3:00}:{4:00}:00' -f $year,$month,$day,$hour,$minute)
        $localArrival = [DateTime]::ParseExact($arrivalText,'yyyy-MM-dd HH:mm:ss',[Globalization.CultureInfo]::InvariantCulture)
        $localArrival = [DateTime]::SpecifyKind($localArrival,[DateTimeKind]::Unspecified)
        if ($localArrival.AddMonths(6) -lt $nowCentral) { $localArrival = $localArrival.AddYears(1) }
        if ($localArrival.AddMonths(-6) -gt $nowCentral) { $localArrival = $localArrival.AddYears(-1) }
        $arrivalUtc = [TimeZoneInfo]::ConvertTimeToUtc($localArrival,$centralZone)
        return [int][Math]::Floor(($arrivalUtc - [DateTime]::UtcNow).TotalMinutes)
    } catch {
        return $null
    }
}


function Format-TimeToArrival($MinutesAway) {
    if ($MinutesAway -eq $null -or $MinutesAway -lt 0) { return '' }
    $total = [int]$MinutesAway
    if ($total -lt 60) { return ('(+' + $total + 'm)') }
    $hours = [int][Math]::Floor($total / 60)
    $minutes = $total % 60
    if ($minutes -eq 0) { return ('(+' + $hours + 'h)') }
    return ('(+' + $hours + 'h' + $minutes + 'm)')
}

function Select-BestStationEntry($Entries) {
    if (-not $Entries -or $Entries.Count -eq 0) { return $null }
    foreach ($entry in $Entries) {
        $entry | Add-Member -MemberType NoteProperty -Name DateRank -Value (Get-DateRank $entry.DateText)
        $sectionRank = 0
        if ($entry.Section -eq 'Arriving') { $sectionRank = 2 }
        elseif ($entry.Section -eq 'Departed') { $sectionRank = 1 }
        $entry | Add-Member -MemberType NoteProperty -Name SectionRank -Value $sectionRank
    }
    # An arriving record always wins. Within a section, use the newest service date.
    return $Entries | Sort-Object SectionRank,DateRank -Descending | Select-Object -First 1
}

function Get-StationStatus([string]$Code) {
    $url = $StationBaseUrl + $Code + '/'
    $text = Get-PageText $url
    $arriving = Get-Section $text 'Arriving Trains' 'Departed Trains'
    $departed = Get-Section $text 'Departed Trains' 'Useful Links'
    $entries = @()
    if ($arriving) { $entries += @(Get-Train3Entries $arriving 'Arriving') }
    if ($departed) { $entries += @(Get-Train3Entries $departed 'Departed') }
    $best = Select-BestStationEntry $entries
    if (-not $best) {
        return New-Record @{ Main='No #3 listed'; Detail='Station page has no Southwest Chief #3'; IsError=$false; MinutesAway=$null }
    }

    $minutesAway = $null
    if ($best.Section -eq 'Arriving' -and $best.Kind -ne 'act') {
        $minutesAway = Get-MinutesUntilStationTime $best.Time $best.DateText
    }

    if ($best.Section -eq 'Departed' -or $best.Kind -eq 'act') {
        $main = 'Arrived ' + $best.Time + ' CT'
        if ($best.Departure) { $detail = 'Departed ' + $best.Departure + ' - ' + (Format-Delay $best.Delay) }
        else { $detail = (Format-Delay $best.Delay) }
    } else {
        $main = 'Est. arrival ' + $best.Time + ' CT'
        $countdown = Format-TimeToArrival $minutesAway
        if ($countdown) { $main += ' ' + $countdown }
        $detail = Format-Delay $best.Delay
    }
    if ($best.DateText) { $detail += ' - ' + $best.DateText }
    return New-Record @{ Main=$main; Detail=$detail; IsError=$false; MinutesAway=$minutesAway }
}

# ----- UI -----
$form = New-Object Windows.Forms.Form
$form.Text = 'Amtrak #3 - FMD / LAP v1.11'
$form.Size = New-Object Drawing.Size(390,178)
$form.MinimumSize = New-Object Drawing.Size(390,178)
$form.MaximumSize = New-Object Drawing.Size(390,178)
$form.StartPosition = 'Manual'
$form.Location = New-Object Drawing.Point(([Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Right - 410), 40)
$form.TopMost = $true
$form.FormBorderStyle = 'None'
$form.ShowInTaskbar = $false
$form.KeyPreview = $true
$form.BackColor = [Drawing.Color]::FromArgb(28,30,34)
$form.ForeColor = [Drawing.Color]::White

$title = New-Object Windows.Forms.Label
$title.Text = 'Southwest Chief #3'
$title.Font = New-Object Drawing.Font('Segoe UI',14,[Drawing.FontStyle]::Bold)
$title.Location = New-Object Drawing.Point(14,10)
$title.Size = New-Object Drawing.Size(346,28)
$form.Controls.Add($title)

function New-StationPanel([string]$Code,[string]$Name,[int]$Y) {
    $box = New-Object Windows.Forms.Panel
    $box.Location = New-Object Drawing.Point(14,$Y)
    $box.Size = New-Object Drawing.Size(346,58)
    $box.BackColor = [Drawing.Color]::FromArgb(42,45,51)
    $labCode = New-Object Windows.Forms.Label
    $labCode.Text = $Code
    $labCode.Font = New-Object Drawing.Font('Segoe UI',12,[Drawing.FontStyle]::Bold)
    $labCode.Location = New-Object Drawing.Point(9,7)
    $labCode.Size = New-Object Drawing.Size(48,22)
    $labName = New-Object Windows.Forms.Label
    $labName.Text = $Name
    $labName.Font = New-Object Drawing.Font('Segoe UI',8)
    $labName.ForeColor = [Drawing.Color]::Silver
    $labName.Location = New-Object Drawing.Point(9,31)
    $labName.Size = New-Object Drawing.Size(95,18)
    $main = New-Object Windows.Forms.Label
    $main.Font = New-Object Drawing.Font('Segoe UI',10,[Drawing.FontStyle]::Bold)
    $main.Location = New-Object Drawing.Point(102,7)
    $main.Size = New-Object Drawing.Size(234,22)
    $detail = New-Object Windows.Forms.Label
    $detail.Font = New-Object Drawing.Font('Segoe UI',9)
    $detail.ForeColor = [Drawing.Color]::LightGray
    $detail.Location = New-Object Drawing.Point(102,31)
    $detail.Size = New-Object Drawing.Size(234,18)
    $box.Controls.Add($labCode)
    $box.Controls.Add($labName)
    $box.Controls.Add($main)
    $box.Controls.Add($detail)
    $form.Controls.Add($box)
    return New-Record @{ Main=$main; Detail=$detail; Panel=$box; Code=$labCode; Name=$labName }
}

$fmdUI = New-StationPanel 'FMD' 'Fort Madison' 43
$lapUI = New-StationPanel 'LAP' 'La Plata' 107

$normalCardColor = [Drawing.Color]::FromArgb(42,45,51)
$normalMainColor = [Drawing.Color]::White
$normalDetailColor = [Drawing.Color]::LightGray
$normalNameColor = [Drawing.Color]::Silver

function Get-UrgencyState($MinutesAway) {
    if ($MinutesAway -ne $null -and $MinutesAway -ge 0) {
        if ($MinutesAway -le 10) { return 'Red' }
        if ($MinutesAway -le 30) { return 'Orange' }
        if ($MinutesAway -le 60) { return 'Yellow' }
    }
    return 'Normal'
}

function Apply-StationCardState($UI,[string]$State) {
    $background = $normalCardColor
    $foreground = $normalMainColor
    $detailColor = $normalDetailColor
    $nameColor = $normalNameColor

    if ($State -eq 'Red') {
        $background = [Drawing.Color]::FromArgb(220,55,45)
        $foreground = [Drawing.Color]::White
        $detailColor = [Drawing.Color]::White
        $nameColor = [Drawing.Color]::White
    } elseif ($State -eq 'Orange') {
        $background = [Drawing.Color]::FromArgb(243,133,32)
        $foreground = [Drawing.Color]::Black
        $detailColor = [Drawing.Color]::FromArgb(35,35,35)
        $nameColor = [Drawing.Color]::FromArgb(35,35,35)
    } elseif ($State -eq 'Yellow') {
        $background = [Drawing.Color]::FromArgb(255,214,64)
        $foreground = [Drawing.Color]::Black
        $detailColor = [Drawing.Color]::FromArgb(35,35,35)
        $nameColor = [Drawing.Color]::FromArgb(35,35,35)
    }

    $UI.Panel.BackColor = $background
    $UI.Code.ForeColor = $foreground
    $UI.Main.ForeColor = $foreground
    $UI.Detail.ForeColor = $detailColor
    $UI.Name.ForeColor = $nameColor
}

# Track urgency separately for each card. The first update establishes the
# starting state without flashing. Later entries into a new warning color flash
# between that warning color and the normal card color for 15 seconds.
$script:cardState = @{ FMD=$null; LAP=$null }
$script:flashUntil = @{ FMD=$null; LAP=$null }
$script:flashVisible = @{ FMD=$true; LAP=$true }
$script:cardUI = @{ FMD=$fmdUI; LAP=$lapUI }

function Set-StationCardColor([string]$Code,$UI,$MinutesAway) {
    $newState = Get-UrgencyState $MinutesAway
    $oldState = $script:cardState[$Code]
    $script:cardState[$Code] = $newState

    if ($oldState -eq $null) {
        Apply-StationCardState $UI $newState
        return
    }

    if ($oldState -ne $newState -and $newState -ne 'Normal') {
        $script:flashUntil[$Code] = [DateTime]::Now.AddSeconds(15)
        $script:flashVisible[$Code] = $true
        Apply-StationCardState $UI $newState
        return
    }

    # Do not restart a running flash merely because the 60-second refresh ran.
    if ($script:flashUntil[$Code] -ne $null -and [DateTime]::Now -lt $script:flashUntil[$Code]) {
        return
    }

    $script:flashUntil[$Code] = $null
    Apply-StationCardState $UI $newState
}

$flashTimer = New-Object Windows.Forms.Timer
$flashTimer.Interval = 500
$flashTimer.Add_Tick({
    foreach ($code in @('FMD','LAP')) {
        $until = $script:flashUntil[$code]
        if ($until -eq $null) { continue }

        $ui = $script:cardUI[$code]
        if ([DateTime]::Now -ge $until) {
            $script:flashUntil[$code] = $null
            $script:flashVisible[$code] = $true
            Apply-StationCardState $ui $script:cardState[$code]
        } else {
            $script:flashVisible[$code] = -not $script:flashVisible[$code]
            if ($script:flashVisible[$code]) {
                Apply-StationCardState $ui $script:cardState[$code]
            } else {
                Apply-StationCardState $ui 'Normal'
            }
        }
    }
})
$flashTimer.Start()

# Right-click menu for the undecorated window.
$contextMenu = New-Object Windows.Forms.ContextMenuStrip
$refreshMenuItem = New-Object Windows.Forms.ToolStripMenuItem
$refreshMenuItem.Text = 'Refresh'
$exitMenuItem = New-Object Windows.Forms.ToolStripMenuItem
$exitMenuItem.Text = 'Exit'
[void]$contextMenu.Items.Add($refreshMenuItem)
[void]$contextMenu.Items.Add($exitMenuItem)

# Borderless window: drag it from anywhere with the left mouse button.
$script:dragging = $false
$script:dragStartMouse = New-Object Drawing.Point(0,0)
$script:dragStartForm = New-Object Drawing.Point(0,0)

$mouseDownHandler = {
    param($sender,$e)
    if ($e.Button -eq [Windows.Forms.MouseButtons]::Left) {
        $script:dragging = $true
        $script:dragStartMouse = [Windows.Forms.Cursor]::Position
        $script:dragStartForm = $form.Location
    }
}
$mouseMoveHandler = {
    param($sender,$e)
    if ($script:dragging) {
        $p = [Windows.Forms.Cursor]::Position
        $dx = $p.X - $script:dragStartMouse.X
        $dy = $p.Y - $script:dragStartMouse.Y
        $form.Location = New-Object Drawing.Point(($script:dragStartForm.X + $dx),($script:dragStartForm.Y + $dy))
    }
}
$mouseUpHandler = {
    param($sender,$e)
    $script:dragging = $false
}

$dragControls = @($form,$title,$fmdUI.Panel,$fmdUI.Main,$fmdUI.Detail,$fmdUI.Code,$fmdUI.Name,$lapUI.Panel,$lapUI.Main,$lapUI.Detail,$lapUI.Code,$lapUI.Name)
foreach ($control in $dragControls) {
    $control.Add_MouseDown($mouseDownHandler)
    $control.Add_MouseMove($mouseMoveHandler)
    $control.Add_MouseUp($mouseUpHandler)
    $control.ContextMenuStrip = $contextMenu
}

# Escape closes the otherwise undecorated panel.
$form.Add_KeyDown({ param($sender,$e) if ($e.KeyCode -eq [Windows.Forms.Keys]::Escape) { $form.Close() } })

$script:busy = $false
function Update-TrainPanel {
    if ($script:busy) { return }
    $script:busy = $true
    try {
        $f = Get-StationStatus 'FMD'
        $fmdUI.Main.Text = $f.Main
        $fmdUI.Detail.Text = $f.Detail
        Set-StationCardColor 'FMD' $fmdUI $f.MinutesAway
    } catch {
        $fmdUI.Main.Text = 'Unable to update'
        $where = ''
        if ($_.InvocationInfo -and $_.InvocationInfo.ScriptLineNumber) { $where = 'Line ' + $_.InvocationInfo.ScriptLineNumber + ': ' }
        $fmdUI.Detail.Text = $where + (Get-FullExceptionText $_)
        Set-StationCardColor 'FMD' $fmdUI $null
    }
    try {
        $l = Get-StationStatus 'LAP'
        $lapUI.Main.Text = $l.Main
        $lapUI.Detail.Text = $l.Detail
        Set-StationCardColor 'LAP' $lapUI $l.MinutesAway
    } catch {
        $lapUI.Main.Text = 'Unable to update'
        $where = ''
        if ($_.InvocationInfo -and $_.InvocationInfo.ScriptLineNumber) { $where = 'Line ' + $_.InvocationInfo.ScriptLineNumber + ': ' }
        $lapUI.Detail.Text = $where + (Get-FullExceptionText $_)
        Set-StationCardColor 'LAP' $lapUI $null
    }
    $script:busy = $false
}

$refreshMenuItem.Add_Click({ Update-TrainPanel })
$exitMenuItem.Add_Click({ $form.Close() })

$timer = New-Object Windows.Forms.Timer
$timer.Interval = $RefreshSeconds * 1000
$timer.Add_Tick({ Update-TrainPanel })
$form.Add_Shown({ Update-TrainPanel; $timer.Start() })
$form.Add_FormClosed({ $timer.Stop(); $timer.Dispose(); $flashTimer.Stop(); $flashTimer.Dispose() })
[void]$form.ShowDialog()
