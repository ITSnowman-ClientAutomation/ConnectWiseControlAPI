# Merged in some functionalities from Justin Grote's ScreenConnect module
# https://gist.github.com/JustinGrote/9e4b984495c3d878415e17714dbeb49b#file-screenconnect-psm1-L226

using namespace Microsoft.PowerShell.Commands
using namespace System.Text

function Invoke-CWCCommand {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True, ParameterSetName = 'ByGUID')]
        [guid]$GUID,
        [Parameter(Mandatory = $True, ParameterSetName = 'ByComputerName')]
        [string]$ComputerName,
        [string]$Command,
        [int]$TimeOut = 30000,
        [int]$MaxLength = [int]::MaxValue,
        [switch]$PowerShell,
        [string]$Group = 'All Machines',
        [switch]$NoWait
    )

    # Add-Type -AssemblyName System.Management.Automation

    $Endpoint = 'Services/PageService.ashx/AddSessionEvents'

    if ($ComputerName) {
        try {
            $session = Get-CWCSession -Name $ComputerName -ErrorAction Stop
        } catch {
            throw "Could not resolve ComputerName '$ComputerName' to a session: $($_.Exception.Message)"
        }
        if (-not $session) { throw "No session returned for '$ComputerName'." }
        $GUID = [guid]$session.sessionId
    }

    $origin = New-Object -Type DateTime -ArgumentList 1970, 1, 1, 0, 0, 0, 0
    $SessionEventType = 44

    #HACK: Connectwise does not give us a unique ID to correlate a command run output to its invocation, so we force output this in any script so that we can correlate the command with the response and avoid a race condition. This is necessary in case two commands with the same format are run in quick succession. For example: sleep (get-random -max 30) if run twice, would return the wrong result if comparing by date if the second invocation completes first. We will trim it off after invocation so the result will be the same.
    $invocationHint = "[[INVOCATIONID:$(New-Guid)]]"


    # Format command
    $FormattedCommand = @()
    if ($Powershell) { $FormattedCommand += '#!ps' }
    $FormattedCommand += "#timeout=$TimeOut"
    $FormattedCommand += "#maxlength=$MaxLength"
    if ($PowerShell) {
        $FormattedCommand += 'try {'
        $FormattedCommand += '  & powershell -nop -noni -o xml -c {'
        $FormattedCommand += '    $ProgressPreference = "SilentlyContinue"'
    }
    $FormattedCommand += $Command
    if ($PowerShell) {
        $FormattedCommand += '  }'
        $FormattedCommand += '} finally {'
        $FormattedCommand += "  '$invocationHint'"
        $FormattedCommand += '}'
    }else{
        $FormattedCommand += ";echo $invocationHint"
    }
    $FormattedCommand = $FormattedCommand | Out-String
    $CommandObject = @{
        SessionID = $GUID
        EventType = $SessionEventType
        Data      = $FormattedCommand
    }
    $Body = (ConvertTo-Json @(@($Group), @($CommandObject))).Replace('\r\n', '\n')
    Write-Verbose $Body

    # Issue command
    $WebRequestArguments = @{
        Endpoint = $Endpoint
        Body     = $Body
        Method   = 'Post'
    }
    $null = Invoke-CWCWebRequest -Arguments $WebRequestArguments
    if ($NoWait) { return }

    # Get Session
    try { $SessionDetails = Get-CWCSessionDetail -Group $Group -GUID $GUID }
    catch { return $_ }

    #Get time command was executed
    $epoch = $((New-TimeSpan -Start $(Get-Date -Date '01/01/1970') -End $(Get-Date)).TotalSeconds)
    $ExecuteTime = $epoch - ((($SessionDetails.events | Where-Object { $_.EventType -eq 44 })[-1]).Time / 1000)
    $ExecuteDate = $origin.AddSeconds($ExecuteTime)

    # Look for results of command
    $Looking = $True
    $TimeOutDateTime = (Get-Date).AddMilliseconds($TimeOut)
    $Body = ConvertTo-Json @($Group, $GUID)
    while ($Looking) {
        try { $SessionDetails = Get-CWCSessionDetail -Group $Group -GUID $GUID }
        catch { return $_ }

        $ConnectionsWithData = @()
        Foreach ($Connection in $SessionDetails.Events) {
            $ConnectionsWithData += $Connection | Where-Object { $_.EventType -eq 70 }
        }

        $Events = ($ConnectionsWithData | Where-Object { $_.EventType -eq 70 -and $_.Time })
        foreach ($Event in $Events) {
            $epoch = $((New-TimeSpan -Start $(Get-Date -Date '01/01/1970') -End $(Get-Date)).TotalSeconds)
            $CheckTime = $epoch - ($Event.Time / 1000)
            $CheckDate = $origin.AddSeconds($CheckTime)
            if ($CheckDate -gt $ExecuteDate) {
                $Looking = $False
                Write-Verbose $Event.Data

                if( $PowerShell ) {
                    $cliXmlOutput = $Event.Data -replace [regex]::Escape($invocationHint) -replace [regex]::Escape('#< CLIXML')
                    $outputObjects = $cliXmlOutput -split '(?=<Objs\b)' |
                        where-object { -not [string]::IsNullOrEmpty($_.Trim() ) } |
                        forEach-Object {
                            try { [System.Management.Automation.PSSerializer]::Deserialize($_) } catch { "$PSItem" }
                    }
                    for ($i = 0; $i -lt $outputObjects.Count; $i++) {
                        for ($j = 0; $j -lt $outputObjects[$i].PSObject.TypeNames.Count; $j++) {
                            if ($outputObjects[$i].PSObject.TypeNames[$j] -like 'Deserialized.*') {
                                $outputObjects[$i].PSObject.TypeNames[$j] = $outputObjects[$i].PSObject.TypeNames[$j] -replace '^Deserialized\.', ''
                            }
                        }
                    }

                    for ($i = 0; $i -lt $outputObjects.Count; $i++) {
                        if( $outputObjects[$i].PSObject.TypeNames -contains 'System.IO.FileSystemInfo'){
                            $outputObjects[$i] = Get-Item -LiteralPath $outputObjects[$i].FullName
                        }
                    }

                    return $outputObjects
                }

                $Output = $Event.Data -split '[\r\n]' | Where-Object {
                    $_ -and
                    $_.Trim() -ine "C:\Windows\System32>$Command" -and
                    $_.Trim() -ine "C:\Windows\System32>echo $invocationHint" -and
                    $_.Trim() -ine $invocationHint
                }
                return $Output
            }
        }

        Start-Sleep -Seconds 1
        if ($(Get-Date) -gt $TimeOutDateTime.AddSeconds(1)) {
            $Looking = $False
            Write-Warning 'Command timed out.'
        }
    }
}