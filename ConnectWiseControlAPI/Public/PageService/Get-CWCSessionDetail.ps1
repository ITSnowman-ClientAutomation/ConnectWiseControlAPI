function Get-CWCSessionDetail {
    [CmdletBinding()]
    param (
        [string]$Group = 'All Machines',

        [Parameter(ValueFromPipeline=$true, Mandatory=$true, ParameterSetName='SessionSet')]
        [PSTypeName('CWC.Session')]
        $Session,

        [Parameter(Mandatory=$true, ParameterSetName='GUIDSet')]
        [guid]
        $GUID
    )

    begin {}

    process {
        # If invoked in the SessionSet, extract SessionID -> GUID
        if ($PSCmdlet.ParameterSetName -eq 'SessionSet') {
            if ($null -ne $Session.SessionID) {
                try {
                    $GUID = [guid]$Session.SessionID
                } catch {
                    Write-Error "Session object's SessionID is not a valid GUID: $($Session.SessionID)"
                    return
                }
            } else {
                Write-Error "Pipeline object does not contain a 'SessionID' property."
                return
            }
        }

        $Endpoint = 'Services/PageService.ashx/GetSessionDetails'

        $Body = ConvertTo-Json @($Group,$GUID)
        Write-Verbose $Body

        $WebRequestArguments = @{
            Endpoint = $Endpoint
            Body = $Body
            Method = 'Post'
        }

        Invoke-CWCWebRequest -Arguments $WebRequestArguments
        | Add-PSType -TypeName 'CWC.SessionDetail' -DefaultDisplayPropertySet Session,Events,Connections,BaseTime
    }

    end {}
}