function Get-CWCSession {
    [CmdletBinding()]
    param (
        [ValidateSet('Support', 'Access', 'Meeting')]
        $Type = 'Access',
        [string]$Group = 'All Machines',
        [Alias('Filter')]
        [string]$Search,
        [string]$Name,
        [string]$FindSessionID,
        [int]$Limit
    )

    $Endpoint = 'Services/PageService.ashx/GetLiveData'

    switch ($Type) {
        'Support' { $Number = 0 }
        'Meeting' { $Number = 1 }
        'Access' { $Number = 2 }
        default { return Write-Error "Unknown Type, $Type" }
    }

    if( -not [string]::IsNullOrWhiteSpace($Name) -and [string]::IsNullOrWhiteSpace($Search) ) {
        $Search = $Name
    }

    $Body = ConvertTo-Json @(
        @{
            HostSessionInfo  = @{
                'sessionType'           = $Number
                'sessionGroupPathParts' = @($Group)
                'filter'                = $Search
                'findSessionID'         = $FindSessionID
                'sessionLimit'          = $Limit
            }            
            ActionCenterInfo = @{}
        }
        0
    ) -Depth 5
    Write-Verbose $Body

    $WebRequestArguments = @{
        Endpoint = $Endpoint
        Body     = $Body
        Method   = 'Post'
    }

    $Data = Invoke-CWCWebRequest -Arguments $WebRequestArguments
    $Sessions = $Data.ResponseInfoMap.HostSessionInfo.Sessions

    if( -not [string]::IsNullOrWhiteSpace($Name) ){
        $Sessions = $Sessions | Where-Object { $_.Name -eq $Name }
    }

    $Sessions
    | Add-Member -Name 'ConnectedOperators' -MemberType ScriptProperty -Value { ($this.ActiveConnections | Measure-Object).Count } -PassThru
    | Add-PSType -TypeName 'CWC.Session' -DefaultDisplayPropertySet Name, GuestLoggedOnUserName, GuestOperatingSystemName, ConnectedOperators

}
