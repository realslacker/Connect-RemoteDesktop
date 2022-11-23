function Connect-RemoteDesktop {
    
    [CmdletBinding( DefaultParameterSetName = 'ComputerName' )]
    [Alias( 'rdp' )]
    param(
            
        [Parameter( ParameterSetName = 'ComputerName', Mandatory, Position = 1 )]
        [Alias( 'Server' )]
        [string]
        $ComputerName,

        [Parameter( ParameterSetName = 'ComputerName' )]
        [ValidateRange( 1, 65535 )]
        [int]
        $Port = 3389,

        [Parameter( ParameterSetName = 'ComputerName' )]
        [string]
        $Gateway,

        [Parameter( ParameterSetName = 'ComputerName' )]
        [Alias( 'FS' )]
        [switch]
        $FullScreen,

        [Parameter( ParameterSetName = 'ComputerName' )]
        [ValidateRange( 640, 4096 )]
        [int]
        $Width,

        [Parameter( ParameterSetName = 'ComputerName' )]
        [ValidateRange( 480, 2048 )]
        [int]
        $Height,

        [Parameter( ParameterSetName = 'ComputerName' )]
        [switch]
        $AdminSession,

        [Parameter( ParameterSetName = 'ComputerName' )]
        [switch]
        $PublicMode,

        [Parameter( ParameterSetName = 'ComputerName' )]
        [switch]
        $SpanMonitors,

        [Parameter( ParameterSetName = 'ComputerName' )]
        [switch]
        $MultiMonitor,

        [Parameter( ParameterSetName = 'ConnectionFile', Mandatory )]
        [Alias( 'Path' )]
        [string]
        $ConnectionFile,
        
        [pscredential]
        $Credential

    )

    if ( $Credential ) {
        cmdkey.exe "/generic:TERMSRV/$ComputerName" "/user:$($Credential.UserName)" "/pass:$($Credential.GetNetworkCredential().Password)" > $null
    }

    [System.Collections.Generic.List[string]]$ArgumentsList = @()

    if ( $PSBoundParameters.ContainsKey( 'FullScreen' ) -or $PSBoundParameters.ContainsKey( 'MultiMonitor' ) ) {
        $PSBoundParameters.Remove( 'Width' ) > $null
        $PSBoundParameters.Remove( 'Height' ) > $null
    }

    switch ( $PSBoundParameters.Keys ) {
        'ComputerName'   { $ArgumentsList.Add( "/v:${ComputerName}:${Port}" ) }
        'ConnectionFile' { $ArgumentsList.Add( "$ConnectionFile" ) }
        'Gateway'        { $ArgumentsList.Add( "/g:$Gateway" ) }
        'AdminSession'   { $ArgumentsList.Add( '/admin' ) }
        'FullScreen'     { $ArgumentsList.Add( '/f' ) }
        'Width'          { $ArgumentsList.Add( "/w:$Width" ) }
        'Height'         { $ArgumentsList.Add( "/h:$Height" ) }
        'PublicMode'     { $ArgumentsList.Add( '/public' ) }
        'SpanMonitors'   { $ArgumentsList.Add( '/span' ) }
        'MultiMonitor'   { $ArgumentsList.Add( '/multimon' ) }
    }

    $MstscProcess = Start-Process -FilePath mstsc.exe -ArgumentList $ArgumentsList -PassThru

    try {
        do {
            Start-Sleep -Milliseconds 100
        } until ( (Get-Process -Id $MstscProcess.Id -ErrorAction Stop).MainWindowTitle -like "$ComputerName*" )
    }
    # window closed while waiting for connection
    catch { return }


    if ( $Credential ) {
        cmdkey.exe "/delete:TERMSRV/$ComputerName" > $null
    }

}