$MstscWindowDefinition = @'
    using System;
    using System.Runtime.InteropServices;
    public class MstscWindow
    {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowRect(
            IntPtr hWnd, out MstscWindowRect lpRect);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public extern static bool MoveWindow( 
            IntPtr handle, int x, int y, int width, int height, bool redraw);

        [DllImport("user32.dll")] 
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ShowWindow(
            IntPtr handle, int state);
    }
    public struct MstscWindowRect
    {
        public int Left;        // x position of upper-left corner
        public int Top;         // y position of upper-left corner
        public int Right;       // x position of lower-right corner
        public int Bottom;      // y position of lower-right corner
    }
'@

if ( -not ([System.Management.Automation.PSTypeName]'MstscWindow').Type ) {
    Add-Type -TypeDefinition $MstscWindowDefinition
}

if ( -not ([System.Management.Automation.PSTypeName]'System.Windows.Forms').Type ) {
    Add-Type -AssemblyName System.Windows.Forms
}

enum AudioModeEnum {
    PlayOnLocalComputer            = 0 # Play sounds on the local computer.
    PlayOnRemoteComputer           = 1 # Play sounds on the remote computer.
    DoNotPlay                      = 2 # Do not play sounds.
}
enum AudioQualityModeEnum {
    Dynamic                        = 0 # Dynamically adjust audio quality based on available bandwidth.
    Medium                         = 1 # Always use medium audio quality.
    Uncompressed                   = 2 # Always use uncompressed audio quality.
}
enum AuthenticationLevelEnum {
    ConnectWithoutWarning          = 0 # If server authentication fails, connect without giving a warning.
    DoNotConnect                   = 1 # If server authentication fails, do not connect.
    WarnAndAllowConnection         = 2 # If server authentication fails, show a warning and allow the user to connect or not.
    None                           = 3 # Server authentication is not required.
}
enum ConnectionTypeEnum {
    Modem                          = 1 # Modem (56 Kbps).
    LowSpeedBroadband              = 2 # Low-speed broadband (256 Kbps - 2 Mbps).
    Satellite                      = 3 # Satellite (2 Mbps - 16 Mbps with high latency).
    HighSpeedBroadband             = 4 # High-speed broadband (2 Mbps - 10 Mbps).
    Wan10MbpsOrHigher              = 5 # WAN (10 Mbps or higher with high latency).
    Lan10MbpsOrHigher              = 6 # LAN (10 Mbps or higher).
    AutoDetect                     = 7 # Automatic bandwidth detection. Requires <b>bandwidthautodetect<b>.</b></b>
}

enum GatewayBrokeringType {
    Undefined                     = 0 # Undefined
}

enum GatewayCredentialsSourceEnum {
    AskForPassword                 = 0 # Ask for password (NTLM).
    SmartCard                      = 1 # Use smart card.
    SelectLater                    = 4 # Allow user to select later.
}
enum GatewayProfileUsageMethodEnum {
    UseDefaultSettings             = 0 # Use the default profile mode, as specified by the administrator.
    UseExplicitSettings            = 1 # Use explicit settings.
}
enum GatewayUsageMethodEnum {
    AlwaysUse                      = 1 # Always use an RD Gateway, even for local connections.
    BypassForLocalAddresses        = 2 # Use the RD Gateway if a direct connection cannot be made to the remote computer (i.e. bypass for local addresses).
    UseDefaultSettings             = 3 # Use the default RD Gateway settings.
    DoNotUse                       = 4 # Do not use an RD Gateway server.
}
enum KeyboardHookEnum {
    LocalComputer                  = 0 # Windows key combinations are applied on the local computer.
    RemoteComputer                 = 1 # Windows key combinations are applied on the remote computer.
    RemoteFullScreenModeOnly       = 2 # Windows key combinations are applied in full-screen mode only.
}
enum RedirectedVideoCaptureEncodingQualityEnum {
    LowQuality                     = 0 # High compression video. Quality may suffer when there's a lot of motion.
    MediumQuality                  = 1 # Medium compression.
    HighQuality                    = 2 # Low compression video with high picture quality.
}
enum ScreenModeIdEnum {
    Windowed                       = 1 # The remote session will appear in a window.
    FullScreen                     = 2 # The remote session will appear full screen.
}
enum SessionBppEnum {
    VGA256Colors                   = 8 # 256 colors (8 bit).
    HighColor15Bit                 = 15 # High color (15 bit).
    HighColor16Bit                 = 16 # High color (16 bit).
    TrueColor24Bit                 = 24 # True color (24 bit).
    TrueColor32Bit                 = 32 # Highest quality (32 bit).
}
enum VideoPlaybackModeEnum {
    NotOptimized                   = 0 # Do not use RDP efficient multimedia streaming for video playback.
    Optimized                      = 1 # Use RDP efficient multimedia streaming for video playback when possible.
}

$ParameterToSettingMap = @{
    ComputerName                          = 'full address:s:'
    Port                                  = 'server port:i:'
    AdministrativeSession                 = 'administrative session:i:'
    AllowDesktopComposition               = 'allow desktop composition:i:'
    AllowFontSmoothing                    = 'allow font smoothing:i:'
    AlternateFullAddress                  = 'alternate full address:s:'
    AlternateShell                        = 'alternate shell:s:'
    AudioCapture                          = 'audiocapturemode:i:'
    AudioMode                             = 'audiomode:i:'
    AudioQualityMode                      = 'audioqualitymode:i:'
    AuthenticationLevel                   = 'authentication level:i:'
    AutoReconnectMaxRetries               = 'autoreconnect max retries:i:'
    AutoReconnectionEnabled               = 'autoreconnection enabled:i:'
    BandwidthAutoDetect                   = 'bandwidthautodetect:i:'
    BitmapCachePersistEnable              = 'bitmapcachepersistenable:i:'
    BitmapCacheSize                       = 'bitmapcachesize:i:'
    CamerasToRedirect                     = 'camerastoredirect:s:'
    Compression                           = 'compression:i:'
    ConnectToConsole                      = 'connect to console:i:'
    ConnectionType                        = 'connection type:i:'
    DesktopHeight                         = 'desktopheight:i:'
    DesktopWidth                          = 'desktopwidth:i:'
    DevicesToRedirect                     = 'devicestoredirect:s:'
    DisableCtrlAltDel                     = 'disable ctrl+alt+del:i:'
    DisableCursorSetting                  = 'disable cursor setting:i:'
    DisableFullWindowDrag                 = 'disable full window drag:i:'
    DisableMenuAnimations                 = 'disable menu anims:i:'
    DisableThemes                         = 'disable themes:i:'
    DisableWallpaper                      = 'disable wallpaper:i:'
    DisableConnectionSharing              = 'disableconnectionsharing:i:'
    DisableRemoteAppCapabilitiesCheck     = 'disableremoteappcapscheck:i:'
    DisplayConnectionBar                  = 'displayconnectionbar:i:'
    Domain                                = 'domain:s:'
    DrivesToRedirect                      = 'drivestoredirect:s:'
    EnableCredSSPSupport                  = 'enablecredsspsupport:i:'
    EnableRDSAADAuth                      = 'enablerdsaadauth:i:'
    EnableSuperPan                        = 'enablesuperpan:i:'
    EnableWorkspaceReconnect              = 'enableworkspacereconnect:i:'
    EncodeRedirectedVideoCapture          = 'encode redirected video capture:i:'
    GatewayCredentialsSource              = 'gatewaycredentialssource:i:'
    Gateway                               = 'gatewayhostname:s:'
    GatewayBrokeringType                  = 'gatewaybrokeringtype:i:'
    GatewayProfileUsageMethod             = 'gatewayprofileusagemethod:i:'
    GatewayUsageMethod                    = 'gatewayusagemethod:i:'
    KDCProxyName                          = 'kdcproxyname:s:'
    KeyboardHook                          = 'keyboardhook:i:'
    NegotiateSecurityLayer                = 'negotiate security layer:i:'
    NetworkAutoDetect                     = 'networkautodetect:i:'
    PinConnectionBar                      = 'pinconnectionbar:i:'
    PromptForCredentials                  = 'prompt for credentials:i:'
    PromptForCredentialsOnClient          = 'prompt for credentials on client:i:'
    PromptCredentialOnce                  = 'promptcredentialonce:i:'
    PublicMode                            = 'public mode:i:'
    RDGIsKDCProxy                         = 'rdgiskdcproxy:i:'
    RedirectedVideoCaptureEncodingQuality = 'redirected video capture encoding quality:i:'
    RedirectClipboard                     = 'redirectclipboard:i:'
    RedirectCOMPorts                      = 'redirectcomports:i:'
    RedirectDirectX                       = 'redirectdirectx:i:'
    RedirectDrives                        = 'redirectdrives:i:'
    RedirectLocation                      = 'redirectlocation:i:'
    RedirectPOSDevices                    = 'redirectposdevices:i:'
    RedirectPrinters                      = 'redirectprinters:i:'
    RedirectSmartCards                    = 'redirectsmartcards:i:'
    RedirectWebAuthn                      = 'redirectwebauthn:i:'
    RemoteApplicationCmdline              = 'remoteapplicationcmdline:s:'
    RemoteApplicationFile                 = 'remoteapplicationfile:s:'
    RemoteApplicationExpandCmdline        = 'remoteapplicationexpandcmdline:i:'
    RemoteApplicationExpandWorkingDir     = 'remoteapplicationexpandworkingdir:i:'
    RemoteApplicationIcon                 = 'remoteapplicationicon:s:'
    RemoteApplicationMode                 = 'remoteapplicationmode:i:'
    RemoteApplicationName                 = 'remoteapplicationname:s:'
    RemoteApplicationProgram              = 'remoteapplicationprogram:s:'
    RemoteAppMouseMoveInject              = 'remoteappmousemoveinject:i:'
    ScreenModeId                          = 'screen mode id:i:'
    SelectedMonitors                      = 'selectedmonitors:s:'
    SessionBpp                            = 'session bpp:i:'
    ShellWorkingDirectory                 = 'shell working directory:s:'
    Signature                             = 'signature:s:'
    SignScope                             = 'signscope:s:'
    SmartSizing                           = 'smartsizing:i:'
    SpanMonitors                          = 'span monitors:i:'
    SuperPanAccelerationFactor            = 'superpan acceleration factor:i:'
    UsbDevicesToRedirect                  = 'usbstoredirect:s:'
    UseMultimon                           = 'use multimon:i:'
    UseRedirectionServerName              = 'use redirection server name:i:'
    Username                              = 'username:s:'
    VideoPlaybackMode                     = 'videoplaybackmode:i:'
    WindowPosition                        = 'winposstr:s:'
    WorkspaceId                           = 'workspaceid:s:'
}

$ConnectionDefaults = @{
    AllowDesktopComposition  = $false
    AllowFontSmoothing       = $false
    AlternateShell           = ''
    AudioCapture             = $true
    AudioMode                = [AudioModeEnum]::PlayOnLocalComputer
    AutoReconnectionEnabled  = $true
    BandwidthAutoDetect      = $true
    BitmapCachePersistEnable = $true
    Compression              = $true
    ConnectionType           = [ConnectionTypeEnum]::AutoDetect
    DesktopHeight            = 600
    DesktopWidth             = 800
    DisableCursorSetting     = $false
    DisableFullWindowDrag    = $true
    DisableMenuAnimations    = $true
    DisableThemes            = $false
    DisableWallpaper         = $false
    DisplayConnectionBar     = $true
    EnableRDSAADAuth         = $false
    EnableWorkspaceReconnect = $false
    GatewayBrokeringType     = [GatewayBrokeringType]::Undefined
    KeyboardHook             = [KeyboardHookEnum]::RemoteFullScreenModeOnly
    KDCProxyName             = ''
    NegotiateSecurityLayer   = $true
    NetworkAutoDetect        = $true
    PromptCredentialOnce     = $false
    PromptForCredentials     = $true
    RDGIsKDCProxy            = $false
    RedirectClipboard        = $true
    RedirectCOMPorts         = $false
    RedirectLocation         = $false
    RedirectPOSDevices       = $false
    RedirectPrinters         = $true
    RedirectSmartCards       = $true
    RedirectWebAuthn         = $true
    RemoteApplicationMode    = $false
    RemoteAppMouseMoveInject = $true
    ScreenModeId             = [ScreenModeIdEnum]::FullScreen
    SessionBpp               = [SessionBppEnum]::TrueColor32Bit
    ShellWorkingDirectory    = ''
    UseMultimon              = $false
    UseRedirectionServerName = $false
    VideoPlaybackMode        = [VideoPlaybackModeEnum]::Optimized
    WindowPosition           = '0,3,0,0,800,600'
}

function Connect-RemoteDesktop {
    
    [CmdletBinding(
        DefaultParameterSetName = 'ComputerName'
    )]
    [Alias( 'rdp' )]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        'PSAvoidUsingPlainTextForPassword',
        'GatewayCredentialsSource',
        Justification = 'This is not a credential parameter.'
    )]
    param(
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            Mandatory        = $true,
            HelpMessage      = "Specifies the name or IP address (and optional port) of the remote computer that you want to connect to."
        )]
        [Alias('V', 'FullAddress')]
        [string]
        $ComputerName,

        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Defines an alternate default port for the Remote Desktop connection."
        )]
        [Nullable[UInt32]]
        $Port,

        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Connect to the administrative session of the remote computer."
        )]
        [switch]
        $AdministrativeSession,

        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether desktop composition (needed for Aero) is permitted when you log on to the remote computer."
        )]
        [switch]
        $AllowDesktopComposition,

        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether font smoothing may be used in the remote session."
        )]
        [switch]
        $AllowFontSmoothing,

        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Specifies an alternate name or IP address of the remote computer that you want to connect to."
        )]
        [string]
        $AlternateFullAddress,

        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Specifies a program to be started automatically when you connect to a remote computer. The value should be a valid path to an executable file."
        )]
        [string]
        $AlternateShell,

        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines how sounds captured (recorded) on the local computer are handled when you are connected to the remote computer."
        )]
        [switch]
        $AudioCapture,

        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines how sounds on a remote computer are handled when you are connected to the remote computer."
        )]
        [Alias('Sound')]
        [Nullable[AudioModeEnum]]
        $AudioMode,

        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines the quality of the audio played in the remote session."
        )]
        [Nullable[AudioQualityModeEnum]]
        $AudioQualityMode,

        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines what should happen when server authentication fails."
        )]
        [Nullable[AuthenticationLevelEnum]]
        $AuthenticationLevel,

        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines the maximum number of times the client computer will try to reconnect to the remote computer if the connection is dropped."
        )]
        [Nullable[UInt32]]
        $AutoReconnectMaxRetries,

        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether the client computer will automatically try to reconnect to the remote computer if the connection is dropped."
        )]
        [switch]
        $AutoReconnectionEnabled,

        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Enables the option for automatic detection of the network type. Used in conjunction with <b>networkautodetect</b>. Also see <b>connection type</b>."
        )]
        [switch]
        $BandwidthAutoDetect,

        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether bitmaps are cached on the local computer (disk-based cache). Bitmap caching can improve the performance of your remote session."
        )]
        [switch]
        $BitmapCachePersistEnable,

        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Specifies the size in kilobytes of the memory-based bitmap cache. The maximum value is 32000."
        )]
        [Nullable[UInt32]]
        $BitmapCacheSize,

        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Configures which cameras to redirect. This setting uses a list of KSCATEGORY_VIDEO_CAMERA interfaces of cameras enabled for redirection."
        )]
        [string[]]
        $CamerasToRedirect,

        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether the connection should use bulk compression."
        )]
        [switch]
        $Compression,

        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Connect to the console session of the remote computer."
        )]
        [switch]
        $ConnectToConsole,

        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Specifies pre-defined performance settings for the Remote Desktop session."
        )]
        [Nullable[ConnectionTypeEnum]]
        $ConnectionType,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "The height (in pixels) of the remote session desktop."
        )]
        [Alias( 'Height' )]
        [Nullable[UInt32]]
        $DesktopHeight,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "The width (in pixels) of the remote session desktop."
        )]
        [Alias( 'Width' )]
        [Nullable[UInt32]]
        $DesktopWidth,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines which supported Plug and Play devices on the client computer will be redirected and available in the remote session."
        )]
        [string[]]
        $DevicesToRedirect,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether you have to press CTRL+ALT+DELETE before entering credentials after you are connected to the remote computer."
        )]
        [switch]
        $DisableCtrlAltDel,

        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether the remote cursor setting is used in the remote session."
        )]
        [switch]
        $DisableCursorSetting,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether window content is displayed when you drag the window to a new location."
        )]
        [switch]
        $DisableFullWindowDrag,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether menus and windows can be displayed with animation effects in the remote session."
        )]
        [switch]
        $DisableMenuAnimations,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether themes are permitted when you log on to the remote computer."
        )]
        [switch]
        $DisableThemes,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether the desktop background is displayed in the remote session."
        )]
        [Alias( 'NoWallpaper' )]
        [switch]
        $DisableWallpaper,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether a new Terminal Server session is started with every launch of a RemoteApp to the same computer and with the same credentials."
        )]
        [switch]
        $DisableConnectionSharing,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Specifies whether the Remote Desktop client should check the remote computer for RemoteApp capabilities."
        )]
        [switch]
        $DisableRemoteAppCapabilitiesCheck,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether the connection bar appears when you are in full screen mode."
        )]
        [switch]
        $DisplayConnectionBar,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Specifies the name of the domain of the user."
        )]
        [string]
        $Domain,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines which local disk drives on the client computer will be redirected and available in the remote session."
        )]
        [string[]]
        $DrivesToRedirect,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether Remote Desktop will use CredSSP for authentication if it's available."
        )]
        [switch]
        $EnableCredSSPSupport,

        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether RDS Azure AD authentication is enabled."
        )]
        [switch]
        $EnableRDSAADAuth,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether SuperPan is enabled or disabled. SuperPan allows the user to navigate a remote desktop in full-screen mode without scroll bars, when the dimensions of the remote desktop are larger than the dimensions of the current client window. The user can point to the window border, and the desktop view will scroll automatically in that direction."
        )]
        [switch]
        $EnableSuperPan,

        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether reconnection to RemoteApp and Desktop Connections workspaces is enabled."
        )]
        [switch]
        $EnableWorkspaceReconnect,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Enables or disables encoding of redirected video."
        )]
        [switch]
        $EncodeRedirectedVideoCapture,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Specifies the credentials that should be used to validate the connection with the RD Gateway."
        )]
        [Nullable[GatewayCredentialsSourceEnum]]
        $GatewayCredentialsSource,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Specifies the hostname of the RD Gateway."
        )]
        [string]
        $Gateway,

        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Specifies the brokering type of the RD Gateway."
        )]
        [Nullable[GatewayBrokeringType]]
        $GatewayBrokeringType,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines the RD Gateway authentication method to be used."
        )]
        [Nullable[GatewayProfileUsageMethodEnum]]
        $GatewayProfileUsageMethod,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Specifies if and how to use a Remote Desktop Gateway (RD Gateway) server."
        )]
        [Nullable[GatewayUsageMethodEnum]]
        $GatewayUsageMethod,

        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Specifies the KDC proxy name when the RD Gateway is a KDC proxy."
        )]
        [string]
        $KDCProxyName,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines how Windows key combinations are applied when you are connected to a remote computer."
        )]
        [Nullable[KeyboardHookEnum]]
        $KeyboardHook,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether the level of security is negotiated or not."
        )]
        [switch]
        $NegotiateSecurityLayer,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether to use auomatic network bandwidth detection or not. Requires the option <b>bandwidthautodetect</b> to be set and correlates with <b>connection type</b> 7."
        )]
        [switch]
        $NetworkAutoDetect,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether or not the connection bar should be pinned to the top of the remote session upon connection when in full screen mode."
        )]
        [switch]
        $PinConnectionBar,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether Remote Desktop Connection will prompt for credentials when connecting to a remote computer for which the credentials have been previously saved."
        )]
        [switch]
        $PromptForCredentials,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether Remote Desktop Connection will prompt for credentials when connecting to a server that does not support server authentication."
        )]
        [switch]
        $PromptForCredentialsOnClient,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "When connecting through an RD Gateway, determines whether RDC should use the same credentials for both the RD Gateway and the remote computer."
        )]
        [switch]
        $PromptCredentialOnce,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether Remote Desktop Connection will be started in public mode."
        )]
        [switch]
        $PublicMode,

        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether the RD Gateway is a KDC proxy."
        )]
        [switch]
        $RDGIsKDCProxy,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Controls the quality of encoded video."
        )]
        [Nullable[RedirectedVideoCaptureEncodingQualityEnum]]
        $RedirectedVideoCaptureEncodingQuality,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether the clipboard on the client computer will be redirected and available in the remote session and vice versa."
        )]
        [switch]
        $RedirectClipboard,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether the COM (serial) ports on the client computer will be redirected and available in the remote session."
        )]
        [switch]
        $RedirectCOMPorts,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether DirectX will be enabled for the remote session."
        )]
        [switch]
        $RedirectDirectX,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether local disk drives on the client computer will be redirected and available in the remote session."
        )]
        [switch]
        $RedirectDrives,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether the location of the local device will be redirected and available in the remote session."
        )]
        [switch]
        $RedirectLocation,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether Microsoft Point of Service (POS) for .NET devices connected to the client computer will be redirected and available in the remote session."
        )]
        [switch]
        $RedirectPOSDevices,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether printers configured on the client computer will be redirected and available in the remote session."
        )]
        [switch]
        $RedirectPrinters,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether smart card devices on the client computer will be redirected and available in the remote session."
        )]
        [switch]
        $RedirectSmartCards,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether WebAuthn requests on the remote computer will be redirected to the local computer allowing the use of local authenticators (such as Windows Hello for Business and security key)."
        )]
        [switch]
        $RedirectWebAuthn,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Optional command line parameters for the RemoteApp."
        )]
        [string]
        $RemoteApplicationCmdline,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Specifies a file to be opened on the remote computer by the RemoteApp."
        )]
        [string]
        $RemoteApplicationFile,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether environment variables contained in the RemoteApp command line parameter should be expanded locally or remotely."
        )]
        [switch]
        $RemoteApplicationExpandCmdline,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether environment variables contained in the RemoteApp working directory parameter should be expanded locally or remotely."
        )]
        [switch]
        $RemoteApplicationExpandWorkingDir,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Specifies the file name of an icon file to be displayed in the Remote Desktop interface while starting the RemoteApp. By default RDC will show the standard Remote Desktop icon."
        )]
        [string]
        $RemoteApplicationIcon,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether a RemoteApp shoud be launched when connecting to the remote computer."
        )]
        [switch]
        $RemoteApplicationMode,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Specifies the name of the RemoteApp in the Remote Desktop interface while starting the RemoteApp."
        )]
        [string]
        $RemoteApplicationName,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Specifies the alias or executable name of the RemoteApp."
        )]
        [string]
        $RemoteApplicationProgram,

        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether mouse movements are injected directly into the RemoteApp session."
        )]
        [switch]
        $RemoteAppMouseMoveInject,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether the remote session window appears full screen when you connect to the remote computer."
        )]
        [Nullable[ScreenModeIdEnum]]
        $ScreenModeId,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Specifies which local displays to use for the remote session. The selected displays must be contiguous. Requires <b>use multimon</b> to be set to 1."
        )]
        [string[]]
        $SelectedMonitors,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines the color depth (in bits) on the remote computer when you connect."
        )]
        [Nullable[SessionBppEnum]]
        $SessionBpp,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "The working directory on the remote computer to be used if an alternate shell is specified."
        )]
        [string]
        $ShellWorkingDirectory,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "The encoded signature when using .rdp file signing."
        )]
        [string]
        $Signature,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Comma-delimited list of .rdp file settings for which the signature is generated when using .rdp file signing."
        )]
        [string]
        $SignScope,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether the client computer should scale the content on the remote computer to fit the window size of the client computer when the window is resized."
        )]
        [switch]
        $SmartSizing,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether the remote session window will be spanned across multiple monitors when you connect to the remote computer."
        )]
        [Alias( 'Span' )]
        [switch]
        $SpanMonitors,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Specifies the number of pixels that the screen view scrolls in a given direction for every pixel of mouse movement by the client when in SuperPan mode"
        )]
        [UInt32]
        $SuperPanAccelerationFactor,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines which supported RemoteFX USB devices on the client computer will be redirected and available in the remote session when you connect to a remote session that supports RemoteFX USB redirection."
        )]
        [string[]]
        $UsbDevicesToRedirect,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether the session should use true multiple monitor support when connecting to the remote computer."
        )]
        [Alias( 'Multimon' )]
        [switch]
        $UseMultimon,

        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether the remote computer's name is used for redirection when reconnecting to a disconnected session."
        )]
        [switch]
        $UseRedirectionServerName,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Specifies the name of the user account that will be used to log on to the remote computer."
        )]
        [string]
        $Username,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Determines whether RDC will use RDP efficient multimedia streaming for video playback."
        )]
        [Nullable[VideoPlaybackModeEnum]]
        $VideoPlaybackMode,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "Specifies the position and dimensions of the session window on the client computer."
        )]
        [string]
        $WindowPosition,
        
        [Parameter(
            ParameterSetName = 'ComputerName',
            HelpMessage      = "This setting defines the RemoteApp and Desktop ID associated with the RDP file that contains this setting."
        )]
        [string]
        $WorkspaceId,

        [Parameter(
            Mandatory        = $true,
            ParameterSetName = 'ConnectionFile',
            HelpMessage      = "Specifies the path to a .rdp file that contains the connection settings for the remote computer."
        )]
        [Alias( 'Path' )]
        [System.IO.FileInfo]
        $ConnectionFile,

        [Parameter(
            HelpMessage = 'Automatically trust hosts and do not prompt to connect.'
        )]
        [switch]
        $AutoTrustHosts,
        
        [Parameter(
            HelpMessage      = 'Credential to use for the remote desktop connection.'
        )]
        [pscredential]
        $Credential,

        [Parameter(
            HelpMessage      = 'Specifies the RD Gateway credential to use for the remote desktop connection.'
        )]
        [Alias( 'GWCredential' )]
        [pscredential]
        $GatewayCredential

    )

    if ( $PSBoundParameters.ContainsKey('DisplayConnectionBar') -and -not $DisplayConnectionBar.IsPresent ) {
        Write-Warning ( 'The connection bar will be hidden by default. To show the connection bar press Ctrl + Alt + Home after the connection is established.' )
    }

    if ( $PSBoundParameters.ContainsKey('Credential') -and -not $PSBoundParameters.ContainsKey('GatewayCredential') -and $PSBoundParameters.ContainsKey('PromptForCredentialOnce') ) {
        $PSBoundParameters['PromptForCredentials'] = $false
        $PSBoundParameters['GatewayCredential']    = $Credential
        $PSBoundParameters.Remove('Credential') > $null
    }

    if ( $PSBoundParameters.ContainsKey('Credential') ) {
        $PSBoundParameters['PromptForCredentials']        = $false
    }

    if ( $PSBoundParameters.ContainsKey('GatewayCredential') ) {
        $PSBoundParameters['GatewayCredentialsSource']  = [GatewayCredentialsSourceEnum]::SelectLater
    }

    if ( $PSBoundParameters.ContainsKey('Gateway') ) {
        $PSBoundParameters['GatewayProfileUsageMethod'] = [GatewayProfileUsageMethodEnum]::UseExplicitSettings
    }

    $Script:ConnectionDefaults.Keys | ForEach-Object {
        if ( -not $PSBoundParameters.ContainsKey($_) ) {
            $PSBoundParameters[$_] = $Script:ConnectionDefaults[$_]
        }
    }

    [string[]] $ConnectionFileContent = $PSBoundParameters.Keys | ForEach-Object {

        if ( -not $Script:ParameterToSettingMap.ContainsKey($_) ) { return }

        if ( $PSBoundParameters[$_] -is [switch] ) {
            return $Script:ParameterToSettingMap[$_] + [int]$PSBoundParameters[$_].IsPresent
        }
        
        if ( $PSBoundParameters[$_] -is [bool] ) {
            return $Script:ParameterToSettingMap[$_] + [int]$PSBoundParameters[$_]
        }

        if ( $PSBoundParameters[$_] -is [enum] ) {
            return $Script:ParameterToSettingMap[$_] + $PSBoundParameters[$_].value__
        }

        if ( $PSBoundParameters[$_] -is [array] ) {
            return $Script:ParameterToSettingMap[$_] + ($PSBoundParameters[$_] -join ';')
        }

        return $Script:ParameterToSettingMap[$_] + $PSBoundParameters[$_]

    }

    $ConnectionFileContent | ForEach-Object { Write-Debug ( '[conection file] {0}' -f $_ ) }

    # create a temporary file to store the connection settings
    if ( $ConnectionFileContent ) {

        # get my documents path
        $DocumentsPath = [Environment]::GetFolderPath( [Environment+SpecialFolder]::MyDocuments )

        # use the Default.rdp
        $ConnectionFile = [System.IO.Path]::Combine($DocumentsPath, 'Default.rdp')
        $ConnectionFileContent -join "`r`n" | Set-Content -Path $ConnectionFile -Encoding UTF8 -Force

    }
    
    Write-Verbose ( 'Connection File: {0}' -f $ConnectionFile.FullName )

    # suppress connection warnings
    if ( $AutoTrustHosts -and $PSBoundParameters.ContainsKey('ComputerName') ) {
        if ( $PSBoundParameters.ContainsKey('Gateway') ) {
            New-ItemProperty -Path 'HKCU:\Software\Microsoft\Terminal Server Client\LocalDevices' -Name "${ComputerName};${Gateway}" -Type DWord -Value 0xffffffff -Force > $null
        } else {
            if ( $Certificate = Get-HostCertificate @PSBoundParameters ) {
                New-ItemProperty -Path "HKCU:\Software\Microsoft\Terminal Server Client\Servers\$ComputerName" -Name CertHash -Type Binary -Value $Certificate.GetCertHash('SHA256') -Force > $null
            }
        }
    }

    if ( $PSBoundParameters.ContainsKey('Credential') ) {
        Set-CmdkeyCredential -RemoteDesktop -ComputerName $ComputerName -Credential $Credential -ErrorAction Stop
    }

    if ( $PSBoundParameters.ContainsKey('Gateway') ) {
        if ( $PSBoundParameters.ContainsKey('GatewayCredential') ) {
            Set-CmdkeyCredential -Gateway -ComputerName $Gateway -Credential $GatewayCredential -ErrorAction Stop
        } elseif ( $PSBoundParameters.ContainsKey('Credential') ) {
            Set-CmdkeyCredential -Gateway -ComputerName $Gateway -Credential $Credential -ErrorAction Stop
        }
    }

    Write-Verbose 'Starting remote desktop connection...'

    [System.Collections.Generic.List[string]] $MstscArguments = @()

    if ( $PSCmdlet.ParameterSetName -eq 'ConnectionFile' ) {

        $MstscArguments.Add( $ConnectionFile.FullName )

        $WindowTitle = '{0} - {1} - Remote Desktop Connection' -f $ConnectionFile.BaseName, $ComputerName

    } else {

        if ( $PSBoundParameters.ContainsKey('Port') ) {
            $ComputerName += ":$Port"
        }

        $MstscArguments.Add( "/v:${ComputerName}" )
        if ( $PSBoundParameters.ContainsKey('Gateway') ) {
            $MstscArguments.Add( "/g:${Gateway}" )
        }

        $WindowTitle = '{0} - Remote Desktop Connection' -f $ComputerName

    }

    $MstscProcess = Start-Process -FilePath mstsc.exe -ArgumentList $MstscArguments -PassThru

    Write-Debug ( 'MSTSC Process Id: {0}' -f $MstscProcess.Id )

    try {
        Write-Verbose 'Waiting for connection window to close...'
        while ( $MstscProcess = Get-Process -Id $MstscProcess.Id -ErrorAction Stop ) {
            Write-Debug ( 'MainWindowTitle: {0}' -f $MstscProcess.MainWindowTitle )
            if ( $MstscProcess.MainWindowTitle -eq $WindowTitle ) { break }
            Start-Sleep -Milliseconds 100
        }
    }
    # window closed while waiting for connection
    catch {
        Write-Debug 'Connection window closed before establishing connection.'
        return
    }

    if ( $PSBoundParameters.ContainsKey('Credential') ) {
        Remove-CmdkeyCredential -RemoteDesktop -ComputerName $ComputerName -ErrorAction Stop
    }

    if ( $PSBoundParameters.ContainsKey('Gateway') ) {
        if ( $PSBoundParameters.ContainsKey('GatewayCredential') ) {
            Remove-CmdkeyCredential -Gateway -ComputerName $Gateway -ErrorAction Stop
        } elseif ( $PSBoundParameters.ContainsKey('Credential') ) {
            Remove-CmdkeyCredential -Gateway -ComputerName $Gateway -ErrorAction Stop
        }
    }

}


function Set-CmdkeyCredential {
    [CmdletBinding( DefaultParameterSetName = 'Remote Desktop' )]
    param(

        [Parameter( Mandatory, ParameterSetName = 'Remote Desktop' )]
        [switch]
        $RemoteDesktop,

        [Parameter( Mandatory, ParameterSetName = 'Gateway' )]
        [switch]
        $Gateway,

        [Parameter( Mandatory, ParameterSetName = 'Remote Desktop' )]
        [Parameter( Mandatory, ParameterSetName = 'Gateway' )]
        [ValidateScript({
            if ( $_ -notmatch '^[a-zA-Z0-9\.\-]+$' ) {
                throw 'Must be a valid hostname or IP address.'
            }
            return $true
        })]
        [string]
        $ComputerName,

        [Parameter( Mandatory, ParameterSetName = 'Remote Desktop' )]
        [Parameter( Mandatory, ParameterSetName = 'Gateway' )]
        [pscredential]
        $Credential

    )

    Write-Verbose ( 'Storing {0} credentials for host {1}...' -f $PSCmdlet.ParameterSetName.ToLower(), $ComputerName )

    if ( $PSCmdlet.ParameterSetName -eq 'Remote Desktop' ) {
        $CmdkeyParams = @(
            '/generic:TERMSRV/{0}' -f $ComputerName
            '/user:{0}'            -f $Credential.UserName
            '/pass:{0}'            -f $Credential.GetNetworkCredential().Password
        )
    } else {
        $CmdkeyParams = @(
            '/add:{0}'             -f $ComputerName
            '/user:{0}'            -f $Credential.UserName
            '/pass:{0}'            -f $Credential.GetNetworkCredential().Password
        )
    }

    $Result = cmdkey.exe @CmdkeyParams *>&1 |
        Out-String -Stream |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        ForEach-Object { $_ -replace '^CMDKEY:\s*' }

    if ( $LASTEXITCODE -ne 0 ) {
        $Result | Write-Information -InformationAction Continue
        throw ( 'Failed to store credentials for {0}' -f $ComputerName )
    } else {
        $Result | Write-Verbose
    }

}

function Remove-CmdkeyCredential {
    [CmdletBinding( DefaultParameterSetName = 'Remote Desktop' )]
    param(

        [Parameter( Mandatory, ParameterSetName = 'Remote Desktop' )]
        [switch]
        $RemoteDesktop,

        [Parameter( Mandatory, ParameterSetName = 'Gateway' )]
        [switch]
        $Gateway,

        [Parameter( Mandatory, ParameterSetName = 'Remote Desktop' )]
        [Parameter( Mandatory, ParameterSetName = 'Gateway' )]
        [ValidateScript({
            if ( $_ -notmatch '^[a-zA-Z0-9\.\-]+$' ) {
                throw 'Must be a valid hostname or IP address.'
            }
            return $true
        })]
        [string]
        $ComputerName

    )

    Write-Verbose ( 'Removing {0} credentials for host {1}...' -f $PSCmdlet.ParameterSetName.ToLower(), $ComputerName )

    if ( $PSCmdlet.ParameterSetName -eq 'Remote Desktop' ) {
        $CmdkeyParams = @(
            '/delete:TERMSRV/{0}' -f $ComputerName
        )
    } else {
        $CmdkeyParams = @(
            '/delete:{0}'          -f $ComputerName
        )
    }

    $Result = cmdkey.exe @CmdkeyParams *>&1 |
        Out-String -Stream |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        ForEach-Object { $_ -replace '^CMDKEY:\s*' }

    if ( $LASTEXITCODE -ne 0 ) {
        $Result | Write-Information -InformationAction Continue
        throw ( 'Failed to remove credentials for {0}' -f $ComputerName )
    } else {
        $Result | Write-Verbose
    }

}

function Get-HostCertificate {
    [CmdletBinding()]
    param(

        [Parameter( Mandatory )]
        [string]
        $ComputerName,

        [UInt32]
        $Port = 3389,

        [Parameter( DontShow, ValueFromRemainingArguments )]
        $IgnoredParameters

    )

    try {

        $TcpClient = New-Object System.Net.Sockets.TcpClient
        $TcpClient.Connect( $ComputerName, $Port )

        $SslStream = New-Object System.Net.Security.SslStream( $TcpClient.GetStream(), $false, { $true } )
        $SslStream.AuthenticateAsClient( $ComputerName )

        $Certificate = $SslStream.RemoteCertificate

    } catch {
        
        Write-Warning $_.Exception.Message
        $Certificate = $null

    } finally {
        
        $SslStream.Close()

        if ( $TcpClient.Connected ) {
            $TcpClient.Close()
        }

    }
    
    return $Certificate

}