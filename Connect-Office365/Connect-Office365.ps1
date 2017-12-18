<# 
    .SYNOPSIS 
    Office 365 connection modules

    .DESCRIPTION 
    Prompts a user to enter username and password to connect to office 365 tenant
        
    .EXAMPLE 
        #Connect to all Office 365 services:
            Connect-Office365
                write-host "Do Some Stuff with Exchange, Skype and Security & Compliance"
            Disconnect-Office365

    .EXAMPLE
    #Connect to individual services:
            Connect-SFBOnline
                write-host "Do Some Skype For Business stuff"
                Disconnect-SFBOnline
            Connect-ExOnline
                write-host "Do Some Exchange stuff"
            Disconnect-ExOnline
            Connect-SCCOnline
                write-host "Do Some Stuff with Security & Compliance"
            Disconnect-SCCOnline

#>



Function Get-O365Credentials {
    <# 
        .SYNOPSIS
            Requests user for Office 365 credentials

        .DESCRIPTION
            Requests user for Office 365 credentials is calles as a subfunction of other functions in this module, do not call manually
        .PARAMETER Force
            -Force [boolean]
                override the check for credentials currently being set, used incase credentials are invalid

        .EXAMPLE
            ...
            write-host "Password incorrect"
            Get-O365Credentials -force $true
    #>
     [CmdletBinding()]Param (
        [Parameter(Mandatory=$false)][boolean]$force
    )
    #This function will be called a few times,
    #Check to see if it has already been entered or the request if forced (eg. bad password check)

    if(($force -eq $true) -or ($Script:O365Credentials -eq $null)){

        Log-write -logpath $Script:LogPath -linevalue "`t`tEnter your Office 365 admin credentials"

        $Script:O365Credentials = Get-Credential -Message "Enter your Office 365 admin credentials"

    }#endif
} #end function


Function Connect-MSOL{
    <# 
        .SYNOPSIS 
            Checks the MSOL session is connected

        .DESCRIPTION 
            Checks the MSOL session is connected and connects it if not
            Also checks to ensure the "credentials" are currently stored in the 'script:' variable scope
        
        .EXAMPLE 
            if(Connect-MSOL){write-host "MSOL is connected"}
        
    #>
    begin{    

        $CurrentMSOLStatus = Get-MsolDomain -ErrorAction SilentlyContinue
    }
    process{
        if($CurrentMSOLStatus){
            #The MSOL may be connected but the credentials may be clear, check this is not the case
            #The other functions in this module use this function to confirm the credentials and MSOL exist/connect
            If ($Script:O365Credentials -eq $null){

                Get-O365Credentials
                $CurrentMSOLStatus = Get-MsolDomain -ErrorAction SilentlyContinue
            }


            #MSOL is already connected, no need for output clutter
            return $true #MSOL is connected

        } else {

            Log-write -logpath $Script:LogPath -linevalue "`tConnecting to Microsoft Online (MSOL)"
            
            try{
                
                # Users credentials may be invalid so try to connect 3 times, any more risks
                # locking the users account.
                $i=0 #Start a counter
                Do {
                    # Check the connect sequence has run less than 3 times
                    if ($i -ge 3){
                        Log-write -logpath $Script:LogPath -linevalue "`t`tThe 3rd login attempt has failed, aborting script to avoid account lockout"
                        throw #cause an error, go to the catch statement
                    }

                    # Attempt to connect to MSOL
                    #Create a local copy of the Credentials, some of the O365 connecting modules seem to empty the variable after connecting
                    $MSOLCredentials = $Script:O365Credentials

                    Connect-MsolService -Credential $MSOLCredentials -ErrorAction SilentlyContinue -ErrorVariable ProcessError
                    $CurrentMSOLStatus = Get-MsolDomain -ErrorAction SilentlyContinue #See if the MSOL is connected by listing the MSOLDomains

                    if($CurrentMSOLStatus){
                        #If the domain list isnt empty MSOL is connected
                        Log-write -logpath $Script:LogPath -linevalue "`t`tMSOL connected"
                        $i = 4
                    } else {
                        # If the domain list is empty MSOL is not connected
                        # Request the credentials again, with force set to true as the credentials variable
                        # is not currently empty, it's just didnt connect, then add 1 (++) to the counter
                        
                        Log-write -logpath $Script:LogPath -linevalue "`t`tCredentials invalid, please try again"
                        
                        Get-O365Credentials -force $true
                        
                        $i ++

                    }
                
                } Until ($CurrentMSOLStatus) #If the MSOL is connected the 'do' loop doesnt need to continue

                return $true #MSOL is connected
                
            }catch{
                Log-Error -LogPath $Script:LogPath -ErrorDesc "Failed to connect to MSOL,check 'Microsoft Online Service Sign-in Assistant for IT Professionals' is installed`n`t$_.Exception" -ExitGracefully $True
            }

        }#EndProcess
    }#EndIf
}#EndFunction



Function Connect-SFBOnline{
    <# 
        .SYNOPSIS 
            Starts a Skype for business sessions

        .DESCRIPTION 
            Checks user is running script as an admin, checks the WinRM service is running and then connects to the Skype for business service.
            Will check MSOL is connected and credentials are valid too
        
        .EXAMPLE 
            Connect-SFBOnline
        
    #>
    begin{
        Log-write -logpath $Script:LogPath -linevalue "`tConnecting to Skype For Business Online"

        #The Skype for business module only seems to work if you run script as admin
        If(([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){
            #Returns true if script is running in admin context
        }else{
            Log-Error -LogPath $Script:LogPath -ErrorDesc "Please run this script with administrator privileges" -ExitGracefully $True
        }
    }
    Process{
        try{
            if (get-module -list SkypeOnlineConnector){
                Import-Module SkypeOnlineConnector
                Log-write -logpath $Script:LogPath -linevalue "`t`tSkypeOnlineConnector module loaded"

            }else{
                log-Error -LogPath $Script:LogPath -ErrorDesc "SkypeOnlineConnector module not found`nCheck 'Skype for Business Online, Windows PowerShell Module' is installed" -ExitGracefully $True
            }

            #The Skype module will need the WinRM module to be running
            if ((get-service winrm).status -ne "Running"){
                 Log-write -logpath $Script:LogPath -linevalue "`t`t`tThe WinRM service is not running, attempting to start..."
                 try{
                     start-service winrm
                 }catch{    
                    Log-write -logpath $Script:LogPath -linevalue "`t`t`tUnable to start, trying as administrator..."
                    #Failed to start, assume the user isnt an admin or hasnt run script as admin
                    #Prompt to run the script as an admin user
                    try {
                        
                        Start-Process powershell -Verb runAs -ArgumentList "start-service winrm" -Wait
                        
                        #Check to see if the service is now running
                        if ((get-service winrm).status -ne "Running"){
                            
                            log-Error -LogPath $Script:LogPath -ErrorDesc "Unable to start the WinRM service, please start manually" -ExitGracefully $True
                        
                        } else {
                                               
                            Log-write -logpath $Script:LogPath -linevalue "`t`tWinRM service started"
                        }
                    } catch {
                        
                    }#EndCatch
                 }#EndCatch
            }#EndIf (WinRM)

            Log-write -logpath $Script:LogPath -linevalue "`t`tCreating Skype PS session"

            # MSOL is connected so we can assume this session has valid O365 credentials stored as $Script:O365Credentials
            
            # The New-CSOnlineSessions command empties the credentials after use, so store a temporary copy of the information    
            $SFBOCredentials = $Script:O365Credentials

            # The correct way to connect to SFBOnline is like this
            try{
                # Sometimes if the autodiscover for the lync service is incorrect is just doesnt work, this will error
                # If the session fails to connect, process the failure as an error and contnue, this should trigger the catch statement
                $script:sfboSession = New-CsOnlineSession -Credential $Script:SFBOCredentials -ErrorAction ProcessError #-Verbose
            
            }Catch{
                Log-write -logpath $Script:LogPath -linevalue "`t`t`tAutomatic Skype For Business endpoint discovery failed, trying to use manual 'AdminDomain'"
                
                # The $Script:AdminDomain defined in the config file (optional) allows us to bypass the autodiscover failure.


                If ($Script:AdminDomain -eq $null){ 
                    Log-Error -LogPath $Script:LogPath -ErrorDesc "Unable to connect using automatic endpoint discovery, add the 'Script:AdminDomain' variable config file and define it manually." -ExitGracefully $True
                }
                


                $script:sfboSession = New-CsOnlineSession -Credential $Script:SFBOCredentials -OverrideAdminDomain $Script:AdminDomain #-Verbose
                Log-write -logpath $Script:LogPath -linevalue "`t`t`tSkype PS session built, connecting..."
             }
              
            # Create the session                                      
            Import-PSSession $script:sfboSession -DisableNameChecking -AllowClobber -Verbose | out-null
            
            Log-write -logpath $Script:LogPath -linevalue "`t`t`t`tConnected"

        }catch{

            Log-Error -LogPath $Script:LogPath -ErrorDesc "Failed to connect to Skype For Business Online`n$_.Exception" -ExitGracefully $True

        }#EndTry
    }#EndProcess
}#EndFunction



Function Disconnect-SFBOnline{
    <# 
        .SYNOPSIS 
            Disconnects a Skype for business sessions

        .DESCRIPTION 
            Checks for and disconnects a Skype for business session
                   
        .EXAMPLE 
            disconnect-SFBOnline
        
    #>
    try{
        # Logic to confirm if the Exchange online session has been disconnected.
        If ($script:sfboSession -eq $null) {
            Log-write -logpath $Script:LogPath -linevalue "No Skype For Business Online session found"
        }
        Else {
            Remove-PSSession $script:sfboSession
            If ($script:sfboSession.State -eq "Closed") {
                Log-write -logpath $Script:LogPath -linevalue "Skype For Business Online session closed"
            }
            ElseIf ($script:sfboSession.state -eq "Open") {
                Log-Error -LogPath $Script:LogPath -ErrorDesc "Skype For Business Online session did not close" -ExitGracefully $false
            }
        }
    }catch{
        Log-Error -LogPath $Script:LogPath -ErrorDesc "Failed to close the Skype For Business Online session" -ExitGracefully $True
    }#EndTry
}#EndFunction






Function Connect-ExOnline{
    <# 
        .SYNOPSIS 
            Starts a Exchange online sessions

        .DESCRIPTION 
            Checks user is running script as an admin, checks the WinRM service is running and then connects to the Exchange online service.
            Will check MSOL is connected and credentials are valid too
        
        .EXAMPLE 
            Connect-ExOnline
        
    #>

    begin{
    
        Log-write -logpath $Script:LogPath -linevalue "`tConnecting to Exchange Online"

    }
    Process{

        try{


            Log-write -logpath $Script:LogPath -linevalue "`t`tCreating Exchange Online PS session"
             
            # MSOL is connected so we can assume this session has valid O365 credentials stored as $Script:O365Credentials
            # Attempt to connect
            
            # Store a temporary copy of the 365 Creds
            $EXOCredentials = $Script:O365Credentials


            $script:exoSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $EXOCredentials -Authentication Basic -AllowRedirection #-verbose
            Log-write -logpath $Script:LogPath -linevalue "`t`t`tExchange Online PS session built, connecting..."
            
            Import-PSSession $script:exoSession -DisableNameChecking -AllowClobber | Out-Null
            

            Log-write -logpath $Script:LogPath -linevalue "`t`t`t`tConnected"
 
        }catch{

            Log-Error -LogPath $Script:LogPath -ErrorDesc "Failed to connect to Exchange Online`n$_.Exception" -ExitGracefully $True
        
        }#EndTry
    }#EndProcess
}#EndFunction



Function Disconnect-ExOnline{
    <# 
        .SYNOPSIS 
            Disconnects a Exchange sessions

        .DESCRIPTION 
            Checks for and disconnects a Exchange session
                   
        .EXAMPLE 
            disconnect-ExOnline
        
    #>
    try{
        # Logic to confirm if the Exchange online session has been disconnected.
        If ($script:exoSession -eq $null) {
            Log-write -logpath $Script:LogPath -linevalue "No Exchange session found"
        }
        Else {
            Remove-PSSession $script:exoSession
            If ($script:exoSession.State -eq "Closed") {
                Log-write -logpath $Script:LogPath -linevalue "Exchange Online session closed"
            }
            ElseIf ($script:exoSession.state -eq "Open") {
                Log-Error -LogPath $Script:LogPath -ErrorDesc "Exchange Online session did not close" -ExitGracefully $false
            }
        }
    }catch{
        Log-Error -LogPath $Script:LogPath -ErrorDesc "Failed to close the Exchange Online session" -ExitGracefully $True
    }#EndTry
}#EndFunction




Function Connect-SCCOnline{
    <# 
        .SYNOPSIS 
            Starts a Security & Compliance Center session

        .DESCRIPTION 
            Connects to the Security & Compliance Center service.
            Will check MSOL is connected and credentials are valid too
        
        .EXAMPLE 
            Connect-SCCOnline
        
    #>
    begin{
        Log-write -logpath $Script:LogPath -linevalue "`tConnecting to Security & Compliance Center Online"
 
    }
    Process{
        try{
            Log-write -logpath $Script:LogPath -linevalue "`t`tCreating Security and Compliance Center PS Session"

            # Store a temporary copy of the 365 Creds
            $SCCOCredentials = $Script:O365Credentials
            
            #MSOL is connected so we can assume this session has valid O365 credentials stored as $Script:O365Credentials
            #Attempt to connect
            $script:sccoSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $SCCOCredentials -Authentication "Basic" -AllowRedirection

            Log-write -logpath $Script:LogPath -linevalue "`t`t`tSecurity and Compliance Center PS session built, connecting..."

            Import-PSSession $script:sccoSession -Prefix cc -DisableNameChecking -AllowClobber | Out-Null

            Log-write -logpath $Script:LogPath -linevalue "`t`t`t`tConnected"
 
        }catch{
            Log-Error -LogPath $Script:LogPath -ErrorDesc "Failed to connect to Security & Compliance Center Online`n$_.Exception" -ExitGracefully $True
        }#EndTry
    }#EndProcess
}#EndFunction



Function Disconnect-SCCOnline{
    <# 
        .SYNOPSIS 
            Disconnects a Security & Compliance Center sessions

        .DESCRIPTION 
            Checks for and disconnects a Security & Compliance Center session
                   
        .EXAMPLE 
            disconnect-SCCOnline
        
    #>
    try{
        # Logic to confirm if the Security & Compliance Center online session has been disconnected.
        If ($script:sccoSession -eq $null) {
            Log-write -logpath $Script:LogPath -linevalue "No Security & Compliance Center Online session found"
        }
        Else {
            Remove-PSSession $script:sccoSession
            If ($script:sccoSession.State -eq "Closed") {
                Log-write -logpath $Script:LogPath -linevalue "Security & Compliance Center Online session closed"
            }
            ElseIf ($script:sccoSession.state -eq "Open") {
                Log-Error -LogPath $Script:LogPath -ErrorDesc "Security & Compliance Center Online session did not close" -ExitGracefully $false
            }
        }
    }catch{
        Log-Error -LogPath $Script:LogPath -ErrorDesc "Failed to close the Security & Compliance Center Online session" -ExitGracefully $True
    }#EndTry
}#EndFunction




Function Connect-Office365{
    <# 
        .SYNOPSIS 
            Starts a session for Skype for business, Exchange online and the Security & Compliance Center sessions

        .DESCRIPTION 
            Checks user is running script as an admin, checks the WinRM service is running and then connects to the Skype for business service.
            Will check MSOL is connected and credentials are valid too
        
        .EXAMPLE 
            Connect-SFBOnline
        
    #>
    try{

        if(Connect-MSOL){
            #If MSOL is connected or connects, nothing to do
        }else{
            #If MSOL fails to connect, the Connect-MSOL module will show the nessessary output
            #Throw an error and abort the script
            Log-Error -LogPath $Script:LogPath -ErrorDesc "$_.Exception" -ExitGracefully $True
        }  


        Connect-ExOnline

        Connect-SCCOnline
        Connect-SFBOnline
               
    }catch{
        Log-Error -LogPath $Script:LogPath -ErrorDesc "Failed to Connect to Office 365, please review logs" -ExitGracefully $True
    }#EndTry
}#EndFunction


Function Disconnect-Office365{
    <# 
        .SYNOPSIS 
            Disconnects a Skype for business sessions

        .DESCRIPTION 
            Checks for and disconnects a Skype for business session
                   
        .EXAMPLE 
            disconnect-SFBOnline
        
    #>
    try{
 

        Disconnect-SFBOnline
        

        Disconnect-ExOnline
             

        Disconnect-SCCOnline

        $Script:O365Credentials = $null

                
    }catch{
        Log-Error -LogPath $Script:LogPath -ErrorDesc "Failed to disconnect to Office 365, please review logs" -ExitGracefully $True
    }#EndTry
}#EndFunction

