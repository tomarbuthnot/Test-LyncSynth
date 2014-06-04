Function Test-LyncSynth {
<#
.SYNOPSIS

.DESCRIPTION

.PARAMETER Path 

.PARAMETER Name 

.PARAMETER Test 

.EXAMPLE
PS C:\> Set-MyScript

.NOTES
Version: 0.3
Author : Tom Arbuthnot
Disclaimer: Use completely at your own risk. Test before using on any production system.
Do not run any script you don't understand.

.INPUTS

.OUTPUTS

.LINK
#>
  
  # Sets that -Whatif and -Confirm should be allowed
  [cmdletbinding(SupportsShouldProcess=$true)]
  
  Param 	(
    [Parameter(Mandatory=$True,
    HelpMessage="Define as '2010' or '2013' ")]
    $LyncServerVersion = '2013',
    
    [Parameter(Mandatory=$true,
    HelpMessage='Enter Test Username1 in the form domain\user')]
    $Username1 = 'defaultvalue1',
    
    [Parameter(Mandatory=$true,
    HelpMessage='Enter Username1 Password')]
    $Password1 = 'defaultvalue1',
    
    [Parameter(Mandatory=$true,
    HelpMessage='Enter Test Username2 in the form domain\user')]
    $Username2 = 'defaultvalue1',
    
    [Parameter(Mandatory=$true,
    HelpMessage='Enter Username2 Password')]
    $Password2 = 'defaultvalue1',
    
    [Parameter(Mandatory=$true,
    HelpMessage='Lync Pool FQDN')]
    $LyncPoolFQDN = 'defaultvalue1',
    
    [Parameter(Mandatory=$true,
    HelpMessage='Lync EDGE Pool FQDN')]
    $EdgePoolFQDN = 'defaultvalue1',
    
    [Parameter(Mandatory=$false,
    HelpMessage='Distribution List for Testing DL Expansion')]
    $GroupExpansionEmailDLToTest = 'defaultvalue1',
    
    [Parameter(Mandatory=$false,
    HelpMessage='Test Phone Number For PSTN Test in foramt +44xxx')]
    $TestPhoneNumber = 'defaultvalue1',
    
    [Parameter(Mandatory=$false,
    HelpMessage='Error Log location, default C:\<Command Name>_ErrorLog.txt')]
    [string]$ErrorLog = "c:\$($myinvocation.mycommand)_ErrorLog.txt",
    [switch]$LogErrors
    
  ) #Close Parameters
  
  Begin 	{
    
    Import-Module Lync
    
    # Create new event log source (error action silently continue if already exists
    # Could tidy this up by catching the specific exception, could also get access denied
    # New-EventLog -LogName Application -Source "Test-LyncServiceHealth" -ErrorAction SilentlyContinue
    
    
    # $VerbosePreference = "Continue"
    
    #$user1 = Read-host "Please enter test user1 in the form domain\user and press enter"
    #$user2 = Read-host "Please enter test user2 in the form domain\user and press enter"
    
    
    $secstr1 = New-Object -TypeName System.Security.SecureString
    $password1.ToCharArray() | ForEach-Object {$secstr1.AppendChar($_)}
    $cred1 = new-object -typename System.Management.Automation.PSCredential -argumentlist $username1, $secstr1
    
    
    $secstr2 = New-Object -TypeName System.Security.SecureString
    $password2.ToCharArray() | ForEach-Object {$secstr2.AppendChar($_)}
    $cred2 = new-object -typename System.Management.Automation.PSCredential -argumentlist $username2, $secstr2
    
    
    # Note, the cmdlets involving users will test their respective pools. In an ideal world you would test users from every pool with every other pool to confirm cross communications. S
    # Similarly, you would test from a remote server to the pool to test the network connectivity, rather than running on an FE for that user (though in fact in a pool the user may be homed
    # to a different FE in that pool.
    
    # Also note in any voice tests the users dial plan will influence the gateway(s) used
    
    # Note this may not catch issues on a particular FE. It test "service" availability not every "server" in the pool working for that service.
    
    # $LyncPoolFQDN = "lonpool.tomuc.com"
    
    # $GroupExpansionEmailDLToTest = "DistGroup1"
    
    #$cred1 = Get-Credential $user1
    
    #$cred2 = Get-Credential $user2
    
    $getuser1 = get-csuser -identity $username1
    $getuser2 = get-csuser -identity $username2
    $SIPURIUser1 = $getuser1.sipaddress
    $SIPURIUser2 = $getuser2.sipaddress
    
    # $testphonenumber = Read-host "Please Enter a test phone number in E164 format i.e. +44123412343"
    
    Write-Verbose "Pool for $SIPURIUser1 is $($getuser1.registrarpool)"
    Write-Verbose "Pool for $SIPURIUser2 is $($getuser2.registrarpool)"
    
    Write-Verbose "Starting $($myinvocation.mycommand)"
    Write-Verbose "Error log will be $ErrorLog"
    
    # Set everytihng ok to true, this is used to stop the script if we have an issue
    # Each Try Catch Finally block, or action (within the process block of the function) depends on $EverythingOK being true
    # A dependancy step will set $everything_ok to $false, therefore other steps will be skipped
    $EverythingOK = $true
    
    # Catch Actions Function to avoid repeating code, don't need to . source within a script
    Function ErrorCatch-Actions 
    {
      Param 	(
        [Parameter(Mandatory=$false,
        HelpMessage='Switch to Allow Errors to be Caught without setting EverythingOK to False, stopping other aspects of the script running')]
        # By default any errors caught will set $EverythingOK to false causing other parts of the script to be skipped
        [switch]$SetEverythingOKVariabletoTrue
      ) # Close Parameters
      # Set Everything OK to false to avoid running dependant actions
      If($SetEverythingOKVariabletoTrue) {$EverythingOK = $true}
      else {$EverythingOK = $false}
      # Print Error to Output
      Write-Output ' '
      Write-Warning '%%% Error Catch Has Been Triggered (To log errors to text file start script with -LogErrors switch) %%%'
      Write-Output ' '
      Write-Warning 'Last Error was:'
      Write-Output ' '
      Write-Error $Error[0]
      if ($LogErrors) {
        # Add Date to Error Log File
        
        Get-Date -format 'dd/MM/yyyy HH:mm' | Out-File $ErrorLog -Append
        # Output Error to Error Log file
        $Error | Out-File $ErrorLog -Append
        '%%%%%%%%%%%%%%%%%%%%%%%%%% LINE BREAK BETWEEN ERRORS %%%%%%%%%%%%%%%%%%%%%%%%%%' | Out-File $ErrorLog -Append
        ' ' | Out-File $ErrorLog -Append
        Write-Warning "Errors Logged to $ErrorLog"
        # Clear Error Log Variable
        $Error.Clear()
      } #Close If
    } # Close Error-CatchActons Function
    
  } #Close Function Begin Block
  
  Process {
    
    
    $loop = 'True'
    # Create a loop (loop keeps going until variable in config file is set to false)
    While ($loop -like 'true')
    # Massive Repeating loop start (LOOP1), with delay between checking for every server
    
    {
      
      
      $OutputCollection =  @()
      
      Write-Verbose 'Beginning Tests....'
      
      
      #################################################################
      
      IF ($LyncServerVersion -eq '2010' -or $LyncServerVersion -eq '2013')
      {
        
        # 2010 and 2013 Tests
        
        $TestCsAddressBookService = Test-CsAddressBookService -targetfqdn $LyncPoolFQDN -usersipaddress $SIPURIUser1
        $Loopoutput = New-Object -TypeName PSobject
        $loopoutput | add-member NoteProperty 'Cmdlet' -value 'Test-CsAddressBookService'
        $loopoutput | add-member NoteProperty 'DateTime' -value "$(Get-Date)"
        $loopoutput | add-member NoteProperty 'Result' -value "$($TestCsAddressBookService.Result)"
        $loopoutput | add-member NoteProperty 'Details' -value "$($TestCsAddressBookService.Error)"
        $OutputCollection += $loopoutput
        
        $TestCsAddressBookWebQuery = Test-CsAddressBookWebQuery -targetfqdn $LyncPoolFQDN -usersipaddress $SIPURIUser1
        $Loopoutput = New-Object -TypeName PSobject
        $loopoutput | add-member NoteProperty 'Cmdlet' -value 'Test-CsAddressBookWebQuery'
        $loopoutput | add-member NoteProperty 'DateTime' -value "$(Get-Date)"
        $loopoutput | add-member NoteProperty 'Result' -value "$($TestCsAddressBookWebQuery.Result)"
        $loopoutput | add-member NoteProperty 'Details' -value "$($TestCsAddressBookWebQuery.Error)"
        $OutputCollection += $loopoutput
        
        $TestCsAVConference = Test-CsAVConference -targetfqdn $LyncPoolFQDN -sendersipaddress $SIPURIUser1 -sendercredential $cred1 -receiversipaddress $SIPURIUser2 -receivercredential $cred2
        $Loopoutput = New-Object -TypeName PSobject
        $loopoutput | add-member NoteProperty 'Cmdlet' -value 'Test-CsAVConference'
        $loopoutput | add-member NoteProperty 'DateTime' -value "$(Get-Date)"
        $loopoutput | add-member NoteProperty 'Result' -value "$($TestCsAVConference.Result)"
        $loopoutput | add-member NoteProperty 'Details' -value "$($TestCsAVConference.Error)"
        $OutputCollection += $loopoutput
        
        
        $TestCsDialinconferencing = Test-CsDialinconferencing -TargetFqdn $LyncPoolFQDN -UserSipAddress $SIPURIUser1 -UserCredential $cred1
        $Loopoutput = New-Object -TypeName PSobject
        $loopoutput | add-member NoteProperty 'Cmdlet' -value 'Test-CsDialinconferencing'
        $loopoutput | add-member NoteProperty 'DateTime' -value "$(Get-Date)"
        $loopoutput | add-member NoteProperty 'Result' -value "$($TestCsDialinconferencing.Result)"
        $loopoutput | add-member NoteProperty 'Details' -value "$($TestCsDialinconferencing.Error)"
        #$OutputCollection += $loopoutput
        
        
        
        IF ($GroupExpansionEmailDLToTest -ne $null)
        {
          $TestCsGroupExpansion = Test-CsGroupExpansion -TargetFqdn $LyncPoolFQDN -GroupEmailAddress $GroupExpansionEmailDLToTest -UserSipAddress $SIPURIUser1 -UserCredential $cred1
          $Loopoutput = New-Object -TypeName PSobject
          $loopoutput | add-member NoteProperty 'Cmdlet' -value 'Test-CsGroupExpansion'
          $loopoutput | add-member NoteProperty 'DateTime' -value "$(Get-Date)"
          $loopoutput | add-member NoteProperty 'Result' -value "$($TestCsGroupExpansion.Result)"
          $loopoutput | add-member NoteProperty 'Details' -value "$($TestCsGroupExpansion.Error)"
          $OutputCollection += $loopoutput
        }
        
        $TestCsGroupIM = Test-CsGroupIM -TargetFqdn $LyncPoolFQDN -SenderSipAddress $SIPURIUser1 -SenderCredential $cred1 -ReceiverSipAddress $SIPURIUser2 -ReceiverCredential $cred2
        $Loopoutput = New-Object -TypeName PSobject
        $loopoutput | add-member NoteProperty 'Cmdlet' -value 'Test-CsGroupIM'
        $loopoutput | add-member NoteProperty 'DateTime' -value "$(Get-Date)"
        $loopoutput | add-member NoteProperty 'Result' -value "$($TestCsGroupIM.Result)"
        $loopoutput | add-member NoteProperty 'Details' -value "$($TestCsGroupIM.Error)"
        $OutputCollection += $loopoutput
        
        
        $TestCsIM = Test-CsIM -TargetFqdn $LyncPoolFQDN -sendersipaddress $SIPURIUser1 -receiversipaddress $SIPURIUser2
        $Loopoutput = New-Object -TypeName PSobject
        $loopoutput | add-member NoteProperty 'Cmdlet' -value 'Test-CsIM'
        $loopoutput | add-member NoteProperty 'DateTime' -value "$(Get-Date)"
        $loopoutput | add-member NoteProperty 'Result' -value "$($TestCsIM.Result)"
        $loopoutput | add-member NoteProperty 'Details' -value "$($TestCsIM.Error)"
        $OutputCollection += $loopoutput
        
        # Fails: Inner Exception:The remote name could not be resolved: 'lonpoolwebext.tomuc.com'
        $TestCsMcxConference = Test-CsMcxConference -TargetFqdn $LyncPoolFQDN -OrganizerSipAddress $SIPURIUser1 -OrganizerCredential $cred1 -UserSipAddress $SIPURIUser1 -UserCredential $cred1 -User2SipAddress $SIPURIUser2 -User2Credential $cred2 -Authentication ClientCertificate
        $Loopoutput = New-Object -TypeName PSobject
        $loopoutput | add-member NoteProperty 'Cmdlet' -value 'Test-CsMcxConference'
        $loopoutput | add-member NoteProperty 'DateTime' -value "$(Get-Date)"
        $loopoutput | add-member NoteProperty 'Result' -value "$($TestCsMcxConference.Result)"
        $loopoutput | add-member NoteProperty 'Details' -value "$($TestCsMcxConference.Error)"
        #$OutputCollection += $loopoutput
        
        
        # Fails
        $TestCsMcxP2PIM = Test-CsMcxP2PIM -TargetFqdn $LyncPoolFQDN -SenderSipAddress $SIPURIUser1 -SenderCredential $cred1 -ReceiverSipAddress $SIPURIUser2 -ReceiverCredential $cred2 -Authentication ClientCertificate
        $Loopoutput = New-Object -TypeName PSobject
        $loopoutput | add-member NoteProperty 'Cmdlet' -value 'Test-CsMcxP2PIM'
        $loopoutput | add-member NoteProperty 'DateTime' -value "$(Get-Date)"
        $loopoutput | add-member NoteProperty 'Result' -value "$($TestCsMcxP2PIM.Result)"
        $loopoutput | add-member NoteProperty 'Details' -value "$($TestCsMcxP2PIM.Error)"
        # $OutputCollection += $loopoutput
        
        
        IF ($EdgePoolFQDN -ne $null)
        {
          $TestCsMcxPushNotification = Test-CsMcxPushNotification -AccessEdgeFqdn $EdgePoolFQDN
          $Loopoutput = New-Object -TypeName PSobject
          $loopoutput | add-member NoteProperty 'Cmdlet' -value 'Test-CsMcxPushNotification'
          $loopoutput | add-member NoteProperty 'DateTime' -value "$(Get-Date)"
          $loopoutput | add-member NoteProperty 'Result' -value "$($TestCsMcxPushNotification.Result)"
          $loopoutput | add-member NoteProperty 'Details' -value "$($TestCsMcxPushNotification.Error)"
          #$OutputCollection += $loopoutput
        }
        
        $TestCSP2PAV = Test-CsP2PAV -TargetFqdn $LyncPoolFQDN -sendersipaddress $SIPURIUser1 -receiversipaddress $SIPURIUser2
        $Loopoutput = New-Object -TypeName PSobject
        $loopoutput | add-member NoteProperty 'Cmdlet' -value 'Test-CsP2PAV'
        $loopoutput | add-member NoteProperty 'DateTime' -value "$(Get-Date)"
        $loopoutput | add-member NoteProperty 'Result' -value "$($TestCSP2PAV.Result)"
        $loopoutput | add-member NoteProperty 'Details' -value "$($TestCSP2PAV.Error)"
        $OutputCollection += $loopoutput
        
        $TestCsPresence = Test-CsPresence -TargetFqdn $LyncPoolFQDN -subscribersipaddress $SIPURIUser1 -publishersipaddress $SIPURIUser2
        $Loopoutput = New-Object -TypeName PSobject
        $loopoutput | add-member NoteProperty 'Cmdlet' -value 'Test-CsPresence'
        $loopoutput | add-member NoteProperty 'DateTime' -value "$(Get-Date)"
        $loopoutput | add-member NoteProperty 'Result' -value "$($TestCsPresence.Result)"
        $loopoutput | add-member NoteProperty 'Details' -value "$($TestCsPresence.Error)"
        $OutputCollection += $loopoutput
        
        
        IF ($TestPhoneNumber -ne $null)
        {
          # Depends on the user voice policy, sends a call to PSTN
          $TestCsPstnOutboundCall = Test-CsPstnOutboundCall -TargetFqdn $LyncPoolFQDN -UserSipAddress $SIPURIUser1 -TargetPstnPhoneNumber $testphonenumber
          $Loopoutput = New-Object -TypeName PSobject
          $loopoutput | add-member NoteProperty 'Cmdlet' -value 'Test-CsPstnOutboundCall'
          $loopoutput | add-member NoteProperty 'DateTime' -value "$(Get-Date)"
          $loopoutput | add-member NoteProperty 'Result' -value "$($TestCsPstnOutboundCall.Result)"
          $loopoutput | add-member NoteProperty 'Details' -value "$($TestCsPstnOutboundCall.Error)"
          $OutputCollection += $loopoutput
        }
        
        # fails no voice path
        $TestCsPstnPeerToPeerCall = Test-CsPstnPeerToPeerCall -TargetFqdn $LyncPoolFQDN -SenderSipAddress $SIPURIUser1 -SenderCredential $cred1 -ReceiverSipAddress $SIPURIUser2 -ReceiverCredential $cred2
        $Loopoutput = New-Object -TypeName PSobject
        $loopoutput | add-member NoteProperty 'Cmdlet' -value 'Test-CsPstnPeerToPeerCall'
        $loopoutput | add-member NoteProperty 'DateTime' -value "$(Get-Date)"
        $loopoutput | add-member NoteProperty 'Result' -value "$($TestCsPstnPeerToPeerCall.Result)"
        $loopoutput | add-member NoteProperty 'Details' -value "$($TestCsPstnPeerToPeerCall.Error)"
        $OutputCollection += $loopoutput
        
        
        $TestCsRegistration = Test-CsRegistration -TargetFqdn $LyncPoolFQDN -UserSipAddress $SIPURIUser1
        $Loopoutput = New-Object -TypeName PSobject
        $loopoutput | add-member NoteProperty 'Cmdlet' -value 'Test-CsRegistration'
        $loopoutput | add-member NoteProperty 'DateTime' -value "$(Get-Date)"
        $loopoutput | add-member NoteProperty 'Result' -value "$($TestCsRegistration.Result)"
        $loopoutput | add-member NoteProperty 'Details' -value "$($TestCsRegistration.Error)"
        $OutputCollection += $loopoutput
        
        
        #Fails
        $TestCsWebApp = Test-CsWebApp -TargetFqdn $LyncPoolFQDN -UserSipAddress $SIPURIUser1 -UserCredential $cred1 -User2SipAddress $SIPURIUser2 -User2Credential $cred2
        $Loopoutput = New-Object -TypeName PSobject
        $loopoutput | add-member NoteProperty 'Cmdlet' -value 'Test-CsWebApp'
        $loopoutput | add-member NoteProperty 'DateTime' -value "$(Get-Date)"
        $loopoutput | add-member NoteProperty 'Result' -value "$($TestCsWebApp.Result)"
        $loopoutput | add-member NoteProperty 'Details' -value "$($TestCsWebApp.Error)"
        # Fails
        $OutputCollection += $loopoutput
        
        
        # Seems to Crash PowerShell, so ignored for now
        # $TestCsWebAppAnonymous = Test-CsWebAppAnonymous -TargetFqdn $LyncPoolFQDN -UserSipAddress $SIPURIUser1 -UserCredential $cred1
      }
      
      
      ##################################################################
      
      IF ($LyncServerVersion -eq '2010')
      {
        # 2010 Only
        
        $TestCsClientAuth2010 =  Test-CsClientAuth -TargetFqdn $LyncPoolFQDN -UserSipAddress $SIPURIUser1 -UserCredential $cred1
        $Loopoutput = New-Object -TypeName PSobject
        $loopoutput | add-member NoteProperty 'Cmdlet' -value 'Test-CsClientAuth'
        $loopoutput | add-member NoteProperty 'DateTime' -value "$(Get-Date)"
        $loopoutput | add-member NoteProperty 'Result' -value "$($TestCsClientAuth2010.Result)"
        $loopoutput | add-member NoteProperty 'Details' -value "$($TestCsClientAuth2010.Error)"
        $OutputCollection += $loopoutput
      }
      
      ##################################################################
      
      IF ($LyncServerVersion -eq '2013')
      {
        
        # 2013 Only
        
        $TestCsClientAuthentication2013 =  Test-CsClientAuthentication -TargetFqdn $LyncPoolFQDN -UserSipAddress $SIPURIUser1 -UserCredential $cred1
        $Loopoutput = New-Object -TypeName PSobject
        $loopoutput | add-member NoteProperty 'Cmdlet' -value 'Test-CsClientAuthentication'
        $loopoutput | add-member NoteProperty 'DateTime' -value "$(Get-Date)"
        $loopoutput | add-member NoteProperty 'Result' -value "$($TestCsClientAuthentication2013.Result)"
        $loopoutput | add-member NoteProperty 'Details' -value "$($TestCsClientAuthentication2013.Error)"
        $OutputCollection += $loopoutput
        
        
        # Seems to Crash PowerShell
        # $TestsASConference = Test-CsASConference -TargetFqdn $LyncPoolFQDN -sendersipaddress $SIPURIUser1 -receiversipaddress $SIPURIUser2
        
        
        IF ($EdgePoolFQDN -ne $null)
        {
          $TestCsAVEdgeConnectivity =  Test-CsAVEdgeConnectivity -TargetFqdn $LyncPoolFQDN -UserSipAddress $SIPURIUser1 -UserCredential $cred1
          $Loopoutput = New-Object -TypeName PSobject
          $loopoutput | add-member NoteProperty 'Cmdlet' -value 'Test-CsAVEdgeConnectivity'
          $loopoutput | add-member NoteProperty 'DateTime' -value "$(Get-Date)"
          $loopoutput | add-member NoteProperty 'Result' -value "$($TestCsAVEdgeConnectivity.Result)"
          $loopoutput | add-member NoteProperty 'Details' -value "$($TestCsAVEdgeConnectivity.Error)"
          $OutputCollection += $loopoutput
        }
        
        
        $TestCsDataConference = Test-CsDataConference -TargetFqdn $LyncPoolFQDN -SenderSipAddress $SIPURIUser1 -SenderCredential $cred1 -ReceiverSipAddress $SIPURIUser2 -ReceiverCredential $cred2 -TestJoinLauncher
        $Loopoutput = New-Object -TypeName PSobject
        $loopoutput | add-member NoteProperty 'Cmdlet' -value 'Test-CsDataConference'
        $loopoutput | add-member NoteProperty 'DateTime' -value "$(Get-Date)"
        $loopoutput | add-member NoteProperty 'Result' -value "$($TestCsDataConference.Result)"
        $loopoutput | add-member NoteProperty 'Details' -value "$($TestCsDataConference.Error)"
        $OutputCollection += $loopoutput
        
        
        
        $TestCsExUMConnectivity = Test-CsExUMConnectivity -TargetFqdn $LyncPoolFQDN -UserSipAddress $SIPURIUser1 -UserCredential $cred1
        $Loopoutput = New-Object -TypeName PSobject
        $loopoutput | add-member NoteProperty 'Cmdlet' -value 'Test-CsExUMConnectivity'
        $loopoutput | add-member NoteProperty 'DateTime' -value "$(Get-Date)"
        $loopoutput | add-member NoteProperty 'Result' -value "$($TestCsDataConference.Result)"
        $loopoutput | add-member NoteProperty 'Details' -value "$($TestCsDataConference.Error)"
        $OutputCollection += $loopoutput
        
        
        
        $TestCsExUMVoiceMail = Test-CsExUMVoiceMail -TargetFqdn $LyncPoolFQDN -SenderSipAddress $SIPURIUser1 -SenderCredential $cred1 -ReceiverSipAddress $SIPURIUser2
        $Loopoutput = New-Object -TypeName PSobject
        $loopoutput | add-member NoteProperty 'Cmdlet' -value 'Test-CsExUMVoiceMail'
        $loopoutput | add-member NoteProperty 'DateTime' -value "$(Get-Date)"
        $loopoutput | add-member NoteProperty 'Result' -value "$($TestCsExUMVoiceMail.Result)"
        $loopoutput | add-member NoteProperty 'Details' -value "$($TestCsExUMVoiceMail.Error)"
        $OutputCollection += $loopoutput
        
        
        $TestCsUcwaConference = Test-CsUcwaConference -TargetFqdn $LyncPoolFQDN -OrganizerSipAddress $SIPURIUser1 -OrganizerCredential $cred1 -ParticipantSipAddress $SIPURIUser2 -ParticipantCredential $cred2
        $Loopoutput = New-Object -TypeName PSobject
        $loopoutput | add-member NoteProperty 'Cmdlet' -value 'Test-CsUcwaConference'
        $loopoutput | add-member NoteProperty 'DateTime' -value "$(Get-Date)"
        $loopoutput | add-member NoteProperty 'Result' -value "$($TestCsUcwaConference.Result)"
        $loopoutput | add-member NoteProperty 'Details' -value "$($TestCsUcwaConference.Error)"
        $OutputCollection += $loopoutput
        
        
        $TestCsWebScheduler = Test-CsWebScheduler -TargetFqdn $LyncPoolFQDN -UserSipAddress $SIPURIUser1 -UserCredential $cred1
        $Loopoutput = New-Object -TypeName PSobject
        $loopoutput | add-member NoteProperty 'Cmdlet' -value 'Test-CsWebScheduler'
        $loopoutput | add-member NoteProperty 'DateTime' -value "$(Get-Date)"
        $loopoutput | add-member NoteProperty 'Result' -value "$($TestCsWebScheduler.Result)"
        $loopoutput | add-member NoteProperty 'Details' -value "$($TestCsWebScheduler.Error)"
        $OutputCollection += $loopoutput
        
        
        # Fails
        $TestCsXmppIM = Test-CsXmppIM -TargetFqdn $LyncPoolFQDN -UserSipAddress $SIPURIUser1 -UserCredential $cred1 -Receiver $SIPURIUser2
        $Loopoutput = New-Object -TypeName PSobject
        $loopoutput | add-member NoteProperty 'Cmdlet' -value 'Test-CsXmppIM'
        $loopoutput | add-member NoteProperty 'DateTime' -value "$(Get-Date)"
        $loopoutput | add-member NoteProperty 'Result' -value "$($TestCsXmppIM.Result)"
        $loopoutput | add-member NoteProperty 'Details' -value "$($TestCsXmppIM.Error)"
        # Fails
        $OutputCollection += $loopoutput
      }
      
      
      ##################################################################
      
      # UCS Only
      
      # Test-CsExStorageConnectivity
      
      # Test-CsExStorageNotification
      
      # Test-CsUnifiedContactStore
      
      
      #################################################################
      
      # Pchat Only
      
      # Test-CsPersistentChatMessage
      
      #################################################################
      
      
      # 2010 and 2013 Results
      
      $OutputCollection | Sort-Object Result,DateTime | Format-Table -AutoSize -Wrap
      
      # Pause may be v3 only
      Pause
      
      ##################################
      
      # Manage Results
      
      # Null out Tests Failed on this run
      $CmdletsFailedOnThisRun = @()
      
      Foreach ($Test in $OutputCollection)
      {
        IF ($test.Result -ne 'Success')
        {
          $Loopoutput = New-Object -TypeName PSobject
          $loopoutput | add-member NoteProperty 'Cmdlet' -value "$($Test.Cmdlet)"
          $CmdletsFailedOnThisRun += $loopoutput
        }
      }
      
      
      Write-Verbose 'Failure Collection Object'
      
      $CmdletsFailedOnThisRun | Sort-Object Result,DateTime | Format-Table -AutoSize -Wrap
      
      # Pause may be v3 only
      Pause
      
      ########################################################################
      
      # If there are some fails, we need to work out if we need to raise an alert
      
      
      
      IF ($CmdletsFailedOnThisRun -ne $null)
      {
        # Some cmdlets have failed on this run, now we need to work out the alert scenario
        
        # If the Fails on the previous run match the fails on this run and it's been less than 24 hours don't alert
        # Check the failure output with compare-object. count gives the number of differences
        If ( $((Compare-Object $TestsFailedOnPreviousRun $CmdletsFailedOnThisRun).Count) -eq 0 )
        {
          Write-Verbose 'The failure state is the same as the last run'
          
        }
        # Note this will error out on first run as tests failed on previous run is null, this is ok, output still does not -eq 0
        If ( $((Compare-Object $TestsFailedOnPreviousRun $CmdletsFailedOnThisRun).Count) -ne 0 )
        {
          Write-Verbose 'The failure state is different to the last run, report failure state has changed'
          Write-EventLog –LogName Application –Source 'Test-LyncServiceHealth' –EntryType Information –EventID 9900 –Message “The failure state is different to the last run, report failure state has changed $OutputCollection”
        }
      }
      
      # If there were no failures on this run, but there were failures on the previous run
      IF ($CmdletsFailedOnThisRun -eq $null -and $TestsFailedOnPreviousRun -ne $null)
      {
        
        Write-Verbose 'Last run was in a failure state but this run was all Success'
        Write-Verbose 'Report Service restored to normail'
      }
      
      # Pause may be v3 only
      Pause
      
      # Set the tests failed on previous run ready for the next run
      
      Write-Verbose 'Tests Failed on Previous Run'
      $TestsFailedOnPreviousRun
      
      Write-Verbose 'Tests Failed on this run'
      $CmdletsFailedOnThisRun
      
      $TestsFailedOnPreviousRun = $CmdletsFailedOnThisRun
      ######################################################
      
      
      
    } # Close Massive Loop
    
    
    
    
  } #Close Function Process Block
  
  End 	{
    Write-Verbose "Ending $($myinvocation.mycommand)"
  } #Close Function End Block
  
  
} #Close Function