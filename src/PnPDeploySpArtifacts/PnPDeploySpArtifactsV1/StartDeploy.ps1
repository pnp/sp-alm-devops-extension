[CmdletBinding()]
param()



# For more information on the VSTS Task SDK:
# https://github.com/Microsoft/vsts-task-lib

Trace-VstsEnteringInvocation $MyInvocation

try {
    Import-VstsLocStrings "$PSScriptRoot/task.json"
	
    . "$PSScriptRoot/ps_modules/CommonScripts/Utility.ps1"
    # get the tmp path of the agent
    $agentTmpPath = "$($env:AGENT_RELEASEDIRECTORY)\_temp"
    $tmpInlineXmlFileName = [System.IO.Path]::GetRandomFileName() + ".xml"

    [string]$SharePointVersion = Get-VstsInput -Name SharePointVersion
		
    [string]$FileOrInline = Get-VstsInput -Name FileOrInline

    [string]$PnPXmlFilePath = ""

    if ($FileOrInline -eq "File") {
        [string]$PnPXmlFilePath = Get-VstsInput -Name PnPXmlFilePath
        if (-not (Test-Path $PnPXmlFilePath)) {
            Write-VstsTaskError -Message "`nFile path '$PnPXmlFilePath' for variable `$PnPXmlFilePath does not exist.`n"
        }
    }
    else {

        #get xml string and check for valid xml
        [string]$PnPXmlInline = (Get-VstsInput -Name PnPXmlInline)
		
        $PnPXml = New-Object System.Xml.XmlDocument
        try {
            $PnPXmlFilePath = "$agentTmpPath/$tmpInlineXmlFileName"
            #if patrh not exists, create it!
            if (-not (Test-Path -Path $agentTmpPath)) {
                New-Item -ItemType Directory -Force -Path $agentTmpPath
            }
            $PnPXml.LoadXml($PnPXmlInline)
            $PnPXml.Save($PnPXmlFilePath)
        }
        catch [System.Xml.XmlException] {
            throw "$($_.toString())"		
        }
    }

    [string]$Handlers = (Get-VstsInput -Name Handlers)

    $TmpParameters = (Get-VstsInput -Name Parameters)

    $ConnectedService = Get-VstsInput -Name ConnectedServiceName -Require
    $ServiceEndpoint = (Get-VstsEndpoint -Name $ConnectedService -Require)

    [string]$WebUrl = $ServiceEndpoint.Url
    if (($WebUrl -match "(http[s]?|[s]?ftp[s]?)(:\/\/)([^\s,]+)") -eq $false) {
       #Write-VstsTaskError -Message "`nweb url '$WebUrl' of the variable `$WebUrl is not a valid url. E.g. http://my.sharepoint.sitecollection.`n"
    }

    [string]$DeployUserName = $ServiceEndpoint.Auth.parameters.username

    [string]$DeployPassword = $ServiceEndpoint.Auth.parameters.password

    [string]$RequiredVersion = Get-VstsInput -Name RequiredVersion

    [bool]$ClearNavigation = Get-VstsInput -Name ClearNavigation -AsBool

    [bool]$IgnoreDuplicateDataRowErrors = Get-VstsInput -Name IgnoreDuplicateDataRowErrors -AsBool

    [bool]$OverwriteSystemPropertyBagValues = Get-VstsInput -Name OverwriteSystemPropertyBagValues -AsBool

    [bool]$ProvisionContentTypesToSubWebs = Get-VstsInput -Name ProvisionContentTypesToSubWebs -AsBool

    #preparing pnp provisioning
    $agentToolsPath = Get-VstsTaskVariable -Name 'agent.toolsDirectory' -Require #"$($env:AGENT_WORKFOLDER)\_tool"
    $null = Load-PnPPackages -SharePointVersion $SharePointVersion -AgentToolPath $agentToolsPath -RequiredVersion $RequiredVersion

    $secpasswd = ConvertTo-SecureString $DeployPassword -AsPlainText -Force
    $adminCredentials = New-Object System.Management.Automation.PSCredential ($DeployUserName, $secpasswd)

    Write-Host "`nConnect to '$WebUrl' as '$DeployUserName'..."
    Connect-PnPOnline -Url $WebUrl -Credentials $adminCredentials
    Write-Host "Successfully connected to '$WebUrl'...`n" 


    $ProvParams = @{ 
        Path                             = $PnPXmlFilePath
        ClearNavigation                  = $ClearNavigation
        IgnoreDuplicateDataRowErrors     = $IgnoreDuplicateDataRowErrors
        OverwriteSystemPropertyBagValues = $OverwriteSystemPropertyBagValues
        ProvisionContentTypesToSubWebs   = $ProvisionContentTypesToSubWebs
    } 

    #check for handlers
    if (-not [string]::IsNullOrEmpty($Handlers)) {
        $ProvParams.Handlers = [System.String]::Join(",",$Handlers.split(",;"))
    }

    #check for parameters
    if (-not [string]::IsNullOrEmpty($TmpParameters)) {
        [System.Collections.Hashtable] $Parameters = ConvertFrom-StringData -StringData $TmpParameters
        $ProvParams.Parameters = $Parameters
    }

    #execute provisioning
    Apply-PnPProvisioningTemplate @ProvParams

}
catch {
    $ErrorMessage = $_.Exception.Message
    Write-VstsTaskError -Message "`nAn Error occured. The error message was: $ErrorMessage. `n Stackstace `n $($_.ScriptStackTrace)`n"
    Write-VstsSetResult -Result 'Failed' -Message "Error detected" -DoNotThrow
}
finally {
    Trace-VstsLeavingInvocation $MyInvocation

    #clean up tmp path
    if ($FileOrInline -eq 'Inline' -and (Test-Path $agentTmpPath)) {
        Remove-Item $agentTmpPath -Recurse       
    }
}
    
