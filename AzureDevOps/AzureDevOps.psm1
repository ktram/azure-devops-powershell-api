function New-AzureDevOpsTokenFile {
    <#
        .SYNOPSIS
            Creates a file containing the token as an encrypted string.

        .DESCRIPTION
            The New-AzureDevOpsTokenFile function stores the Azure DevOps personal access token as an encrypted string in a
            file. The same user who created the file must use it and on the same computer the file was created.

        .PARAMETER  Path
            Path to file to be created.

        .EXAMPLE
            PS C:\> New-AzureDevOpsTokenFile -Path C:\temp\pat.txt
            
            Create the file at C:\temp\pat.txt containing the personal access token as an encrypted string.
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(Position=0, Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Path
    )
    Read-Host -Prompt 'Enter personal access token' -AsSecureString | ConvertFrom-SecureString | Set-Content -Path $Path
}
Export-ModuleMember -Function New-AzureDevOpsTokenFile

function Get-AzureDevOpsTokenCredential {
    <#
        .SYNOPSIS
            Gets a credential object based on the personal access token.

        .DESCRIPTION
            The Get-AzureDevOpsTokenCredential creates a credential object for the personal access token. The user is not used.
            If no parameters are specified, a dialog box will appear that prompts the user to enter a personal access token.

        .PARAMETER  Token
            Path to file containing the personal access token as an encrypted string.

        .PARAMETER  TokenFile
            Path to file containing the personal access token as an encrypted string.

        .EXAMPLE
            PS C:\> Get-AzureDevOpsTokenCredential
            
            Creates a credential after prompting the user to enter the personal access token.

        .EXAMPLE
            PS C:\> Get-AzureDevOpsTokenCredential -TokenFile C:\temp\pat.txt
            
            Creates a credential from the encrypted token stored in C:\temp\pat.txt

        .EXAMPLE
            PS C:\> Get-AzureDevOpsTokenCredential -Token C:\temp\pat.txt
            
            Creates a credential from the personal access token.
    #>
    [CmdletBinding(DefaultParameterSetName='Interactive')]
    Param
    (
        [Parameter(Position=0, Mandatory=$false, ParameterSetName='TokenFile')]
        [ValidateScript({Test-Path -Path $_ -PathType Leaf})]
        [String]
        $TokenFile,

        [Parameter(Position=0, Mandatory=$false, ParameterSetName='Token')]
        [ValidateNotNullOrEmpty()]
        [String]
        $Token
    )
    $dummyUser = 'AzureDevOps'
    if ($Token) {
        $secureToken = $Token | ConvertTo-SecureString -AsPlainText -Force
        New-Object System.Management.Automation.PSCredential($dummyUser, $secureToken)
    } elseif ($TokenFile) {
        New-Object System.Management.Automation.PSCredential($dummyUser, (Get-Content -Path $TokenFile | ConvertTo-SecureString))
    } else {
        $Host.UI.PromptForCredential('AzureDevOps personal access token',
            'Please enter personal access token',
            $dummyUser, $null)
    }
}
Export-ModuleMember -Function Get-AzureDevOpsTokenCredential

function Invoke-AzureDevOpsRestMethod {
    <#
        .SYNOPSIS
           Sends a request to the Azure DevOps Services RESTful web service.

        .DESCRIPTION
            The Invoke-AzureDevOpsRestMethod function sends a request to the Azure DevOps Services RESTful web service. 

        .PARAMETER  Uri
            URI to send the request.

        .PARAMETER  Method 
            Specifies the method used for the web request. The default is GET.

        .PARAMETER  Body 
            Specifies the body of the request.

            The Body parameter can be used to specify a list of query parameters or specify the content of the response.

            When the input is a GET request the body is added to the URI as query parameters. For a POST request, the body
            will be converted to JSON.

        .PARAMETER  Credential 
            Credential that contains the personal access token.

        .PARAMETER  Token 
            Personal access token.

        .EXAMPLE
            PS C:\> Invoke-AzureDevOpsRestMethod -Uri 'https://<organization>.visualstudio.com/_apis/distributedtask/pools/<id>/agents' -Body @{includeCapabilities='true'} -Credential $cred
            
            Sends a request to Azure DevOps to get the test agents and their capabilities.
    #>
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter(Position=0, Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [Uri]
        $Uri,

        [Parameter(Position=1, Mandatory=$true, ParameterSetName='Credential')]
        [ValidateNotNullOrEmpty()]
        [PSCredential]
        $Credential,

        [Parameter(Position=1, Mandatory=$true, ParameterSetName='Token')]
        [ValidateNotNullOrEmpty()]
        [String]
        $Token,

        [Parameter(Position=2)]
        [Microsoft.PowerShell.Commands.WebRequestMethod]
        $Method = 'GET',

        [Parameter(Position=3)]
        $Body = $null
    )
    $encodedToken = if ($Token) {
        [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(":$Token"))
    } else {
        [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(":$($Credential.GetNetworkCredential().Password)"))
    }

    $restParams = @{
        'Uri' = $Uri
        'Headers' = @{Authorization = "Basic $encodedToken"}
        'Method' = $Method
    }

    if ($Method -eq 'GET') {
        $restParams.Add('Body', $Body)
    }

    if ($Method -eq 'POST') {
        $restParams.Add('Body', $Body)
        $restParams.Add('ContentType', 'application/json')
    }
    
    if ($PSCmdlet.ShouldProcess($Uri, $Method)) {
        Invoke-RestMethod @restParams
    }
}
Export-ModuleMember -Function Invoke-AzureDevOpsRestMethod