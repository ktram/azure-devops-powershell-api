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
            Personal access token.

        .PARAMETER  TokenFile
            Path to file containing the personal access token as an encrypted string.

        .EXAMPLE
            PS C:\> Get-AzureDevOpsTokenCredential
            
            Creates a credential after prompting the user to enter the personal access token.

        .EXAMPLE
            PS C:\> Get-AzureDevOpsTokenCredential -TokenFile C:\temp\pat.txt
            
            Creates a credential from the encrypted token stored in C:\temp\pat.txt

        .EXAMPLE
            PS C:\> Get-AzureDevOpsTokenCredential -Token (Get-Content -Path C:\temp\mypat.txt)
            
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

function Get-AzureDevOpsBase64EncodedToken {
    <#
        .SYNOPSIS
           Get the personal access token encoded as Base64.

        .DESCRIPTION
            The Get-AzureDevOpsBase64EncodedToken function gets the personal access token encoded as Base64 for use in
            the authorization header for the web request.

        .PARAMETER  Credential
            Credential that contains the personal access token.

        .PARAMETER  Token
            Personal access token.

        .EXAMPLE
            PS C:\> Get-AzureDevOpsBase64EncodedToken -Credential $cred
            
            Get the base64 encoded personal access token from a Credential object.

        .EXAMPLE
            PS C:\> Get-AzureDevOpsBase64EncodedToken -Token (Get-Credential -Path C:\temp\mypat.txt)
            
            Get the base64 encoded personal access token from the token.
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(Position=0, Mandatory=$true, ParameterSetName='Credential')]
        [ValidateNotNullOrEmpty()]
        [PSCredential]
        $Credential,

        [Parameter(Position=0, Mandatory=$true, ParameterSetName='Token')]
        [ValidateNotNullOrEmpty()]
        [String]
        $Token
    )
    if ($Token) {
        [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(":$Token"))
    } else {
        [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(":$($Credential.GetNetworkCredential().Password)"))
    }
}
Export-ModuleMember -Function Get-AzureDevOpsBase64EncodedToken

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
        [Hashtable]
        $Body = $null
    )
    $encodedToken = if ($Token) {
        Get-AzureDevOpsBase64EncodedToken -Token $Token
    } else {
        Get-AzureDevOpsBase64EncodedToken -Credential $Credential
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
        $restParams.Add('Body', ($Body | ConvertTo-Json))
        $restParams.Add('ContentType', 'application/json')
    }
    
    if ($PSCmdlet.ShouldProcess($Uri, $Method)) {
        Invoke-RestMethod @restParams
    }
}
Export-ModuleMember -Function Invoke-AzureDevOpsRestMethod

function Invoke-AzureDevOpsWebRequest {
    <#
        .SYNOPSIS
           Sends a request to the Azure DevOps Services RESTful web service.

        .DESCRIPTION
            The Invoke-AzureDevOpsWebRequest function sends a request to the Azure DevOps Services RESTful web service.

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
            PS C:\> Invoke-AzureDevOpsWebRequest -Uri 'https://<organization>.visualstudio.com/_apis/distributedtask/pools/<id>/agents' -Body @{includeCapabilities='true'} -Credential $cred
            
            Sends a request to Azure DevOps to get the test agents and their capabilities.

        .NOTES
            This will return the status code, headers, and other information from the web request unlike Invoke-AzureDevOpsRestMethod.
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
        [Hashtable]
        $Body = $null
    )
    $encodedToken = if ($Token) {
        Get-AzureDevOpsBase64EncodedToken -Token $Token    
    } else {
        Get-AzureDevOpsBase64EncodedToken -Credential $Credential
    }

    $webRequestParams = @{
        'Uri' = $Uri
        'Headers' = @{Authorization = "Basic $encodedToken"}
        'Method' = $Method
    }

    if ($Method -eq 'GET') {
        $webRequestParams.Add('Body', $Body)
    }

    if ($Method -eq 'POST') {
        $webRequestParams.Add('Body', ($Body | ConvertTo-Json))
        $webRequestParams.Add('ContentType', 'application/json')
    }
    
    if ($PSCmdlet.ShouldProcess($Uri, $Method)) {
        Invoke-WebRequest @webRequestParams -UseBasicParsing
    }
}
Export-ModuleMember -Function Invoke-AzureDevOpsWebRequest

function Get-AzureDevOpsBaseUri {
    <#
        .SYNOPSIS
            Get the base URI for Azure DevOps Services RESTful web service.

        .DESCRIPTION
            The Get-AzureDevOpsBaseUri function gets the base URI for most requests to the Azure DevOps Services RESTful web service.

        .PARAMETER  Organization
            The name of the Azure DevOps organization.
        
        .PARAMETER  Project
            Project ID or project name.

        .EXAMPLE
            PS C:\> Get-AzureDevOpsBaseUri -Organization MyOrganization -Project MyProject
            
            Get the base URI of https://dev.azure.com/MyOrganization/MyProject used for most Azure DevOps requests.
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(Position=0, Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Organization,

        [Parameter(Position=1, Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Project
    )
    "https://dev.azure.com/$Organization/$Project"
}
Export-ModuleMember -Function Get-AzureDevOpsBaseUri

function Get-TfvcHistory {
    <#
        .SYNOPSIS
            Get the TFVC changeset history of the specified item.

        .DESCRIPTION
            The Get-TfvcHistory function gets the TFVC changeset history of the specified item.

        .PARAMETER  Organization
            The name of the Azure DevOps organization.
        
        .PARAMETER  Project
            Project ID or project name.
        
        .PARAMETER  Credential
            Credential that contains the personal access token.
        
        .PARAMETER  ApiVersion
            Version of the API to use.

        .PARAMETER  After
            Speficies the date and time that this cmdlet gets changesets after. The default is one day ago.

        .PARAMETER  Before
            Speficies the date and time that this cmdlet gets changesets before. The default is current time.
        
        .EXAMPLE
            PS C:\> Get-TfvcHistory -Organization MyOrganization -Project MyProject -ItemPath $itemPath
            
            Get the changeset history for the specified item path.
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(Position=0, Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Organization,

        [Parameter(Position=0, Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Project,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $ItemPath,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [String]
        $ApiVersion = 4.1,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [Datetime]
        $After = (Get-Date).AddDays(-1),

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [Datetime]
        $Before = (Get-Date), 

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [PSCredential]
        $Credential
    )
    $body = @{
        'api-version' = $ApiVersion
        'searchCriteria.itemPath' = $ItemPath
        'searchCriteria.fromDate' = $After.ToString('s')
        'searchCriteria.toDate' = $Before.ToString('s')
    }

    $baseUri = Get-AzureDevOpsBaseUri -Organization $Organization -Project $Project
    $uri = "$($baseUri)/_apis/tfvc/changesets"
    Invoke-AzureDevOpsRestMethod -Uri $uri -Body $body -Credential $Credential
}
Export-ModuleMember -Function Get-TfvcHistory

function Get-TestHistory {
    <#
        .SYNOPSIS
            Get the test history for the specified test.

        .DESCRIPTION
            The Get-TestHistory function gets the test results history for the specified test.

        .PARAMETER  Organization
            The name of the Azure DevOps organization.
        
        .PARAMETER  Project
            Project ID or project name.
        
        .PARAMETER  Credential
            Credential that contains the personal access token.
        
        .PARAMETER  ApiVersion
            Version of the API to use.
        
        .PARAMETER  TestName
            Automated test name of the TestCase. This must be the full qualified name of the test.
        
        .PARAMETER  GroupBy
            Group the result on the basis of TestResultGroupBy.

        .EXAMPLE
            PS C:\> Get-TestCaseHistory -Organization MyOrganization -Project MyProject -TestName NameSpace.ClassName.MyAwesomeTest
            
            Get the test history of MyAwesomeTest.

        .NOTES
            It was not particular clear on how to use the REST API for this. Information from https://github.com/MicrosoftDocs/vsts-rest-api-specs/issues/71 was helpful.
    #>
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter(Position=0, Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Organization,

        [Parameter(Position=1, Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Project,
        
        [Parameter(Position=2, Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $TestName,

        [Parameter(Position=3)]
        [ValidateNotNullOrEmpty()]
        [String]
        $ApiVersion = '5.0-preview.1',

        [Parameter(Position=4)]
        [ValidateNotNullOrEmpty()]
        [uint16]
        $GroupBy = 1,

        [Parameter(Position=5, Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [PSCredential]
        $Credential
    )
    $body = @{
        #'api-version' = $ApiVersion
        'automatedTestName' = $TestName
        'GroupBy' = $GroupBy
    }

    $baseUri = Get-AzureDevOpsBaseUri -Organization $Organization -Project $Project
    $uri = "$($baseUri)/_apis/test/Results/testhistory?api-version=$ApiVersion"
    Invoke-AzureDevOpsRestMethod -Uri $uri -Body $body -Method Post -Credential $Credential
}
Export-ModuleMember -Function Get-TestHistory
