<#
Begin license text.
Copyright (c) 2013-2018 Raffael Herrmann
Copyright (c) 2019 Dr. Tobias Weltner
Copyright (c) 2020 Sam Jansson

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

End license text.
#>

<#
.SYNOPSIS
This script generates QR codes based on attributes of user objects within a hardcoded LDAP directory.

.DESCRIPTION
This script extends the functionality of 'QRCodeGenerator' by Dr. Tobias Weltner by reading vCard attributes directly from Active Directory.
vCards are used as a basis for generating QR codes that holds contact information for Active Directory users.
The QR codes are saved as .PNG-files in the subdirectory 'QR Codes' of the script location.

This script uses the 'QRCoder' library.

NOTE:
1) This script is incompatible with PowerShell Core.
2) This script must be run with administrative privileges to enable installing/loading of necessary modules.
3) Change the values of $DarkModColor and $LightModColor for different colors. 
4) You must edit this script and specify LDAP -SearchBase.
5) All AD users in -SearchBase must have the following attributes populated:

Name
GivenName
Surname
Company
EmailAddress
OfficePhone
MobilePhone
Title

.LINK
https://github.com/TobiasPSP/Modules.QRCodeGenerator
https://github.com/codebude/QRCoder
https://github.com/sakeja/ADQR
#>

<#
Required for installing/loading of necessary modules.
https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.security/set-executionpolicy?view=powershell-7
#>
Set-ExecutionPolicy Unrestricted -Scope Process -Force

<#
We can't load QRCoder.dll by
[Reflection.Assembly]::LoadFrom( (Resolve-Path "Path to QRCoder.dll"))
since the ability to execute code in assemblies loaded from remote locations is disabled by default starting with .NET Framework 4.
https://docs.microsoft.com/en-us/dotnet/api/system.reflection.assembly.loadfrom?view=netstandard-2.1
#>

# So we base64 encode QRCoder.dll into a string so that we can bypass this restriction.
$QRCoderDLL = Get-Content -Path "$(Get-Location)\QRCoder\bin\Release\netstandard2.0\QRCoder.dll" -Encoding Byte
$Base64Payload = [System.Convert]::ToBase64String($QRCoderDLL)
$null = [System.Reflection.Assembly]::Load([System.Convert]::FromBase64String($Base64Payload))

function GenVCardQRCode {
    param (
        [Parameter(Mandatory)]
        [string]
        $FirstName,

        [Parameter(Mandatory)]
        [string]
        $LastName,

        [Parameter(Mandatory)]
        [string]
        $Company,

        [Parameter(Mandatory)]
        [string]
        $TelWork,

        [Parameter(Mandatory)]
        [string]
        $TelCell,

        [Parameter(Mandatory)]
        [string]
        $Title,

        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]
        $Email,

        <#
        'Width' = 'pixelsPerModule' - "The pixel size each b/w module is drawn"
        https://github.com/codebude/QRCoder/wiki/Advanced-usage---QR-Code-renderers#25-pngbyteqrcode-renderer-in-detail
        #>
        [ValidateRange(10,2000)]
        [int]
        $Width = 100,

        [Parameter(Mandatory)]
        [string]
        $OutPath
    )

    $Name = "$FirstName $LastName"

    <#
    For info on how 'QRCoder' handles color, see:
    https://github.com/codebude/QRCoder

    and

    https://github.com/codebude/QRCoder/wiki/Advanced-usage---QR-Code-renderers#25-pngbyteqrcode-renderer-in-detail

    NOTE:
    Some color combinations can't be read by QR scanners due to insufficient contrast.
    #>

    <#
    These are the standard black and white combinations used in most QR-codes.
    #[Byte[]] $DarkModColor = 0x00, 0x00, 0x00
    #[Byte[]] $LightModColor = 0xff, 0xff, 0xff
    #>

    # An example combination with sufficient contrast for most QR scanners.
    [Byte[]] $DarkModColor = 0x2b, 0x57, 0x95
    [Byte[]] $LightModColor = 0xff, 0xff, 0xff
    #>

    <#
    For reference, the vCard 3.0 specification is published here:
    https://tools.ietf.org/html/rfc6350

    A summarized version can be found on Wikipedia:
    https://en.wikipedia.org/wiki/VCard#vCard_3.0

    Documentation for the implemented parts of the vCard 3.0 specification in the 'QRCoder' library:
    https://github.com/codebude/QRCoder/wiki/Advanced-usage---Payload-generators#35-contactdata-mecardvcard
    #>

    $VCard = @"
BEGIN:VCARD
VERSION:3.0
N:$LastName;$FirstName
FN:$Name
ORG:$Company
EMAIL:$Email
TEL;TYPE=WORK:$TelWork
TEL;TYPE=CELL:$TelCell
TITLE:$Title
END:VCARD
"@

    $QRGenerator = New-Object -TypeName QRCoder.QRCodeGenerator

    <#
    The parameters for the 'CreateQrCode' function are documented here:
    https://github.com/codebude/QRCoder/wiki/How-to-use-QRCoder

    Below we force UTF-8 encoding (represented by the argument '$True') and use the error correction level 'Q' (25%).
    #>
    $Data = $QRGenerator.CreateQrCode($VCard, 'Q', $True)

    $QRCode = New-Object -TypeName QRCoder.PngByteQRCode -ArgumentList ($Data)

    # This (base) method can be used for saving black and white QR codes.
    #$ByteArray = $QRCode.GetGraphic($Width)

    # This overloaded method is used for saving colored QR codes.
    $ByteArray = $QRCode.GetGraphic($Width, $DarkModColor, $LightModColor)

    [System.IO.File]::WriteAllBytes($OutPath, $ByteArray)
}

$OutputFolderName = "QR Codes"
$OutputFilenameExtension = ".png"

# Silently ('Out-Null') create the output directory.
New-Item -Path "$(Get-Location)\" -Name $OutputFolderName -ItemType "Directory" -Force | Out-Null

# Load all AD users.
$ADUsers = @(Get-ADUser -SearchBase "OU=Contoso,OU=Users,DC=Contoso,DC=SE" -Filter *)

# Load the contact info for every AD user.
ForEach ($ADUser in $ADUsers) {
    $OutputFilename = Get-ADUser $ADUser | ForEach-Object {$_.Name + " - " + $_.GivenName + " " + $_.Surname + " QR vCard" + $OutputFilenameExtension}
    $GivenName = Get-ADUser $ADUser | ForEach-Object {$_.GivenName}
    $Surname = Get-ADUser $ADUser | ForEach-Object {$_.Surname}
    $Company = Get-ADUser $ADUser -Properties * | Select-Object Company | ForEach-Object {$_.Company}
    $EmailAddress = Get-ADUser $ADUser -Properties * | Select-Object EmailAddress | ForEach-Object {$_.EmailAddress}
    $OfficePhone = Get-ADUser $ADUser -Properties * | Select-Object OfficePhone | ForEach-Object {$_.OfficePhone}
    $MobilePhone = Get-ADUser $ADUser -Properties * | Select-Object MobilePhone | ForEach-Object {$_.MobilePhone}
    $Title = Get-ADUser $ADUser -Properties * | Select-Object Title | ForEach-Object {$_.Title}

    # Create a vCard and QR Code for every AD user.
    GenVCardQRCode -FirstName $GivenName -LastName $Surname -Company $Company -Title $Title -TelWork $OfficePhone -TelCell $MobilePhone -Email $EmailAddress -OutPath "$(Get-Location)\$OutputFolderName\$OutputFilename"
}
