<#

Script Name:  fix_ACCD_ServiceConfig.ps1
By:  Zack Thompson / Created:  5/2/2019
Version:  1.0.0 / Updated:  5/2/2019 / By:  ZT

Description:  This script writes the desired values to the Adobe ServiceConfig.xml file to resolve common issues with the ACCD.

#>

# ============================================================
# Define Variables
# ============================================================

$path_ServiceConfig = "%ProgramFiles(x86)%\Common Files\Adobe\OOBE\Configs"
$ServiceConfig = "${path_ServiceConfig}\ServiceConfig.xml"
$change_made = 0

# ============================================================
# Logic Functions
# ============================================================

function AddXMLElement($xml, $Parent, $Child) {
    # Creation of a node and its text
    # Write-Host "Here 1"
    $xmlElt = $xml.CreateElement("${Child}")
    # Add the node to the document
    # Write-Host "Here 2"
    $xml.SelectSingleNode("//${Parent}").AppendChild($xmlElt) | Out-Null
    # return $xml
}

function AddXMLText($xml, $Parent, $Element, $ElementText) {
    # Creation of a node and its text
    $xmlElt = $xml.CreateElement("${Element}")
    $xmlText = $xml.CreateTextNode("${ElementText}")
    $xmlElt.AppendChild($xmlText) | Out-Null
    # Add the node to the document
    $xml.SelectSingleNode("//${Parent}").AppendChild($xmlElt) | Out-Null
    # return $xml
}

# ============================================================
# Bits Staged...
# ============================================================

if ( Test-Path -Path $ServiceConfig ) {
    Write-Host "Config file exists, checking configuration..."

    # Read the xml file.
    [xml]$xml_ServiceConfig = Get-Content -Path "${ServiceConfig}"

    # Get the element so we can check if it exists.
    $AppsPanel = $( $xml_ServiceConfig.SelectSingleNode("//panel") | Where-Object { $_.name -eq "AppsPanel" } )
    if ( $AppsPanel ) {
        if ( $AppsPanel.visible -eq $false ){
            Write-Host "AppsPanel is hidden, correcting..."
            $AppsPanel.visible = "${true}"
            $change_made = 1
        }
    }
    else {
        Write-host "AppsPanel is not configured, setting proper values..."
        AddXMLElement $xml_ServiceConfig "config" "panel"
        AddXMLText $xml_ServiceConfig "panel" "name" "AppsPanel"
        AddXMLText $xml_ServiceConfig "panel" "visible" $true
    }

    Write-host "AppsPanel visible: " $($xml_ServiceConfig.SelectSingleNode("//panel") | Where-Object { $_.name -eq "AppsPanel" }).visible

    # Get the element so we can check if it exists.
    $SelfServeInstalls = $( $xml_ServiceConfig.SelectSingleNode("//feature") | Where-Object { $_.name -eq "SelfServeInstalls" } )
    if ( $SelfServeInstalls ) {
        if ( $SelfServeInstalls.enabled -eq $false ){
            Write-Host "SelfServeInstalls is disabled, correcting..."
            $SelfServeInstalls.enabled = "${true}"
            $change_made = 1
        }
    }
    else {
        Write-host "SelfServeInstalls is not configured, setting proper values..."
        AddXMLElement $xml_ServiceConfig "config" "feature"
        AddXMLText $xml_ServiceConfig "feature" "name" "SelfServeInstalls"
        AddXMLText $xml_ServiceConfig "feature" "enabled" $true
    }

    Write-host "SelfServeInstalls enabled: " $($xml_ServiceConfig.SelectSingleNode("//feature") | Where-Object { $_.name -eq "SelfServeInstalls" }).enabled

    if ($change_made -eq 1 ) {
        Write-Host "Saving configuration..."
        $xml_ServiceConfig.Save($ServiceConfig)
        Write-Host "Result:  Updates made."
    }
    else {
        Write-Host "Result:  No changes made."
    }
}
else {
    Write-Host "Config file does not exist."

    if ( !( Test-Path -Path $path_ServiceConfig ) ) {
        Write-Host "Creating directory structure..."
        New-Item -Path "${path_ServiceConfig}" -ItemType Directory | Out-Null
    }

    # Create an XML Object
    [xml]$xml_Contents = New-Object Xml

    # Creat the root node
    $xmlElt = $xml_Contents.CreateElement("config")
    # Add the node to the XML Object
    $xml_Contents.AppendChild($xmlElt) | Out-Null

    # # Build the xml elements.
    AddXMLElement $xml_Contents "config" "panel"
    AddXMLText $xml_Contents "panel" "name" "AppsPanel"
    AddXMLText $xml_Contents "panel" "visible" $true
    AddXMLElement $xml_Contents "config" "feature"
    AddXMLText $xml_Contents "feature" "name" "SelfServeInstalls"
    AddXMLText $xml_Contents "feature" "enabled" $true

    Write-Host "Saving configuration to disk..."
    $xml_Contents.Save($ServiceConfig)
    Write-Host "Result:  Updates made."
}
