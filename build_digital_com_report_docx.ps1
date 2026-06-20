param(
    [string]$OutputPath = "Digital_Com_Project_Report.docx"
)

$ErrorActionPreference = "Stop"

function Escape-Xml {
    param([Parameter(Mandatory=$true)][string]$Text)
    return [System.Security.SecurityElement]::Escape($Text)
}

function Run-Xml {
    param(
        [Parameter(Mandatory=$true)][string]$Text,
        [int]$SizeHalfPoints = 22,
        [string]$Font = "Calibri",
        [bool]$Bold = $false,
        [bool]$Italic = $false,
        [string]$Color = "000000"
    )
    $rPr = "<w:rPr><w:rFonts w:ascii='$Font' w:hAnsi='$Font'/><w:sz w:val='$SizeHalfPoints'/><w:szCs w:val='$SizeHalfPoints'/><w:color w:val='$Color'/>"
    if ($Bold) { $rPr += "<w:b/>" }
    if ($Italic) { $rPr += "<w:i/>" }
    $rPr += "</w:rPr>"
    return "<w:r>$rPr<w:t xml:space='preserve'>$(Escape-Xml $Text)</w:t></w:r>"
}

function Para-Xml {
    param(
        [Parameter(Mandatory=$true)][string]$Text,
        [string]$Align = "left",
        [int]$Before = 0,
        [int]$After = 120,
        [int]$Line = 276,
        [int]$SizeHalfPoints = 22,
        [string]$Font = "Calibri",
        [bool]$Bold = $false,
        [bool]$Italic = $false,
        [string]$Color = "000000",
        [string]$Style = $null,
        [object]$NumId = $null
    )
    $pPr = "<w:pPr>"
    if ($Style) { $pPr += "<w:pStyle w:val='$Style'/>" }
    if ($Align -ne "left") { $pPr += "<w:jc w:val='$Align'/>" }
    $pPr += "<w:spacing w:before='$Before' w:after='$After' w:line='$Line' w:lineRule='auto'/>"
    if ($null -ne $NumId) {
        $pPr += "<w:numPr><w:ilvl w:val='0'/><w:numId w:val='$NumId'/></w:numPr>"
    }
    $pPr += "</w:pPr>"
    return "<w:p>$pPr$(Run-Xml -Text $Text -SizeHalfPoints $SizeHalfPoints -Font $Font -Bold $Bold -Italic $Italic -Color $Color)</w:p>"
}

function Heading-Xml {
    param(
        [Parameter(Mandatory=$true)][string]$Text,
        [int]$Level = 1
    )
    switch ($Level) {
        1 { return (Para-Xml -Text $Text -Before 200 -After 80 -Line 276 -SizeHalfPoints 32 -Bold $true -Color "2E74B5") }
        2 { return (Para-Xml -Text $Text -Before 160 -After 60 -Line 276 -SizeHalfPoints 26 -Bold $true -Color "2E74B5") }
        default { return (Para-Xml -Text $Text -Before 120 -After 40 -Line 276 -SizeHalfPoints 24 -Bold $true -Color "1F4D78") }
    }
}

function Cell-Xml {
    param(
        [Parameter(Mandatory=$true)][string]$Text,
        [int]$Width,
        [bool]$Bold = $false,
        [string]$Align = "left"
    )
    $tcPr = "<w:tcPr><w:tcW w:w='$Width' w:type='dxa'/><w:tcMar><w:top w:w='80' w:type='dxa'/><w:start w:w='120' w:type='dxa'/><w:bottom w:w='80' w:type='dxa'/><w:end w:w='120' w:type='dxa'/></w:tcMar></w:tcPr>"
    $p = Para-Xml -Text $Text -Align $Align -Before 0 -After 0 -Line 240 -SizeHalfPoints 20 -Bold:$Bold
    return "<w:tc>$tcPr$p</w:tc>"
}

function Table-Xml {
    param(
        [Parameter(Mandatory=$true)][string[]]$Headers,
        [Parameter(Mandatory=$true)][object[][]]$Rows,
        [int[]]$Widths
    )
    $tblPr = @"
<w:tblPr>
  <w:tblW w:w='9360' w:type='dxa'/>
  <w:tblInd w:w='0' w:type='dxa'/>
  <w:tblBorders>
    <w:top w:val='single' w:sz='4' w:space='0' w:color='D9D9D9'/>
    <w:left w:val='single' w:sz='4' w:space='0' w:color='D9D9D9'/>
    <w:bottom w:val='single' w:sz='4' w:space='0' w:color='D9D9D9'/>
    <w:right w:val='single' w:sz='4' w:space='0' w:color='D9D9D9'/>
    <w:insideH w:val='single' w:sz='4' w:space='0' w:color='D9D9D9'/>
    <w:insideV w:val='single' w:sz='4' w:space='0' w:color='D9D9D9'/>
  </w:tblBorders>
</w:tblPr>
"@
    $grid = "<w:tblGrid>" + (($Widths | ForEach-Object { "<w:gridCol w:w='$_'/>" }) -join "") + "</w:tblGrid>"
    $out = New-Object System.Text.StringBuilder
    [void]$out.Append("<w:tbl>$tblPr$grid")
    $headerCells = for ($i=0; $i -lt $Headers.Count; $i++) { Cell-Xml -Text $Headers[$i] -Width $Widths[$i] -Bold $true }
    [void]$out.Append("<w:tr>" + ($headerCells -join "") + "</w:tr>")
    foreach ($row in $Rows) {
        $cells = for ($i=0; $i -lt $row.Count; $i++) { Cell-Xml -Text ([string]$row[$i]) -Width $Widths[$i] }
        [void]$out.Append("<w:tr>" + ($cells -join "") + "</w:tr>")
    }
    [void]$out.Append("</w:tbl>")
    return $out.ToString()
}

$sections = @()

$sections += Para-Xml -Text "Digital Com Project Report" -Align center -Before 0 -After 40 -Line 240 -SizeHalfPoints 52 -Font "Calibri" -Bold $true
$sections += Para-Xml -Text "Digital commerce system for smart homes, villas, city assets, commercial spaces, and yachts" -Align center -Before 0 -After 180 -Line 240 -SizeHalfPoints 22 -Font "Calibri"
$sections += Para-Xml -Text "This report defines the product vision, scope, user workflow, delivery phases, and technical direction for the Digital Com platform." -Before 0 -After 160 -Line 276 -SizeHalfPoints 22 -Font "Calibri"

$sections += Heading-Xml -Text "1. Project Overview" -Level 1
$sections += Para-Xml -Text "Digital Com is a digital commerce and smart-environment platform for the smart homes industry. It allows a customer to create an account, upload or scan a property blueprint, and generate a 3D digital twin in Unity before any physical implementation begins. The system is intended to support residential villas, commercial spaces, smart city assets, yachts, and specialized environments through a common commerce and visualization layer."
$sections += Para-Xml -Text "The core idea is to move the selling, planning, and commissioning process into a simulation-first workflow. Instead of purchasing smart-home or smart-building services blindly, the customer can see the property model, place devices and services inside it, test interactions, review the effect of each service, and decide what to deploy. The platform then becomes a bridge between commercial selection, digital design, implementation, and later live synchronization with the real installation. After implementation, the same digital twin can be connected to the real Crestron environment so the user can visualize the live state of the property, control supported functions more accurately, and speed up debugging and support."

$sections += Heading-Xml -Text "2. Business Problem and Project Objective" -Level 1
$sections += Para-Xml -Text "The smart-home and automation market often suffers from disconnected planning tools, unclear purchasing decisions, weak visualization of the final result, and limited support after installation. Customers usually need to imagine the end result from a drawing or from a technical bill of materials. That makes sales slower, support harder, and upgrades less transparent."
$sections += Para-Xml -Text "Digital Com addresses that gap by combining e-commerce, digital twin generation, AI-assisted modeling, augmented reality, simulation, billing, customer support, and post-deployment system monitoring in one platform. The objective is to reduce friction for the customer, improve the accuracy of system design, accelerate sales, and provide a long-term operational layer for maintenance and debugging."

$sections += Heading-Xml -Text "3. Scope of Services" -Level 1
$sections += Para-Xml -Text "The project is based on the service groups shown in the original concept description. These groups define the solution catalog that the customer can browse and insert into a property model."
$sections += Table-Xml -Headers @("Vertical","Included Services") -Widths @(1800,7560) -Rows @(
    @("Smart City","Traffic management, waste management systems, water monitoring and control, smart street lighting"),
    @("Commercial","Automated power management, smart lighting, smart digital signage, integrated surveillance systems, interactive access control"),
    @("Residential","IP 4K video distribution, smart energy meters, smart AC and heating, IP multi-zone audio, lighting and shading systems, home security, geo-fencing"),
    @("Yacht","Hidden motorized TV systems, monitored cloud services, marine audio, marine communication"),
    @("Other","Guest room management, digitized nursing systems, lighting management systems, automated pro audio")
)

$sections += Heading-Xml -Text "4. Platform Concept" -Level 1
$sections += Para-Xml -Text "Digital Com functions as a digital commerce platform with a built-in digital twin. The commercial layer is not separated from the technical layer. The customer can discover services, configure them in the model, test the result, and then proceed toward implementation or subscription-based live synchronization."
$sections += Para-Xml -Text "The platform is designed around Crestron and KNX-based automation ecosystems, so the product catalog and downstream implementation logic can align with professional smart-building hardware and control standards."

$sections += Heading-Xml -Text "5. User Workflow" -Level 1
$sections += Para-Xml -Text "The user signs up and creates a personal account." -NumId 1
$sections += Para-Xml -Text "The user uploads or scans the blueprint of a villa, house, city property, commercial site, or yacht." -NumId 1
$sections += Para-Xml -Text "The system converts the blueprint into a layered 3D model using Unity-based processing." -NumId 1
$sections += Para-Xml -Text "The user can optionally upload a photo taken through AR so the platform can place real objects and contextual details inside the model." -NumId 1
$sections += Para-Xml -Text "The user browses the available services from the catalog and inserts them into the digital twin." -NumId 1
$sections += Para-Xml -Text "The user moves through the model in 2D or 3D, similar to a game, to test how the selected features behave." -NumId 1
$sections += Para-Xml -Text "The user may activate VR access depending on budget and sponsorship model." -NumId 1
$sections += Para-Xml -Text "The user can interact with the assistant chatbot or create support conversations with the IT team." -NumId 1
$sections += Para-Xml -Text "The user can pay deposits and maintenance fees through the platform." -NumId 1
$sections += Para-Xml -Text "After implementation, the same digital twin can be reused for live monitoring, version tracking, support, and future upgrades." -NumId 1

$sections += Heading-Xml -Text "6. Digital Twin and 3D Model Generation" -Level 1
$sections += Para-Xml -Text "A key feature of the platform is the conversion of blueprint files into a navigable model. The system accepts PDF, JPG, PNG, or similar property-source inputs and converts them into a structured digital twin. Each floor can be represented as a separate layer so that the user can inspect the architecture independently or view the building as a complete model."
$sections += Para-Xml -Text "Unity is used as the game-engine environment for rendering, interaction, movement, object placement, and simulation. The model is not only a static viewer. It is intended to behave as an interactive environment where the user can walk through rooms, test service layouts, evaluate experience flow, and understand how the property will function after deployment."
$sections += Para-Xml -Text "The AR scanning step adds another layer of realism. Instead of relying only on a flat blueprint, the user can scan or photograph the physical home and insert real-world objects or visual references into the model. This helps the digital twin reflect the intended final environment more accurately."

$sections += Heading-Xml -Text "7. Interaction, Visualization, and Control" -Level 1
$sections += Para-Xml -Text "The platform supports both a 2D planning view and a 3D immersive view. The 2D view is useful for quick layout checks, service placement, and floor-by-floor comparison. The 3D view is used to move through rooms, inspect device positions, and test the appearance and operation of the configured solution."
$sections += Para-Xml -Text "The interface is intended to mirror the real smart-home interface as closely as possible. Tablets and mobile devices can be used inside the simulation model with the same style of controls that the real installation will expose later. That makes the digital twin a practical preview of the final system, not just a marketing model."
$sections += Para-Xml -Text "Lightmaps and heatmaps can be used to visualize physical properties of the space, such as lighting behavior, energy usage, comfort zones, or traffic patterns through the property. These overlays help the customer and the support team understand how the environment behaves under different scenarios."

$sections += Heading-Xml -Text "8. VR Support" -Level 1
$sections += Para-Xml -Text "VR support is optional and budget-dependent. The project can be delivered in two possible ways. The first option is a lower-cost setup where a controller connects to a mobile phone and is used together with a PS5 controller or similar input device. The second option is a full VR set, such as a complete headset and controller package, for a more immersive and premium experience."
$sections += Para-Xml -Text "VR is valuable when the customer needs to validate room scale, interior placement, visual comfort, and the overall feeling of the space before implementation."

$sections += Heading-Xml -Text "9. Commerce, Billing, and Customer Support" -Level 1
$sections += Para-Xml -Text "Digital Com includes a billing system that can collect deposits and maintenance fees. This is important because the platform is not only a visualization tool. It is also the transactional entry point for the project lifecycle."
$sections += Para-Xml -Text "An AI chatbot assists the user during browsing and configuration. The same chat layer can be used to open direct communication with the support or IT department. This creates a faster escalation path when the user needs help selecting services, checking compatibility, or troubleshooting a deployment issue."

$sections += Heading-Xml -Text "10. Phase 1 and Phase 2 Delivery Model" -Level 1
$sections += Para-Xml -Text "The project is best understood in two phases."
$sections += Table-Xml -Headers @("Phase","Description") -Widths @(1800,7560) -Rows @(
    @("Phase 1","Digital commerce and digital twin creation before implementation. The user uploads the blueprint, configures services, tests the model, pays required fees, and reviews the expected outcome."),
    @("Phase 2","Live synchronization with the real installation after deployment. The saved blueprint and model are reused so the platform can mirror the actual property state and support long-term maintenance.")
)
$sections += Para-Xml -Text "In phase 2, the platform uses a gateway server that validates whether the user has an active subscription. Only valid users are forwarded to the central server. Periodic traffic is collected, and when the user opens the app the server returns the latest state data so the digital twin can reflect the real environment."
$sections += Para-Xml -Text "This phase turns the platform into a live operational layer for home, yacht, or commercial assets. The twin becomes a continuously updated representation of the deployed system rather than a one-time sales model."

$sections += Heading-Xml -Text "11. Remote Support and Debugging Model" -Level 1
$sections += Para-Xml -Text "A major operational benefit of the project is remote debugging. If a support issue requires deeper intervention, a small tunneling server can be used to create a secure support path. That server remains closed by default and is opened only with the proper encryption key and controlled user approval."
$sections += Para-Xml -Text "The user can expose a port or a controlled access point to support, and the support engineer connects as a client to the Crestron-side user environment to apply the fix. This approach reduces service time and avoids unnecessary physical intervention when a remote fix is possible."
$sections += Para-Xml -Text "The architecture also assumes message-broker-based updates from the smart-home processor into a database. Engineers can inspect the home state, review event history, and inject approved code updates or configuration changes when the workflow requires it."

$sections += Heading-Xml -Text "12. Version Tracking and Data Retention" -Level 1
$sections += Para-Xml -Text "The system stores blueprint versions and digital twin versions so the customer can rebuild the same environment later if required. This is important for upgrades, rework, multi-phase construction, and long-term support."
$sections += Para-Xml -Text "Version tracking also helps the support team understand what changed between deployments, which layout was active at a given time, and which service set was selected in each phase of the project."

$sections += Heading-Xml -Text "13. Security and Access Control" -Level 1
$sections += Para-Xml -Text "Because the platform interacts with live automation environments, security has to be treated as a core feature. Account access, subscription validation, gateway filtering, encrypted tunneling, and role-based support access should all be enforced."
$sections += Para-Xml -Text "The goal is to ensure that only authenticated customers and authorized support staff can interact with the live environment, while all other traffic remains blocked or read-only."

$sections += Heading-Xml -Text "14. Expected Benefits" -Level 1
$sections += Para-Xml -Text "Better sales conversion because customers can see the final result before buying." -NumId 1
$sections += Para-Xml -Text "Lower design ambiguity because services are tested in a realistic digital twin." -NumId 1
$sections += Para-Xml -Text "Faster commissioning and clearer handover between sales, engineering, and support." -NumId 1
$sections += Para-Xml -Text "Improved maintenance because the live system can be monitored and versioned." -NumId 1
$sections += Para-Xml -Text "Stronger customer confidence through AR, VR, and interactive simulations." -NumId 1
$sections += Para-Xml -Text "A reusable platform that can serve smart homes, villas, yachts, city assets, and commercial projects." -NumId 1

$sections += Heading-Xml -Text "15. Conclusion" -Level 1
$sections += Para-Xml -Text "Digital Com is a smart-home commerce and digital twin platform that combines property scanning, AI-assisted model generation, Unity-based simulation, optional VR, customer billing, chatbot support, and post-deployment synchronization. Its main value is that it connects the commercial purchase process to the real technical implementation in one continuous workflow."
$sections += Para-Xml -Text "The project is strong because it solves a real market problem: customers want to understand what they are buying before the hardware is installed. By giving them a live, interactive, versioned digital twin, Digital Com makes the sales process clearer, the support process faster, and the final smart environment easier to manage."

$bodyXml = ($sections -join "`r`n")

$documentXml = @"
<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<w:document xmlns:wpc='http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas'
 xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006'
 xmlns:o='urn:schemas-microsoft-com:office:office'
 xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
 xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'
 xmlns:v='urn:schemas-microsoft-com:vml'
 xmlns:wp14='http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing'
 xmlns:wp='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
 xmlns:w10='urn:schemas-microsoft-com:office:word'
 xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
 xmlns:w14='http://schemas.microsoft.com/office/word/2010/wordml'
 xmlns:wpg='http://schemas.microsoft.com/office/word/2010/wordprocessingGroup'
 xmlns:wpi='http://schemas.microsoft.com/office/word/2010/wordprocessingInk'
 xmlns:wne='http://schemas.microsoft.com/office/word/2006/wordml'
 mc:Ignorable='w14 wp14'>
  <w:body>
    $bodyXml
    <w:sectPr>
      <w:pgSz w:w='12240' w:h='15840'/>
      <w:pgMar w:top='1440' w:right='1440' w:bottom='1440' w:left='1440' w:header='720' w:footer='720' w:gutter='0'/>
      <w:cols w:space='720'/>
      <w:docGrid w:linePitch='360'/>
    </w:sectPr>
  </w:body>
</w:document>
"@

$stylesXml = @"
<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<w:styles xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:style w:type='paragraph' w:default='1' w:styleId='Normal'>
    <w:name w:val='Normal'/>
    <w:qFormat/>
    <w:rPr>
      <w:rFonts w:ascii='Calibri' w:hAnsi='Calibri'/>
      <w:sz w:val='22'/>
      <w:szCs w:val='22'/>
    </w:rPr>
  </w:style>
  <w:style w:type='paragraph' w:styleId='Heading1'>
    <w:name w:val='heading 1'/>
    <w:basedOn w:val='Normal'/>
    <w:next w:val='Normal'/>
    <w:qFormat/>
    <w:rPr>
      <w:rFonts w:ascii='Calibri' w:hAnsi='Calibri'/>
      <w:b/>
      <w:sz w:val='32'/>
      <w:szCs w:val='32'/>
      <w:color w:val='2E74B5'/>
    </w:rPr>
  </w:style>
  <w:style w:type='paragraph' w:styleId='Heading2'>
    <w:name w:val='heading 2'/>
    <w:basedOn w:val='Normal'/>
    <w:next w:val='Normal'/>
    <w:qFormat/>
    <w:rPr>
      <w:rFonts w:ascii='Calibri' w:hAnsi='Calibri'/>
      <w:b/>
      <w:sz w:val='26'/>
      <w:szCs w:val='26'/>
      <w:color w:val='2E74B5'/>
    </w:rPr>
  </w:style>
  <w:style w:type='paragraph' w:styleId='Heading3'>
    <w:name w:val='heading 3'/>
    <w:basedOn w:val='Normal'/>
    <w:next w:val='Normal'/>
    <w:qFormat/>
    <w:rPr>
      <w:rFonts w:ascii='Calibri' w:hAnsi='Calibri'/>
      <w:b/>
      <w:sz w:val='24'/>
      <w:szCs w:val='24'/>
      <w:color w:val='1F4D78'/>
    </w:rPr>
  </w:style>
</w:styles>
"@

$numberingXml = @"
<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<w:numbering xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:abstractNum w:abstractNumId='0'>
    <w:lvl w:ilvl='0'>
      <w:start w:val='1'/>
      <w:numFmt w:val='bullet'/>
      <w:lvlText w:val='•'/>
      <w:lvlJc w:val='left'/>
      <w:pPr><w:ind w:left='720' w:hanging='360'/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId='1'>
    <w:lvl w:ilvl='0'>
      <w:start w:val='1'/>
      <w:numFmt w:val='decimal'/>
      <w:lvlText w:val='%1.'/>
      <w:lvlJc w:val='left'/>
      <w:pPr><w:ind w:left='720' w:hanging='360'/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId='1'><w:abstractNumId w:val='0'/></w:num>
  <w:num w:numId='2'><w:abstractNumId w:val='1'/></w:num>
</w:numbering>
"@

$settingsXml = @"
<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<w:settings xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:zoom w:percent='100'/>
  <w:compat/>
</w:settings>
"@

$coreXml = @"
<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<cp:coreProperties xmlns:cp='http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
 xmlns:dc='http://purl.org/dc/elements/1.1/'
 xmlns:dcterms='http://purl.org/dc/terms/'
 xmlns:dcmitype='http://purl.org/dc/dcmitype/'
 xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'>
  <dc:title>Digital Com Project Report</dc:title>
  <dc:subject>Smart home digital commerce report</dc:subject>
  <dc:creator>Codex</dc:creator>
  <cp:keywords>Digital Com, smart home, digital twin, Crestron, KNX</cp:keywords>
  <dc:description>Project report for a smart home digital commerce and digital twin platform.</dc:description>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type='dcterms:W3CDTF'>2026-06-19T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type='dcterms:W3CDTF'>2026-06-19T00:00:00Z</dcterms:modified>
</cp:coreProperties>
"@

$appXml = @"
<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Properties xmlns='http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'
 xmlns:vt='http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'>
  <Application>Microsoft Word</Application>
</Properties>
"@

$contentTypesXml = @"
<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>
  <Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/>
  <Default Extension='xml' ContentType='application/xml'/>
  <Override PartName='/word/document.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'/>
  <Override PartName='/word/styles.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml'/>
  <Override PartName='/word/numbering.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml'/>
  <Override PartName='/word/settings.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'/>
  <Override PartName='/docProps/core.xml' ContentType='application/vnd.openxmlformats-package.core-properties+xml'/>
  <Override PartName='/docProps/app.xml' ContentType='application/vnd.openxmlformats-officedocument.extended-properties+xml'/>
</Types>
"@

$relsXml = @"
<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>
  <Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/>
  <Relationship Id='rId2' Type='http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties' Target='docProps/core.xml'/>
  <Relationship Id='rId3' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties' Target='docProps/app.xml'/>
</Relationships>
"@

$docRelsXml = @"
<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>
  <Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles' Target='styles.xml'/>
  <Relationship Id='rId2' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering' Target='numbering.xml'/>
  <Relationship Id='rId3' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings' Target='settings.xml'/>
</Relationships>
"@

$tempRoot = Join-Path $env:TEMP ("digital_com_report_" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tempRoot | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tempRoot "_rels") | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tempRoot "docProps") | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tempRoot "word") | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tempRoot "word/_rels") -Force | Out-Null

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText((Join-Path $tempRoot "[Content_Types].xml"), $contentTypesXml, $utf8NoBom)
[System.IO.File]::WriteAllText((Join-Path $tempRoot "_rels/.rels"), $relsXml, $utf8NoBom)
[System.IO.File]::WriteAllText((Join-Path $tempRoot "docProps/core.xml"), $coreXml, $utf8NoBom)
[System.IO.File]::WriteAllText((Join-Path $tempRoot "docProps/app.xml"), $appXml, $utf8NoBom)
[System.IO.File]::WriteAllText((Join-Path $tempRoot "word/document.xml"), $documentXml, $utf8NoBom)
[System.IO.File]::WriteAllText((Join-Path $tempRoot "word/styles.xml"), $stylesXml, $utf8NoBom)
[System.IO.File]::WriteAllText((Join-Path $tempRoot "word/numbering.xml"), $numberingXml, $utf8NoBom)
[System.IO.File]::WriteAllText((Join-Path $tempRoot "word/settings.xml"), $settingsXml, $utf8NoBom)
[System.IO.File]::WriteAllText((Join-Path $tempRoot "word/_rels/document.xml.rels"), $docRelsXml, $utf8NoBom)

$target = Join-Path (Get-Location) $OutputPath
if (Test-Path $target) { Remove-Item $target -Force }
$zipTarget = [System.IO.Path]::ChangeExtension($target, ".zip")
if (Test-Path $zipTarget) { Remove-Item $zipTarget -Force }
Compress-Archive -Path (Join-Path $tempRoot "*") -DestinationPath $zipTarget
Move-Item -Force $zipTarget $target
Remove-Item -Recurse -Force $tempRoot
