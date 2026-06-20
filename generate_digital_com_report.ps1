param(
    [string]$OutputPath = "Digital_Com_Project_Report.docx"
)

$ErrorActionPreference = "Stop"

function Set-RunFont {
    param(
        [Parameter(Mandatory=$true)] $Run,
        [string]$Name = "Calibri",
        [int]$Size = 11,
        [bool]$Bold = $false
    )
    $Run.Font.Name = $Name
    $Run.Font.Size = $Size
    $Run.Font.Bold = $Bold
}

function Add-Paragraph {
    param(
        [Parameter(Mandatory=$true)] $Doc,
        [Parameter(Mandatory=$true)] [string]$Text,
        [string]$Style = "Normal",
        [string]$FontName = "Calibri",
        [int]$FontSize = 11,
        [bool]$Bold = $false,
        [int]$SpaceAfter = 6,
        [int]$SpaceBefore = 0,
        [switch]$PassThru
    )
    $p = $Doc.Paragraphs.Add()
    $p.Range.Text = $Text
    $p.Range.Style = $Style
    $p.Format.SpaceAfter = $SpaceAfter
    $p.Format.SpaceBefore = $SpaceBefore
    $p.Format.LineSpacingRule = 5
    $p.Range.Font.Name = $FontName
    $p.Range.Font.Size = $FontSize
    $p.Range.Font.Bold = $Bold
    if ($PassThru) { return $p }
}

function Add-Bullet {
    param(
        [Parameter(Mandatory=$true)] $Doc,
        [Parameter(Mandatory=$true)] [string]$Text
    )
    $p = $Doc.Paragraphs.Add()
    $p.Range.Text = $Text
    $p.Range.ListFormat.ApplyBulletDefault()
    $p.Format.SpaceAfter = 4
    $p.Format.LineSpacingRule = 5
    $p.Range.Font.Name = "Calibri"
    $p.Range.Font.Size = 11
}

function Add-Numbered {
    param(
        [Parameter(Mandatory=$true)] $Doc,
        [Parameter(Mandatory=$true)] [string]$Text
    )
    $p = $Doc.Paragraphs.Add()
    $p.Range.Text = $Text
    $p.Range.ListFormat.ApplyNumberDefault()
    $p.Format.SpaceAfter = 4
    $p.Format.LineSpacingRule = 5
    $p.Range.Font.Name = "Calibri"
    $p.Range.Font.Size = 11
}

function Add-Heading {
    param(
        [Parameter(Mandatory=$true)] $Doc,
        [Parameter(Mandatory=$true)] [string]$Text,
        [int]$Level = 1
    )
    $style = if ($Level -eq 1) { "Heading 1" } elseif ($Level -eq 2) { "Heading 2" } else { "Heading 3" }
    $size = if ($Level -eq 1) { 16 } elseif ($Level -eq 2) { 13 } else { 12 }
    [void](Add-Paragraph -Doc $Doc -Text $Text -Style $style -FontName "Calibri" -FontSize $size -Bold $true -SpaceAfter 6 -SpaceBefore 10)
}

function Set-CellText {
    param(
        [Parameter(Mandatory=$true)] $Cell,
        [Parameter(Mandatory=$true)] [string]$Text,
        [bool]$Bold = $false
    )
    $Cell.Range.Text = $Text
    $Cell.Range.Font.Name = "Calibri"
    $Cell.Range.Font.Size = 10
    $Cell.Range.Font.Bold = $Bold
    $Cell.Range.ParagraphFormat.SpaceAfter = 0
    $Cell.Range.ParagraphFormat.SpaceBefore = 0
    $Cell.VerticalAlignment = 1
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
    $doc = $word.Documents.Add()

    $section = $doc.Sections.Item(1)
    $section.PageSetup.TopMargin = 72
    $section.PageSetup.BottomMargin = 72
    $section.PageSetup.LeftMargin = 72
    $section.PageSetup.RightMargin = 72
    $section.PageSetup.HeaderDistance = 35
    $section.PageSetup.FooterDistance = 35

    $title = Add-Paragraph -Doc $doc -Text "Digital Com Project Report" -Style "Normal" -FontName "Calibri" -FontSize 24 -Bold $true -SpaceAfter 3 -SpaceBefore 0 -PassThru
    $title.Alignment = 1

    $subtitle = Add-Paragraph -Doc $doc -Text "Digital commerce system for smart homes, villas, city assets, commercial spaces, and yachts" -Style "Normal" -FontName "Calibri" -FontSize 11 -Bold $false -SpaceAfter 12 -PassThru
    $subtitle.Alignment = 1

    Add-Paragraph -Doc $doc -Text "Prepared from the project description supplied by the client. This report defines the product vision, scope, user workflow, delivery phases, and technical direction for the Digital Com platform." -SpaceAfter 12

    Add-Heading -Doc $doc -Text "1. Project Overview" -Level 1
    Add-Paragraph -Doc $doc -Text "Digital Com is a digital commerce and smart-environment platform for the smart homes industry. It allows a customer to create an account, upload or scan a property blueprint, and generate a 3D digital twin in Unity before any physical implementation begins. The system is intended to support residential villas, commercial spaces, smart city assets, yachts, and specialized environments through a common commerce and visualization layer."
    Add-Paragraph -Doc $doc -Text "The core idea is to move the selling, planning, and commissioning process into a simulation-first workflow. Instead of purchasing smart-home or smart-building services blindly, the customer can see the property model, place devices and services inside it, test interactions, review the effect of each service, and decide what to deploy. The platform then becomes a bridge between commercial selection, digital design, implementation, and later live synchronization with the real installation."

    Add-Heading -Doc $doc -Text "2. Business Problem and Project Objective" -Level 1
    Add-Paragraph -Doc $doc -Text "The smart-home and automation market often suffers from disconnected planning tools, unclear purchasing decisions, weak visualization of the final result, and limited support after installation. Customers usually need to imagine the end result from a drawing or from a technical bill of materials. That makes sales slower, support harder, and upgrades less transparent."
    Add-Paragraph -Doc $doc -Text "Digital Com addresses that gap by combining e-commerce, digital twin generation, AI-assisted modeling, augmented reality, simulation, billing, customer support, and post-deployment system monitoring in one platform. The objective is to reduce friction for the customer, improve the accuracy of system design, accelerate sales, and provide a long-term operational layer for maintenance and debugging."

    Add-Heading -Doc $doc -Text "3. Scope of Services" -Level 1
    Add-Paragraph -Doc $doc -Text "The project is based on the service groups shown in the original concept description. These groups define the solution catalog that the customer can browse and insert into a property model."

    $table = $doc.Tables.Add($doc.Range($doc.Content.End - 1, $doc.Content.End - 1), 6, 2)
    $table.Borders.Enable = 1
    $table.Range.ParagraphFormat.SpaceAfter = 0
    $table.Range.ParagraphFormat.SpaceBefore = 0
    $table.Columns.Item(1).Width = 120
    $table.Columns.Item(2).Width = 348

    Set-CellText -Cell $table.Cell(1,1) -Text "Vertical" -Bold $true
    Set-CellText -Cell $table.Cell(1,2) -Text "Included Services" -Bold $true
    Set-CellText -Cell $table.Cell(2,1) -Text "Smart City"
    Set-CellText -Cell $table.Cell(2,2) -Text "Traffic management, waste management systems, water monitoring and control, smart street lighting"
    Set-CellText -Cell $table.Cell(3,1) -Text "Commercial"
    Set-CellText -Cell $table.Cell(3,2) -Text "Automated power management, smart lighting, smart digital signage, integrated surveillance systems, interactive access control"
    Set-CellText -Cell $table.Cell(4,1) -Text "Residential"
    Set-CellText -Cell $table.Cell(4,2) -Text "IP 4K video distribution, smart energy meters, smart AC and heating, IP multi-zone audio, lighting and shading systems, home security, geo-fencing"
    Set-CellText -Cell $table.Cell(5,1) -Text "Yacht"
    Set-CellText -Cell $table.Cell(5,2) -Text "Hidden motorized TV systems, monitored cloud services, marine audio, marine communication"
    Set-CellText -Cell $table.Cell(6,1) -Text "Other"
    Set-CellText -Cell $table.Cell(6,2) -Text "Guest room management, digitized nursing systems, lighting management systems, automated pro audio"

    Add-Heading -Doc $doc -Text "4. Platform Concept" -Level 1
    Add-Paragraph -Doc $doc -Text "Digital Com functions as a digital commerce platform with a built-in digital twin. The commercial layer is not separated from the technical layer. The customer can discover services, configure them in the model, test the result, and then proceed toward implementation or subscription-based live synchronization."
    Add-Paragraph -Doc $doc -Text "The platform is designed around Crestron and KNX-based automation ecosystems, so the product catalog and downstream implementation logic can align with professional smart-building hardware and control standards."

    Add-Heading -Doc $doc -Text "5. User Workflow" -Level 1
    Add-Numbered -Doc $doc -Text "The user signs up and creates a personal account."
    Add-Numbered -Doc $doc -Text "The user uploads or scans the blueprint of a villa, house, city property, commercial site, or yacht."
    Add-Numbered -Doc $doc -Text "The system converts the blueprint into a layered 3D model using Unity-based processing."
    Add-Numbered -Doc $doc -Text "The user can optionally upload a photo taken through AR so the platform can place real objects and contextual details inside the model."
    Add-Numbered -Doc $doc -Text "The user browses the available services from the catalog and inserts them into the digital twin."
    Add-Numbered -Doc $doc -Text "The user moves through the model in 2D or 3D, similar to a game, to test how the selected features behave."
    Add-Numbered -Doc $doc -Text "The user may activate VR access depending on budget and sponsorship model."
    Add-Numbered -Doc $doc -Text "The user can interact with the assistant chatbot or create support conversations with the IT team."
    Add-Numbered -Doc $doc -Text "The user can pay deposits and maintenance fees through the platform."
    Add-Numbered -Doc $doc -Text "After implementation, the same digital twin can be reused for live monitoring, version tracking, support, and future upgrades."

    Add-Heading -Doc $doc -Text "6. Digital Twin and 3D Model Generation" -Level 1
    Add-Paragraph -Doc $doc -Text "A key feature of the platform is the conversion of blueprint files into a navigable model. The system accepts PDF, JPG, PNG, or similar property-source inputs and converts them into a structured digital twin. Each floor can be represented as a separate layer so that the user can inspect the architecture independently or view the building as a complete model."
    Add-Paragraph -Doc $doc -Text "Unity is used as the game-engine environment for rendering, interaction, movement, object placement, and simulation. The model is not only a static viewer. It is intended to behave as an interactive environment where the user can walk through rooms, test service layouts, evaluate experience flow, and understand how the property will function after deployment."
    Add-Paragraph -Doc $doc -Text "The AR scanning step adds another layer of realism. Instead of relying only on a flat blueprint, the user can scan or photograph the physical home and insert real-world objects or visual references into the model. This helps the digital twin reflect the intended final environment more accurately."

    Add-Heading -Doc $doc -Text "7. Interaction, Visualization, and Control" -Level 1
    Add-Paragraph -Doc $doc -Text "The platform supports both a 2D planning view and a 3D immersive view. The 2D view is useful for quick layout checks, service placement, and floor-by-floor comparison. The 3D view is used to move through rooms, inspect device positions, and test the appearance and operation of the configured solution."
    Add-Paragraph -Doc $doc -Text "The interface is intended to mirror the real smart-home interface as closely as possible. Tablets and mobile devices can be used inside the simulation model with the same style of controls that the real installation will expose later. That makes the digital twin a practical preview of the final system, not just a marketing model."
    Add-Paragraph -Doc $doc -Text "Lightmaps and heatmaps can be used to visualize physical properties of the space, such as lighting behavior, energy usage, comfort zones, or traffic patterns through the property. These overlays help the customer and the support team understand how the environment behaves under different scenarios."

    Add-Heading -Doc $doc -Text "8. VR Support" -Level 1
    Add-Paragraph -Doc $doc -Text "VR support is optional and budget-dependent. The project can be delivered in two possible ways. The first option is a lower-cost setup where a controller connects to a mobile phone and is used together with a PS5 controller or similar input device. The second option is a full VR set, such as a complete headset and controller package, for a more immersive and premium experience."
    Add-Paragraph -Doc $doc -Text "VR is valuable when the customer needs to validate room scale, interior placement, visual comfort, and the overall feeling of the space before implementation."

    Add-Heading -Doc $doc -Text "9. Commerce, Billing, and Customer Support" -Level 1
    Add-Paragraph -Doc $doc -Text "Digital Com includes a billing system that can collect deposits and maintenance fees. This is important because the platform is not only a visualization tool. It is also the transactional entry point for the project lifecycle."
    Add-Paragraph -Doc $doc -Text "An AI chatbot assists the user during browsing and configuration. The same chat layer can be used to open direct communication with the support or IT department. This creates a faster escalation path when the user needs help selecting services, checking compatibility, or troubleshooting a deployment issue."

    Add-Heading -Doc $doc -Text "10. Phase 1 and Phase 2 Delivery Model" -Level 1
    Add-Paragraph -Doc $doc -Text "The project is best understood in two phases."

    $phaseTable = $doc.Tables.Add($doc.Range($doc.Content.End - 1, $doc.Content.End - 1), 3, 2)
    $phaseTable.Borders.Enable = 1
    $phaseTable.Columns.Item(1).Width = 120
    $phaseTable.Columns.Item(2).Width = 348
    Set-CellText -Cell $phaseTable.Cell(1,1) -Text "Phase" -Bold $true
    Set-CellText -Cell $phaseTable.Cell(1,2) -Text "Description" -Bold $true
    Set-CellText -Cell $phaseTable.Cell(2,1) -Text "Phase 1"
    Set-CellText -Cell $phaseTable.Cell(2,2) -Text "Digital commerce and digital twin creation before implementation. The user uploads the blueprint, configures services, tests the model, pays required fees, and reviews the expected outcome."
    Set-CellText -Cell $phaseTable.Cell(3,1) -Text "Phase 2"
    Set-CellText -Cell $phaseTable.Cell(3,2) -Text "Live synchronization with the real installation after deployment. The saved blueprint and model are reused so the platform can mirror the actual property state and support long-term maintenance."

    Add-Paragraph -Doc $doc -Text "In phase 2, the platform uses a gateway server that validates whether the user has an active subscription. Only valid users are forwarded to the central server. Periodic traffic is collected, and when the user opens the app the server returns the latest state data so the digital twin can reflect the real environment."
    Add-Paragraph -Doc $doc -Text "This phase turns the platform into a live operational layer for home, yacht, or commercial assets. The twin becomes a continuously updated representation of the deployed system rather than a one-time sales model."

    Add-Heading -Doc $doc -Text "11. Remote Support and Debugging Model" -Level 1
    Add-Paragraph -Doc $doc -Text "A major operational benefit of the project is remote debugging. If a support issue requires deeper intervention, a small tunneling server can be used to create a secure support path. That server remains closed by default and is opened only with the proper encryption key and controlled user approval."
    Add-Paragraph -Doc $doc -Text "The user can expose a port or a controlled access point to support, and the support engineer connects as a client to the Crestron-side user environment to apply the fix. This approach reduces service time and avoids unnecessary physical intervention when a remote fix is possible."
    Add-Paragraph -Doc $doc -Text "The architecture also assumes message-broker-based updates from the smart-home processor into a database. Engineers can inspect the home state, review event history, and inject approved code updates or configuration changes when the workflow requires it."

    Add-Heading -Doc $doc -Text "12. Version Tracking and Data Retention" -Level 1
    Add-Paragraph -Doc $doc -Text "The system stores blueprint versions and digital twin versions so the customer can rebuild the same environment later if required. This is important for upgrades, rework, multi-phase construction, and long-term support."
    Add-Paragraph -Doc $doc -Text "Version tracking also helps the support team understand what changed between deployments, which layout was active at a given time, and which service set was selected in each phase of the project."

    Add-Heading -Doc $doc -Text "13. Security and Access Control" -Level 1
    Add-Paragraph -Doc $doc -Text "Because the platform interacts with live automation environments, security has to be treated as a core feature. Account access, subscription validation, gateway filtering, encrypted tunneling, and role-based support access should all be enforced."
    Add-Paragraph -Doc $doc -Text "The goal is to ensure that only authenticated customers and authorized support staff can interact with the live environment, while all other traffic remains blocked or read-only."

    Add-Heading -Doc $doc -Text "14. Expected Benefits" -Level 1
    Add-Bullet -Doc $doc -Text "Better sales conversion because customers can see the final result before buying."
    Add-Bullet -Doc $doc -Text "Lower design ambiguity because services are tested in a realistic digital twin."
    Add-Bullet -Doc $doc -Text "Faster commissioning and clearer handover between sales, engineering, and support."
    Add-Bullet -Doc $doc -Text "Improved maintenance because the live system can be monitored and versioned."
    Add-Bullet -Doc $doc -Text "Stronger customer confidence through AR, VR, and interactive simulations."
    Add-Bullet -Doc $doc -Text "A reusable platform that can serve smart homes, villas, yachts, city assets, and commercial projects."

    Add-Heading -Doc $doc -Text "15. Conclusion" -Level 1
    Add-Paragraph -Doc $doc -Text "Digital Com is a smart-home commerce and digital twin platform that combines property scanning, AI-assisted model generation, Unity-based simulation, optional VR, customer billing, chatbot support, and post-deployment synchronization. Its main value is that it connects the commercial purchase process to the real technical implementation in one continuous workflow."
    Add-Paragraph -Doc $doc -Text "The project is strong because it solves a real market problem: customers want to understand what they are buying before the hardware is installed. By giving them a live, interactive, versioned digital twin, Digital Com makes the sales process clearer, the support process faster, and the final smart environment easier to manage."

    $outputFull = Join-Path (Get-Location) $OutputPath
    if (Test-Path $outputFull) { Remove-Item $outputFull -Force }
    $doc.SaveAs2($outputFull, 16)
    $doc.Close()
}
finally {
    $word.Quit()
}
