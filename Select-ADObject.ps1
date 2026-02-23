function Select-ADObject {
    <#
    .SYNOPSIS
        Displays an interactive GUI dialog to select an Active Directory object from the forest hierarchy.

    .DESCRIPTION
        Select-ADObject provides a graphical user interface (GUI) for browsing and selecting Active Directory 
        objects such as Organizational Units (OUs), users, groups, and computers. The function displays a 
        tree view of the forest structure with custom icons for different object types. It supports browsing 
        across multiple domains and optionally highlights Tier 0 (high-security) objects.

    .PARAMETER Title
        This text displayed in the dialog window title. Default is "Select AD Object".

    .PARAMETER LocalDomainOnly
        If specified, only the current domain and its contents are displayed. By default, all domains 
        in the forest are shown.

    .PARAMETER DomainSelectionOnly
        If specified, only domain objects can be selected. The tree view is not expanded to show child objects.

    .PARAMETER IncludeUsers
        If specified, user objects are included in the browsable tree view.

    .PARAMETER IncludeGroups
        If specified, group objects are included in the browsable tree view.

    .PARAMETER IncludeComputers
        If specified, computer objects are included in the browsable tree view.

    .PARAMETER MarkTier0
        If specified, Tier 0 (high-security) objects are marked with a red overlay on their icons. 
        This helps identify highly sensitive AD objects.

    .PARAMETER Tier0Regex
        A regular expression pattern used to identify Tier 0 objects. Default pattern matches:
        - CN=Administrator
        - OU=Domain Controllers
        - OU=Tier 0 (case-insensitive)
        - OU=T0 (case-insensitive)

    .OUTPUTS
        System.String
        Returns the distinguished name (DN) of the selected Active Directory object, or $null if the 
        user cancels the dialog.

    .EXAMPLE
        PS> Select-ADObject -Title "Select Target OU"
        Description: Displays the AD selection dialog with a custom title title. Only OUs and containers 
        are available for selection.

    .EXAMPLE
        PS> Select-ADObject -LocalDomainOnly -IncludeUsers -IncludeGroups
        Description: Shows only the current domain's contents, allowing selection of users and groups 
        in addition to OUs.

    .EXAMPLE
        PS> Select-ADObject -DomainSelectionOnly
        Description: Restricts the selection to domain objects only.

    .EXAMPLE
        PS> Select-ADObject -MarkTier0 -IncludeUsers -Title "Select Tier 0 Target"
        Description: Displays all users and OUs, highlighting Tier 0 objects with a red indicator.

    .NOTES
        - Author:      zimmermann.holger@live.de
        - Version:     1.0
        - Last Update: 2026-02-22

        - This is a private function and requires a graphical environment (Windows Forms).
        - The function uses Active Directory modules to query the forest structure.
        - All logging is handled through Write-Log.
    #>

    [CmdletBinding()]
    param(
        [string]$Title = "Select AD Object",
        [switch]$LocalDomainOnly,
        [switch]$DomainSelectionOnly,
        [switch]$IncludeUsers,
        [switch]$IncludeGroups,
        [switch]$IncludeComputers,
        [switch]$MarkTier0,
        [Switch]$ShowAdvanced,
        [string]$Tier0Regex = '(?i)(^CN=Administrator,|^OU=Domain Controllers,|^OU=Tier\s*0|^OU=T0)'
    )

    #################### main code | out- host #####################

    # =========================
    # Assemblies
    # =========================
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.DirectoryServices
    [System.Windows.Forms.Application]::EnableVisualStyles()

    # =========================
    # Modern Icon Creation
    # =========================

    function New-RoundedRectPath($X, $Y, $W, $H, $R) {
        $path = New-Object System.Drawing.Drawing2D.GraphicsPath
        $path.AddArc($X, $Y, $R, $R, 180, 90)
        $path.AddArc($X + $W - $R, $Y, $R, $R, 270, 90)
        $path.AddArc($X + $W - $R, $Y + $H - $R, $R, $R, 0, 90)
        $path.AddArc($X, $Y + $H - $R, $R, $R, 90, 90)
        $path.CloseFigure()
        $path
    }

    function New-IconBitmap($Kind) {
        $bmp = New-Object System.Drawing.Bitmap 16, 16
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.SmoothingMode = 'AntiAlias'
        $g.Clear([System.Drawing.Color]::Transparent)

        $stroke = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(60, 60, 60), 1)
        $fill = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(220, 226, 232))

        switch ($Kind) {
            'Forest' {

    # Top node (Forest root)
    $g.FillEllipse($fill, 6, 2, 4, 4)
    $g.DrawEllipse($stroke, 6, 2, 4, 4)

    # Left child (Domain)
    $g.FillEllipse($fill, 3, 8, 4, 4)
    $g.DrawEllipse($stroke, 3, 8, 4, 4)

    # Right child (Domain)
    $g.FillEllipse($fill, 9, 8, 4, 4)
    $g.DrawEllipse($stroke, 9, 8, 4, 4)

    # Connection lines
    $g.DrawLine($stroke, 8, 6, 5, 8)
    $g.DrawLine($stroke, 8, 6, 11, 8)
}
            'Forest1' {
                $g.DrawEllipse($stroke, 2, 2, 12, 12)
                $g.DrawLine($stroke, 2, 8, 14, 8)
                $g.DrawLine($stroke, 8, 2, 8, 14)
                $g.DrawEllipse($stroke, 5, 5, 6, 6)
            }		
            'Domain' {
                $points = [System.Drawing.Point[]]@(
                    (New-Object System.Drawing.Point 8, 3),
                    (New-Object System.Drawing.Point 13, 12),
                    (New-Object System.Drawing.Point 3, 12)
                )

                $g.FillPolygon($fill, $points)
                $g.DrawPolygon($stroke, $points)
                $g.DrawLine($stroke, 5, 9, 11, 9)
            }
            'OU' {
                $p = New-RoundedRectPath 2 5 12 8 3
                $g.FillPath($fill, $p)
                $g.DrawPath($stroke, $p)
            }
            'Container' {
                $p = New-RoundedRectPath 3 4 10 10 3
                $g.FillPath($fill, $p)
                $g.DrawPath($stroke, $p)
                $g.DrawLine($stroke, 4, 8, 12, 8)
            }
            'User' {
                $g.FillEllipse($fill, 5, 3, 6, 6)
                $g.DrawEllipse($stroke, 5, 3, 6, 6)

                $p = New-RoundedRectPath 3 9 10 5 3
                $g.FillPath($fill, $p)
                $g.DrawPath($stroke, $p)
            }
            'Computer' {
                $p = New-RoundedRectPath 2 3 12 8 2
                $g.FillPath($fill, $p)
                $g.DrawPath($stroke, $p)

                $g.DrawLine($stroke, 6, 12, 10, 12)
                $g.DrawLine($stroke, 8, 11, 8, 14)
            }
            'Group' {
                $g.FillEllipse($fill, 4, 4, 4, 4)
                $g.DrawEllipse($stroke, 4, 4, 4, 4)

                $g.FillEllipse($fill, 8, 4, 4, 4)
                $g.DrawEllipse($stroke, 8, 4, 4, 4)

                $p = New-RoundedRectPath 4 9 8 4 3
                $g.FillPath($fill, $p)
                $g.DrawPath($stroke, $p)
            }
        }
        $g.Dispose()
        $bmp
    }

    function Add-TierOverlay($bmp) {
        $clone = $bmp.Clone()
        $g = [System.Drawing.Graphics]::FromImage($clone)
        $g.FillEllipse([System.Drawing.Brushes]::Red, 10, 10, 5, 5)
        $g.Dispose()
        $clone
    }

    # =========================
    # ImageList
    # =========================

    $imageList = New-Object System.Windows.Forms.ImageList
    $imageList.ImageSize = '16,16'

    $imageList.Images.Add("Forest", (New-IconBitmap Forest)) | Out-Null
    $imageList.Images.Add("Domain", (New-IconBitmap Domain)) | Out-Null
    $imageList.Images.Add("OU", (New-IconBitmap OU)) | Out-Null
    $imageList.Images.Add("Container", (New-IconBitmap Container)) | Out-Null
    $imageList.Images.Add("User", (New-IconBitmap User)) | Out-Null
    $imageList.Images.Add("Computer", (New-IconBitmap Computer)) | Out-Null
    $imageList.Images.Add("Group", (New-IconBitmap Group)) | Out-Null

    if ($MarkTier0) {
        foreach ($k in @("User", "OU", "Container", "Computer", "Group")) {
            $imageList.Images.Add("${k}_T0", (Add-TierOverlay $imageList.Images[$k])) | Out-Null
        }
    }

    # =========================
    # Helpers
    # =========================

    function Convert-DomainToDN($d) { 'DC=' + ($d -replace '\.', ',DC=') }

    function Get-ForestDomains {
        ([System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()).Domains |
        Select-Object -ExpandProperty Name
    }

    function Get-ChildNodes($node) {

        if ($DomainSelectionOnly) { return }

        if ($node.Nodes.Count -eq 1 -and $node.Nodes[0].Text -eq 'dummy') {

            $node.Nodes.Clear()

            $root = [ADSI]"LDAP://$($node.Name)"
            $searcher = New-Object System.DirectoryServices.DirectorySearcher($root)
            $searcher.SearchScope = 'OneLevel'
            
            $filters = @(
                "(objectClass=organizationalUnit)",
                "(objectClass=container)"
                "(objectClass=builtinDomain)"
            )

            if ($IncludeUsers) { $filters += "(objectCategory=user)" }
            if ($IncludeGroups) { $filters += "(objectCategory=group)" }
            if ($IncludeComputers) { $filters += "(objectCategory=computer)" }

            $orFilter = "(|" + ($filters -join "") + ")"

            if (-not $ShowAdvanced) {
                $searcher.Filter = "(& $orFilter (!(showInAdvancedViewOnly=TRUE)))"
            }
            else {
                $searcher.Filter = $orFilter
            }
          
            write-verbose $orFilter
		  
            $null = $searcher.PropertiesToLoad.AddRange(@("name", "distinguishedName", "objectClass"))

            foreach ($res in $searcher.FindAll()) {

                $dn = $res.Properties.distinguishedname[0]
                $name = $res.Properties.name[0]
                $cls = $res.Properties.objectclass[-1]
                # change icon here
                #$iconKey = if ($cls -eq "organizationalUnit") { "OU" } else { "User" }

                $iconKey = switch ($cls) {
                    'organizationalUnit' { 'OU' }
                    'container' { 'Container' }
                    'user' { 'User' }
                    'group' { 'Group' }
                    'computer' { 'Computer' }
                    default { 'Container' }
                }

                if ($MarkTier0 -and $dn -match $Tier0Regex -and $imageList.Images.ContainsKey("${iconKey}_T0")) {
                    $iconKey = "${iconKey}_T0"
                }

                $child = New-Object System.Windows.Forms.TreeNode
                $child.Text = $name
                $child.Name = $dn
                $child.ImageKey = $iconKey
                $child.SelectedImageKey = $iconKey
                $child.Nodes.Add("dummy") | Out-Null

                $node.Nodes.Add($child) | Out-Null
            }
        }
    }

    # =========================
    # FORM
    # =========================

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = '520,700'
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.Padding = '10,10,10,10'

    $font = New-Object System.Drawing.Font("Segoe UI", 9)

    $tree = New-Object System.Windows.Forms.TreeView
    $tree.Location = '10,10'
    $tree.Size = '480,560'
    $tree.Font = $font
    $tree.ImageList = $imageList

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Location = '10,580'
    $statusLabel.Size = '480,20'
    $statusLabel.Text = "Selected: none"
    $statusLabel.Font = $font

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = '310,610'
    $btnCancel.Size = '80,28'
    $btnCancel.Font = $font

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.Location = '410,610'
    $btnOK.Size = '80,28'
    $btnOK.Font = $font
    $btnOK.Enabled = $false

    $form.AcceptButton = $btnOK
    $form.CancelButton = $btnCancel

    # =========================
    # Build Forest
    # =========================

    $forestName = ([System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()).Name

    $forestNode = New-Object System.Windows.Forms.TreeNode
    $forestNode.Text = $forestName
    $forestNode.ImageKey = "Forest"
    $forestNode.SelectedImageKey = "Forest"
    $tree.Nodes.Add($forestNode) | Out-Null

    $domains = Get-ForestDomains
    $currentDomain = ([System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()).Name

    foreach ($d in $domains) {

        if ($LocalDomainOnly -and $d -ne $currentDomain) { continue }

        $dn = Convert-DomainToDN $d

        $domainNode = New-Object System.Windows.Forms.TreeNode
        $domainNode.Text = $d
        $domainNode.Name = $dn
        $domainNode.ImageKey = "Domain"
        $domainNode.SelectedImageKey = "Domain"
        $domainNode.Nodes.Add("dummy") | Out-Null

        $forestNode.Nodes.Add($domainNode) | Out-Null
    }

    $forestNode.Expand()

    # =========================
    # Events
    # =========================

    $tree.add_BeforeExpand({ Get-ChildNodes $_.Node })

    $tree.add_AfterSelect({
            $node = $_.Node
            $statusLabel.Text = "Selected: $($node.Text)"

            if ($DomainSelectionOnly) {
                $btnOK.Enabled = ($node.Parent -and -not $node.Parent.Parent)
                return
            }

            if ($AllowDomainSwitch) {
                $btnOK.Enabled = [bool]($node.Parent)
            }
            else {
                $btnOK.Enabled = [bool]($node.Parent -and $node.Parent.Parent)
            }
        })

    $btnOK.add_Click({
            $form.Tag = $tree.SelectedNode.Name
            $form.Close()
        })

    $btnCancel.add_Click({
            $form.Tag = $null
            $form.Close()
        })

    $form.Controls.AddRange(@($tree, $statusLabel, $btnCancel, $btnOK))

    $null = $form.ShowDialog()

    return $form.Tag

}
