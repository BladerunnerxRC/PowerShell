[CmdletBinding()]
param(
	[switch]$NoUI
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-AppBasePath {
	if ($PSScriptRoot) {
		return $PSScriptRoot
	}

	if ($PSCommandPath) {
		return (Split-Path -Path $PSCommandPath -Parent)
	}

	return (Get-Location).Path
}

function Get-DefaultTaxOptions {
	return @(
		[PSCustomObject]@{ Name = 'No Tax'; Rate = 0.00 }
		[PSCustomObject]@{ Name = 'WA Seattle (10.35%)'; Rate = 10.35 }
		[PSCustomObject]@{ Name = 'CA Los Angeles (9.50%)'; Rate = 9.50 }
		[PSCustomObject]@{ Name = 'TX Austin (8.25%)'; Rate = 8.25 }
	)
}

function Resolve-TaxOptions {
	param(
		[AllowNull()]
		[object[]]$TaxOptions
	)

	$normalized = New-Object System.Collections.Generic.List[object]

	if ($null -ne $TaxOptions) {
		foreach ($option in $TaxOptions) {
			if ($null -eq $option) {
				continue
			}

			$name = [string]$option.Name
			$name = $name.Trim()
			if ([string]::IsNullOrWhiteSpace($name)) {
				continue
			}

			$rate = 0.0
			if (-not [double]::TryParse([string]$option.Rate, [ref]$rate)) {
				continue
			}

			if ($rate -lt 0.0 -or $rate -gt 100.0) {
				continue
			}

			$normalized.Add([PSCustomObject]@{
					Name = $name
					Rate = [Math]::Round([decimal]$rate, 4)
				})
		}
	}

	if ($normalized.Count -eq 0) {
		foreach ($defaultOption in (Get-DefaultTaxOptions)) {
			$normalized.Add($defaultOption)
		}
	}

	return $normalized.ToArray()
}

function Resolve-DefaultTaxName {
	param(
		[Parameter(Mandatory)]
		[object[]]$TaxOptions,

		[AllowNull()]
		[string]$DefaultTaxName
	)

	$options = @(Resolve-TaxOptions -TaxOptions $TaxOptions)
	if ($options.Count -eq 0) {
		return ''
	}

	if (-not [string]::IsNullOrWhiteSpace($DefaultTaxName)) {
		foreach ($option in $options) {
			if ($option.Name.Equals($DefaultTaxName, [System.StringComparison]::OrdinalIgnoreCase)) {
				return [string]$option.Name
			}
		}
	}

	return [string]$options[0].Name
}

function Resolve-ThemeMode {
	param(
		[AllowNull()]
		[string]$ThemeMode
	)

	if ($ThemeMode -and $ThemeMode.Equals('Dark', [System.StringComparison]::OrdinalIgnoreCase)) {
		return 'Dark'
	}

	return 'Light'
}

function Resolve-ShowSplash {
	param(
		[AllowNull()]
		[object]$ShowSplash
	)

	if ($null -eq $ShowSplash) {
		return $true
	}

	if ($ShowSplash -is [bool]) {
		return [bool]$ShowSplash
	}

	$parsedShowSplash = $true
	if ([bool]::TryParse([string]$ShowSplash, [ref]$parsedShowSplash)) {
		return $parsedShowSplash
	}

	return $true
}

function Get-ThemePalette {
	param(
		[Parameter(Mandatory)]
		[string]$ThemeMode
	)

	$resolvedThemeMode = Resolve-ThemeMode -ThemeMode $ThemeMode
	if ($resolvedThemeMode -eq 'Dark') {
		return @{
			FormBack = [System.Drawing.Color]::FromArgb(32, 32, 36)
			InputBack = [System.Drawing.Color]::FromArgb(46, 46, 52)
			ButtonBack = [System.Drawing.Color]::FromArgb(58, 58, 66)
			Text = [System.Drawing.Color]::FromArgb(235, 235, 235)
			Border = [System.Drawing.Color]::FromArgb(88, 88, 96)
			GrandTotal = [System.Drawing.Color]::FromArgb(120, 230, 140)
		}
	}

	return @{
		FormBack = [System.Drawing.Color]::FromArgb(240, 240, 240)
		InputBack = [System.Drawing.Color]::White
		ButtonBack = [System.Drawing.Color]::FromArgb(245, 245, 245)
		Text = [System.Drawing.Color]::Black
		Border = [System.Drawing.Color]::FromArgb(200, 200, 200)
		GrandTotal = [System.Drawing.Color]::DarkGreen
	}
}

function Set-ThemeOnControl {
	param(
		[Parameter(Mandatory)]
		[object]$Control,

		[Parameter(Mandatory)]
		[hashtable]$Palette,

		[switch]$PreserveStatusColor
	)

	if ($null -eq $Control) {
		return
	}

	if ($Control -is [System.Windows.Forms.Form] -or $Control -is [System.Windows.Forms.GroupBox]) {
		$Control.BackColor = $Palette.FormBack
		$Control.ForeColor = $Palette.Text
	}
	elseif ($Control -is [System.Windows.Forms.TextBox] -or $Control -is [System.Windows.Forms.ComboBox] -or $Control -is [System.Windows.Forms.ListBox]) {
		$Control.BackColor = $Palette.InputBack
		$Control.ForeColor = $Palette.Text
	}
	elseif ($Control -is [System.Windows.Forms.Button]) {
		$Control.BackColor = $Palette.ButtonBack
		$Control.ForeColor = $Palette.Text
		$Control.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
		$Control.FlatAppearance.BorderColor = $Palette.Border
	}
	elseif ($Control -is [System.Windows.Forms.Label]) {
		$Control.BackColor = [System.Drawing.Color]::Transparent
		if ($Control.Name -eq 'valueGrandTotal') {
			$Control.ForeColor = $Palette.GrandTotal
		}
		elseif (-not ($PreserveStatusColor -and ($Control.Name -eq 'statusMain' -or $Control.Name -eq 'statusOptions'))) {
			$Control.ForeColor = $Palette.Text
		}
	}
	elseif ($Control -is [System.Windows.Forms.CheckBox]) {
		$Control.BackColor = [System.Drawing.Color]::Transparent
		$Control.ForeColor = $Palette.Text
	}

	if ($Control.Controls -and $Control.Controls.Count -gt 0) {
		foreach ($child in $Control.Controls) {
			Set-ThemeOnControl -Control $child -Palette $Palette -PreserveStatusColor:$PreserveStatusColor
		}
	}
}

function Save-TaxOptions {
	param(
		[Parameter(Mandatory)]
		[string]$ConfigPath,

		[Parameter(Mandatory)]
		[object[]]$TaxOptions,

		[AllowNull()]
		[string]$DefaultTaxName,

		[AllowNull()]
		[string]$ThemeMode,

		[AllowNull()]
		[object]$ShowSplash
	)

	$directory = Split-Path -Path $ConfigPath -Parent
	if (-not (Test-Path -LiteralPath $directory)) {
		New-Item -Path $directory -ItemType Directory -Force | Out-Null
	}

	$normalized = @(Resolve-TaxOptions -TaxOptions $TaxOptions)
	$resolvedDefaultTaxName = Resolve-DefaultTaxName -TaxOptions $normalized -DefaultTaxName $DefaultTaxName
	$resolvedThemeMode = Resolve-ThemeMode -ThemeMode $ThemeMode
	$resolvedShowSplash = Resolve-ShowSplash -ShowSplash $ShowSplash

	[PSCustomObject]@{
		DefaultTaxName = $resolvedDefaultTaxName
		ThemeMode = $resolvedThemeMode
		ShowSplash = $resolvedShowSplash
		TaxOptions = $normalized
	} |
		ConvertTo-Json -Depth 4 |
		Set-Content -LiteralPath $ConfigPath -Encoding UTF8
}

function Import-TaxOptions {
	param(
		[Parameter(Mandatory)]
		[string]$ConfigPath
	)

	if (-not (Test-Path -LiteralPath $ConfigPath)) {
		$defaults = @(Resolve-TaxOptions -TaxOptions (Get-DefaultTaxOptions))
		$defaultTaxName = Resolve-DefaultTaxName -TaxOptions $defaults -DefaultTaxName ''
		$themeMode = Resolve-ThemeMode -ThemeMode 'Light'
		$showSplash = Resolve-ShowSplash -ShowSplash $true
		Save-TaxOptions -ConfigPath $ConfigPath -TaxOptions $defaults -DefaultTaxName $defaultTaxName -ThemeMode $themeMode -ShowSplash $showSplash
		return [PSCustomObject]@{
			TaxOptions = $defaults
			DefaultTaxName = $defaultTaxName
			ThemeMode = $themeMode
			ShowSplash = $showSplash
		}
	}

	try {
		$raw = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8
		if ([string]::IsNullOrWhiteSpace($raw)) {
			$defaults = @(Resolve-TaxOptions -TaxOptions (Get-DefaultTaxOptions))
			$defaultTaxName = Resolve-DefaultTaxName -TaxOptions $defaults -DefaultTaxName ''
			$themeMode = Resolve-ThemeMode -ThemeMode 'Light'
			$showSplash = Resolve-ShowSplash -ShowSplash $true
			Save-TaxOptions -ConfigPath $ConfigPath -TaxOptions $defaults -DefaultTaxName $defaultTaxName -ThemeMode $themeMode -ShowSplash $showSplash
			return [PSCustomObject]@{
				TaxOptions = $defaults
				DefaultTaxName = $defaultTaxName
				ThemeMode = $themeMode
				ShowSplash = $showSplash
			}
		}

		$loaded = $raw | ConvertFrom-Json

		$loadedOptions = @()
		$loadedDefaultTaxName = ''
		$loadedThemeMode = 'Light'
		$loadedShowSplash = $true
		if ($loaded.PSObject.Properties['TaxOptions']) {
			$loadedOptions = @($loaded.TaxOptions)
			if ($loaded.PSObject.Properties['DefaultTaxName']) {
				$loadedDefaultTaxName = [string]$loaded.DefaultTaxName
			}
			if ($loaded.PSObject.Properties['ThemeMode']) {
				$loadedThemeMode = [string]$loaded.ThemeMode
			}
			if ($loaded.PSObject.Properties['ShowSplash']) {
				$loadedShowSplash = Resolve-ShowSplash -ShowSplash $loaded.ShowSplash
			}
		}
		else {
			# Backward compatibility for legacy config that stored only an array.
			$loadedOptions = @($loaded)
		}

		$normalized = @(Resolve-TaxOptions -TaxOptions $loadedOptions)
		$resolvedDefaultTaxName = Resolve-DefaultTaxName -TaxOptions $normalized -DefaultTaxName $loadedDefaultTaxName
		$resolvedThemeMode = Resolve-ThemeMode -ThemeMode $loadedThemeMode
		$resolvedShowSplash = Resolve-ShowSplash -ShowSplash $loadedShowSplash
		Save-TaxOptions -ConfigPath $ConfigPath -TaxOptions $normalized -DefaultTaxName $resolvedDefaultTaxName -ThemeMode $resolvedThemeMode -ShowSplash $resolvedShowSplash

		return [PSCustomObject]@{
			TaxOptions = $normalized
			DefaultTaxName = $resolvedDefaultTaxName
			ThemeMode = $resolvedThemeMode
			ShowSplash = $resolvedShowSplash
		}
	}
	catch {
		$defaults = @(Resolve-TaxOptions -TaxOptions (Get-DefaultTaxOptions))
		$defaultTaxName = Resolve-DefaultTaxName -TaxOptions $defaults -DefaultTaxName ''
		$themeMode = Resolve-ThemeMode -ThemeMode 'Light'
		$showSplash = Resolve-ShowSplash -ShowSplash $true
		Save-TaxOptions -ConfigPath $ConfigPath -TaxOptions $defaults -DefaultTaxName $defaultTaxName -ThemeMode $themeMode -ShowSplash $showSplash
		return [PSCustomObject]@{
			TaxOptions = $defaults
			DefaultTaxName = $defaultTaxName
			ThemeMode = $themeMode
			ShowSplash = $showSplash
		}
	}
}

function Get-FreightFee {
	param(
		[Parameter(Mandatory)]
		[decimal]$BidAmount
	)

	if ($BidAmount -gt [decimal]5.00) {
		return [decimal]1.00
	}

	return [decimal]0.25
}

function Get-AuctionTotals {
	param(
		[Parameter(Mandatory)]
		[decimal]$BidAmount,

		[Parameter(Mandatory)]
		[decimal]$TaxRatePercent
	)

	$buyerPremium = [Math]::Round($BidAmount * [decimal]0.1725, 2)
	$freightFee = Get-FreightFee -BidAmount $BidAmount
	$taxableSubtotal = [Math]::Round($BidAmount + $buyerPremium + $freightFee, 2)
	$taxAmount = [Math]::Round($taxableSubtotal * ($TaxRatePercent / [decimal]100.0), 2)
	$grandTotal = [Math]::Round($taxableSubtotal + $taxAmount, 2)

	return [PSCustomObject]@{
		BidAmount       = [Math]::Round($BidAmount, 2)
		BuyerPremium    = $buyerPremium
		FreightFee      = $freightFee
		TaxableSubtotal = $taxableSubtotal
		TaxRatePercent  = [Math]::Round($TaxRatePercent, 4)
		TaxAmount       = $taxAmount
		GrandTotal      = $grandTotal
	}
}

function Format-Currency {
	param(
		[Parameter(Mandatory)]
		[decimal]$Value
	)

	return ('$' + $Value.ToString('N2'))
}

function Format-Percent {
	param(
		[Parameter(Mandatory)]
		[decimal]$Value
	)

	return ($Value.ToString('N2') + '%')
}

function Get-TaxOptionDisplay {
	param(
		[Parameter(Mandatory)]
		[object]$Option,

		[switch]$IsDefault
	)

	$display = ('{0} - {1}' -f $Option.Name, (Format-Percent -Value ([decimal]$Option.Rate)))
	if ($IsDefault) {
		$display += ' (Default)'
	}

	return $display
}

function Get-TaxOptionIndex {
	param(
		[Parameter(Mandatory)]
		[object[]]$TaxOptions,

		[AllowNull()]
		[string]$TaxName
	)

	$options = @(Resolve-TaxOptions -TaxOptions $TaxOptions)
	if ($options.Count -eq 0) {
		return -1
	}

	if (-not [string]::IsNullOrWhiteSpace($TaxName)) {
		for ($i = 0; $i -lt $options.Count; $i++) {
			if ($options[$i].Name.Equals($TaxName, [System.StringComparison]::OrdinalIgnoreCase)) {
				return $i
			}
		}
	}

	return 0
}

function Restart-AppInstance {
	param(
		[Parameter(Mandatory)]
		[string]$ScriptPath
	)

	if ([string]::IsNullOrWhiteSpace($ScriptPath)) {
		return $false
	}

	try {
		Start-Process -FilePath 'powershell.exe' -ArgumentList @(
			'-NoProfile',
			'-STA',
			'-ExecutionPolicy',
			'Bypass',
			'-File',
			$ScriptPath
		) | Out-Null

		return $true
	}
	catch {
		return $false
	}
}

function Show-AboutMessage {
	[void][System.Windows.Forms.MessageBox]::Show(
		'This application created at the behest of T.A.Smith',
		'About',
		[System.Windows.Forms.MessageBoxButtons]::OK,
		[System.Windows.Forms.MessageBoxIcon]::Information
	)
}

function Show-StartupSplash {
	param(
		[Parameter(Mandatory)]
		[string]$ThemeMode
	)

	$palette = Get-ThemePalette -ThemeMode $ThemeMode

	$splash = New-Object System.Windows.Forms.Form
	$splash.StartPosition = 'CenterScreen'
	$splash.FormBorderStyle = 'None'
	$splash.ShowInTaskbar = $false
	$splash.TopMost = $true
	$splash.Size = New-Object System.Drawing.Size(520, 160)
	$splash.BackColor = $palette.FormBack

	$label = New-Object System.Windows.Forms.Label
	$label.AutoSize = $false
	$label.Dock = [System.Windows.Forms.DockStyle]::Fill
	$label.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
	$label.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 12)
	$label.ForeColor = $palette.Text
	$label.Text = 'This application created at the behest of T.A.Smith'

	$splash.Controls.Add($label)

	$timer = New-Object System.Windows.Forms.Timer
	$timer.Interval = 2000
	$timer.Add_Tick({
		$timer.Stop()
		$splash.Close()
	})

	$timer.Start()
	[void]$splash.ShowDialog()
	$timer.Dispose()
	$splash.Dispose()
}

function Show-TaxOptionsDialog {
	param(
		[Parameter(Mandatory)]
		[System.Windows.Forms.Form]$Owner,

		[Parameter(Mandatory)]
		[string]$ConfigPath,

		[Parameter(Mandatory)]
		[object[]]$TaxOptions,

		[AllowNull()]
		[string]$CurrentDefaultTaxName,

		[AllowNull()]
		[string]$CurrentThemeMode,

		[AllowNull()]
		[object]$CurrentShowSplash
	)

	$workingOptions = New-Object System.Collections.Generic.List[object]
	foreach ($option in (Resolve-TaxOptions -TaxOptions $TaxOptions)) {
		$workingOptions.Add([PSCustomObject]@{ Name = [string]$option.Name; Rate = [decimal]$option.Rate })
	}
	$defaultTaxName = Resolve-DefaultTaxName -TaxOptions $workingOptions.ToArray() -DefaultTaxName $CurrentDefaultTaxName
	$themeMode = Resolve-ThemeMode -ThemeMode $CurrentThemeMode
	$showSplash = Resolve-ShowSplash -ShowSplash $CurrentShowSplash
	$initialThemeMode = $themeMode

	$dialog = New-Object System.Windows.Forms.Form
	$dialog.Text = 'Tax Options'
	$dialog.StartPosition = 'CenterParent'
	$dialog.FormBorderStyle = 'FixedDialog'
	$dialog.MaximizeBox = $false
	$dialog.MinimizeBox = $false
	$dialog.ClientSize = New-Object System.Drawing.Size(500, 560)

	$listBox = New-Object System.Windows.Forms.ListBox
	$listBox.Location = New-Object System.Drawing.Point(16, 16)
	$listBox.Size = New-Object System.Drawing.Size(468, 185)

	$labelName = New-Object System.Windows.Forms.Label
	$labelName.Text = 'State + Local label:'
	$labelName.Location = New-Object System.Drawing.Point(16, 220)
	$labelName.AutoSize = $true

	$textName = New-Object System.Windows.Forms.TextBox
	$textName.Location = New-Object System.Drawing.Point(16, 242)
	$textName.Size = New-Object System.Drawing.Size(468, 25)

	$labelRate = New-Object System.Windows.Forms.Label
	$labelRate.Text = 'Combined tax percentage:'
	$labelRate.Location = New-Object System.Drawing.Point(16, 276)
	$labelRate.AutoSize = $true

	$textRate = New-Object System.Windows.Forms.TextBox
	$textRate.Location = New-Object System.Drawing.Point(16, 298)
	$textRate.Size = New-Object System.Drawing.Size(180, 25)

	$buttonAdd = New-Object System.Windows.Forms.Button
	$buttonAdd.Text = 'Add Option'
	$buttonAdd.Location = New-Object System.Drawing.Point(214, 296)
	$buttonAdd.Size = New-Object System.Drawing.Size(130, 30)

	$buttonDelete = New-Object System.Windows.Forms.Button
	$buttonDelete.Text = 'Delete Selected'
	$buttonDelete.Location = New-Object System.Drawing.Point(354, 296)
	$buttonDelete.Size = New-Object System.Drawing.Size(130, 30)

	$buttonSetDefault = New-Object System.Windows.Forms.Button
	$buttonSetDefault.Text = 'Set Selected As Default'
	$buttonSetDefault.Location = New-Object System.Drawing.Point(16, 332)
	$buttonSetDefault.Size = New-Object System.Drawing.Size(190, 30)

	$labelDefault = New-Object System.Windows.Forms.Label
	$labelDefault.Location = New-Object System.Drawing.Point(214, 338)
	$labelDefault.Size = New-Object System.Drawing.Size(270, 22)
	$labelDefault.AutoEllipsis = $true

	$labelTheme = New-Object System.Windows.Forms.Label
	$labelTheme.Text = 'Theme mode:'
	$labelTheme.Location = New-Object System.Drawing.Point(16, 370)
	$labelTheme.AutoSize = $true

	$comboTheme = New-Object System.Windows.Forms.ComboBox
	$comboTheme.Location = New-Object System.Drawing.Point(16, 392)
	$comboTheme.Size = New-Object System.Drawing.Size(190, 25)
	$comboTheme.DropDownStyle = 'DropDownList'
	[void]$comboTheme.Items.Add('Light')
	[void]$comboTheme.Items.Add('Dark')
	$comboTheme.SelectedItem = $themeMode

	$checkShowSplash = New-Object System.Windows.Forms.CheckBox
	$checkShowSplash.Text = 'Show splash message at startup'
	$checkShowSplash.Location = New-Object System.Drawing.Point(16, 425)
	$checkShowSplash.Size = New-Object System.Drawing.Size(280, 24)
	$checkShowSplash.Checked = $showSplash

	$statusLabel = New-Object System.Windows.Forms.Label
	$statusLabel.Name = 'statusOptions'
	$statusLabel.Location = New-Object System.Drawing.Point(16, 460)
	$statusLabel.Size = New-Object System.Drawing.Size(468, 22)
	$statusLabel.ForeColor = [System.Drawing.Color]::Firebrick

	$buttonDone = New-Object System.Windows.Forms.Button
	$buttonDone.Text = 'Done'
	$buttonDone.Location = New-Object System.Drawing.Point(354, 494)
	$buttonDone.Size = New-Object System.Drawing.Size(130, 32)

	$dialogState = [PSCustomObject]@{ DidChange = $false; AnyChange = $false; RestartRequested = $false }

	$refreshList = {
		$listBox.Items.Clear()
		foreach ($entry in $workingOptions) {
			$isDefault = $entry.Name.Equals($defaultTaxName, [System.StringComparison]::OrdinalIgnoreCase)
			[void]$listBox.Items.Add((Get-TaxOptionDisplay -Option $entry -IsDefault:$isDefault))
		}

		$labelDefault.Text = ('Current default: ' + $defaultTaxName)
	}

	$validateRate = {
		param(
			[string]$RateText,
			[ref]$OutRate
		)

		$parsedRate = [decimal]0
		$ok = [decimal]::TryParse($RateText, [ref]$parsedRate)
		if (-not $ok) {
			return $false
		}

		if ($parsedRate -lt 0 -or $parsedRate -gt 100) {
			return $false
		}

		$OutRate.Value = [Math]::Round($parsedRate, 4)
		return $true
	}

	$buttonAdd.Add_Click({
			$statusLabel.Text = ''
			$statusLabel.ForeColor = [System.Drawing.Color]::Firebrick
			$name = $textName.Text.Trim()
			if ([string]::IsNullOrWhiteSpace($name)) {
				$statusLabel.Text = 'Enter a name for this tax option.'
				return
			}

			$rate = [decimal]0
			if (-not (& $validateRate -RateText $textRate.Text.Trim() -OutRate ([ref]$rate))) {
				$statusLabel.Text = 'Tax percent must be a numeric value between 0 and 100.'
				return
			}

			foreach ($entry in $workingOptions) {
				if ($entry.Name.Equals($name, [System.StringComparison]::OrdinalIgnoreCase)) {
					$statusLabel.Text = 'That name already exists. Use a unique label.'
					return
				}
			}

			$workingOptions.Add([PSCustomObject]@{ Name = $name; Rate = $rate })
			if ([string]::IsNullOrWhiteSpace($defaultTaxName)) {
				$defaultTaxName = $name
			}
			$dialogState.DidChange = $true
			$dialogState.AnyChange = $true
			& $refreshList

			$textName.Clear()
			$textRate.Clear()
			$statusLabel.ForeColor = [System.Drawing.Color]::ForestGreen
			$statusLabel.Text = 'Tax option added.'
		})

	$buttonDelete.Add_Click({
			$statusLabel.ForeColor = [System.Drawing.Color]::Firebrick
			$statusLabel.Text = ''

			if ($listBox.SelectedIndex -lt 0) {
				$statusLabel.Text = 'Select a tax option to delete.'
				return
			}

			if ($workingOptions.Count -le 1) {
				$statusLabel.Text = 'At least one tax option must remain.'
				return
			}

			$index = $listBox.SelectedIndex
			$removedName = [string]$workingOptions[$index].Name
			$workingOptions.RemoveAt($index)
			if ($removedName.Equals($defaultTaxName, [System.StringComparison]::OrdinalIgnoreCase)) {
				$defaultTaxName = Resolve-DefaultTaxName -TaxOptions $workingOptions.ToArray() -DefaultTaxName ''
			}
			$dialogState.DidChange = $true
			$dialogState.AnyChange = $true
			& $refreshList
			$statusLabel.ForeColor = [System.Drawing.Color]::ForestGreen
			$statusLabel.Text = 'Tax option deleted.'
		})

	$buttonSetDefault.Add_Click({
			$statusLabel.ForeColor = [System.Drawing.Color]::Firebrick
			$statusLabel.Text = ''

			if ($listBox.SelectedIndex -lt 0) {
				$statusLabel.Text = 'Select a tax option to set as default.'
				return
			}

			$defaultTaxName = [string]$workingOptions[$listBox.SelectedIndex].Name

			try {
				$finalOptions = @(Resolve-TaxOptions -TaxOptions $workingOptions.ToArray())
				$resolvedDefaultTaxName = Resolve-DefaultTaxName -TaxOptions $finalOptions -DefaultTaxName $defaultTaxName
				Save-TaxOptions -ConfigPath $ConfigPath -TaxOptions $finalOptions -DefaultTaxName $resolvedDefaultTaxName -ThemeMode $themeMode -ShowSplash $showSplash
				$defaultTaxName = $resolvedDefaultTaxName
				$dialogState.DidChange = $false
				$dialogState.AnyChange = $true
				& $refreshList
				$statusLabel.ForeColor = [System.Drawing.Color]::ForestGreen
				$statusLabel.Text = 'Default tax option saved immediately.'
			}
			catch {
				$statusLabel.ForeColor = [System.Drawing.Color]::Firebrick
				$statusLabel.Text = ('Could not save default option: ' + $_.Exception.Message)
			}
		})

	$comboTheme.Add_SelectedIndexChanged({
			if ($comboTheme.SelectedIndex -lt 0) {
				return
			}

			$themeMode = [string]$comboTheme.SelectedItem

			try {
				$finalOptions = @(Resolve-TaxOptions -TaxOptions $workingOptions.ToArray())
				$resolvedDefaultTaxName = Resolve-DefaultTaxName -TaxOptions $finalOptions -DefaultTaxName $defaultTaxName
				$resolvedThemeMode = Resolve-ThemeMode -ThemeMode $themeMode
				Save-TaxOptions -ConfigPath $ConfigPath -TaxOptions $finalOptions -DefaultTaxName $resolvedDefaultTaxName -ThemeMode $resolvedThemeMode -ShowSplash $showSplash
				$defaultTaxName = $resolvedDefaultTaxName
				$themeMode = $resolvedThemeMode
				$dialogState.DidChange = $false
				$dialogState.AnyChange = $true

				$palette = Get-ThemePalette -ThemeMode $themeMode
				Set-ThemeOnControl -Control $dialog -Palette $palette -PreserveStatusColor

				$statusLabel.ForeColor = [System.Drawing.Color]::ForestGreen
				$statusLabel.Text = 'Theme mode saved.'

				$applyNow = [System.Windows.Forms.MessageBox]::Show(
					'Theme saved. Do you want to apply the theme now? Restart required.',
					'Apply Theme',
					[System.Windows.Forms.MessageBoxButtons]::YesNo,
					[System.Windows.Forms.MessageBoxIcon]::Question
				)

				if ($applyNow -eq [System.Windows.Forms.DialogResult]::Yes) {
					$dialogState.RestartRequested = $true
					$dialog.DialogResult = [System.Windows.Forms.DialogResult]::OK
					$dialog.Close()
				}
			}
			catch {
				$dialogState.DidChange = $true
				$dialogState.AnyChange = $true
				$statusLabel.ForeColor = [System.Drawing.Color]::Firebrick
				$statusLabel.Text = ('Could not save theme mode: ' + $_.Exception.Message)
			}
		})

	$checkShowSplash.Add_CheckedChanged({
			$showSplash = $checkShowSplash.Checked
			$dialogState.DidChange = $true
			$dialogState.AnyChange = $true
			$statusLabel.ForeColor = [System.Drawing.Color]::ForestGreen
			$statusLabel.Text = 'Splash startup preference updated. Save with Done (or close Options).'
		})

	$saveChanges = {
		if (-not $dialogState.DidChange) {
			return
		}

		$finalOptions = @(Resolve-TaxOptions -TaxOptions $workingOptions.ToArray())
		$resolvedDefaultTaxName = Resolve-DefaultTaxName -TaxOptions $finalOptions -DefaultTaxName $defaultTaxName
		Save-TaxOptions -ConfigPath $ConfigPath -TaxOptions $finalOptions -DefaultTaxName $resolvedDefaultTaxName -ThemeMode $themeMode -ShowSplash $showSplash
		$defaultTaxName = $resolvedDefaultTaxName
		$dialogState.DidChange = $false
	}

	$buttonDone.Add_Click({
			& $saveChanges
			$dialog.DialogResult = [System.Windows.Forms.DialogResult]::OK
			$dialog.Close()
		})

	$dialog.Add_FormClosing({
			& $saveChanges
		})

	$dialog.Controls.AddRange(@(
			$listBox,
			$labelName,
			$textName,
			$labelRate,
			$textRate,
			$buttonAdd,
			$buttonDelete,
			$buttonSetDefault,
			$labelDefault,
			$labelTheme,
			$comboTheme,
			$checkShowSplash,
			$statusLabel,
			$buttonDone
		))

	& $refreshList
	$palette = Get-ThemePalette -ThemeMode $themeMode
	Set-ThemeOnControl -Control $dialog -Palette $palette -PreserveStatusColor
	$statusLabel.ForeColor = [System.Drawing.Color]::Firebrick

	[void]$dialog.ShowDialog($Owner)

	return [PSCustomObject]@{
		Changed = $dialogState.AnyChange
		Options = @(Resolve-TaxOptions -TaxOptions $workingOptions.ToArray())
		DefaultTaxName = $defaultTaxName
		ThemeMode = $themeMode
		ShowSplash = $showSplash
		ThemeChanged = ($themeMode -ne $initialThemeMode)
		RestartRequested = $dialogState.RestartRequested
	}
}

if ($NoUI) {
	Write-Output 'NoUI specified. Validation mode only.'
	return
}

if ([Threading.Thread]::CurrentThread.ApartmentState -ne [Threading.ApartmentState]::STA) {
	throw 'This application must run in STA mode. Start it with: powershell -STA -ExecutionPolicy Bypass -File .\\bidFTA_Bid_Price_Calculator'
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$appBasePath = Get-AppBasePath
$configPath = Join-Path -Path $appBasePath -ChildPath 'bidFTA_tax_options.json'
$taxConfig = Import-TaxOptions -ConfigPath $configPath
$taxOptions = @($taxConfig.TaxOptions)
$defaultTaxName = [string]$taxConfig.DefaultTaxName
$themeMode = Resolve-ThemeMode -ThemeMode ([string]$taxConfig.ThemeMode)
$showSplash = Resolve-ShowSplash -ShowSplash $taxConfig.ShowSplash

$form = New-Object System.Windows.Forms.Form
$form.Text = 'BidFTA Auction Total Calculator'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(700, 470)
$form.MinimumSize = New-Object System.Drawing.Size(700, 470)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.HelpButton = $true

$fontLabel = New-Object System.Drawing.Font('Segoe UI', 10)
$fontValue = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
$fontGrandValue = New-Object System.Drawing.Font('Segoe UI', 14, [System.Drawing.FontStyle]::Bold)
$fontGrandLabel = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)

$labelBid = New-Object System.Windows.Forms.Label
$labelBid.Text = 'Auction Bid Amount:'
$labelBid.Location = New-Object System.Drawing.Point(20, 20)
$labelBid.AutoSize = $true
$labelBid.Font = $fontLabel

$textBid = New-Object System.Windows.Forms.TextBox
$textBid.Location = New-Object System.Drawing.Point(20, 46)
$textBid.Size = New-Object System.Drawing.Size(220, 24)
$textBid.Font = $fontLabel

$labelTax = New-Object System.Windows.Forms.Label
$labelTax.Text = 'State + Local Tax Selection:'
$labelTax.Location = New-Object System.Drawing.Point(265, 20)
$labelTax.AutoSize = $true
$labelTax.Font = $fontLabel

$comboTax = New-Object System.Windows.Forms.ComboBox
$comboTax.Location = New-Object System.Drawing.Point(265, 46)
$comboTax.Size = New-Object System.Drawing.Size(280, 28)
$comboTax.DropDownStyle = 'DropDownList'
$comboTax.Font = $fontLabel

$buttonOptions = New-Object System.Windows.Forms.Button
$buttonOptions.Text = 'Options'
$buttonOptions.Location = New-Object System.Drawing.Point(560, 44)
$buttonOptions.Size = New-Object System.Drawing.Size(100, 30)
$buttonOptions.Font = $fontLabel

$statusMain = New-Object System.Windows.Forms.Label
$statusMain.Name = 'statusMain'
$statusMain.Location = New-Object System.Drawing.Point(20, 98)
$statusMain.Size = New-Object System.Drawing.Size(640, 24)
$statusMain.Font = $fontLabel
$statusMain.ForeColor = [System.Drawing.Color]::Firebrick

$groupResults = New-Object System.Windows.Forms.GroupBox
$groupResults.Text = 'Cost Breakdown'
$groupResults.Location = New-Object System.Drawing.Point(20, 140)
$groupResults.Size = New-Object System.Drawing.Size(640, 260)
$groupResults.Font = $fontLabel

$breakdownRows = @(
	'Winning Bid',
	"Buyer's Premium (17.25%)",
	'Freight Fee',
	'Taxable Subtotal',
	'Tax Amount',
	'Grand Total'
)

$valueLabels = @{}
$y = 35
foreach ($row in $breakdownRows) {
	$left = New-Object System.Windows.Forms.Label
	$left.Text = $row
	$left.Location = New-Object System.Drawing.Point(24, $y)
	$left.Size = New-Object System.Drawing.Size(260, 24)
	$left.Font = $fontLabel

	$right = New-Object System.Windows.Forms.Label
	$right.Text = '$0.00'
	$right.Location = New-Object System.Drawing.Point(310, $y)
	$right.Size = New-Object System.Drawing.Size(300, 24)
	$right.Font = $fontValue

	if ($row -eq 'Grand Total') {
		$left.Name = 'labelGrandTotal'
		$right.Name = 'valueGrandTotal'
		$left.Font = $fontGrandLabel
		$right.Font = $fontGrandValue
	}

	$groupResults.Controls.Add($left)
	$groupResults.Controls.Add($right)
	$valueLabels[$row] = $right
	$y += 34
}

$form.Controls.AddRange(@(
		$labelBid,
		$textBid,
		$labelTax,
		$comboTax,
		$buttonOptions,
		$statusMain,
		$groupResults
	))

$currentTaxOptions = @(Resolve-TaxOptions -TaxOptions $taxOptions)

$refreshTaxDropdown = {
	param(
		[int]$PreferredIndex = 0
	)

	$comboTax.Items.Clear()

	foreach ($option in $currentTaxOptions) {
		[void]$comboTax.Items.Add((Get-TaxOptionDisplay -Option $option))
	}

	if ($comboTax.Items.Count -gt 0) {
		if ($PreferredIndex -lt 0 -or $PreferredIndex -ge $comboTax.Items.Count) {
			$PreferredIndex = 0
		}

		$comboTax.SelectedIndex = $PreferredIndex
	}
}

$resetOutput = {
	foreach ($key in $valueLabels.Keys) {
		$valueLabels[$key].Text = '$0.00'
	}
}

$updateTotals = {
	$statusMain.Text = ''
	$statusMain.ForeColor = [System.Drawing.Color]::Firebrick

	$bidText = $textBid.Text.Trim()
	$bid = [decimal]0
	if ([string]::IsNullOrWhiteSpace($bidText)) {
		& $resetOutput
		return
	}

	if (-not [decimal]::TryParse($bidText, [ref]$bid)) {
		& $resetOutput
		$statusMain.Text = 'Bid must be a valid number.'
		return
	}

	if ($bid -lt 0) {
		& $resetOutput
		$statusMain.Text = 'Bid cannot be negative.'
		return
	}

	if ($comboTax.SelectedIndex -lt 0 -or $comboTax.SelectedIndex -ge @($currentTaxOptions).Count) {
		& $resetOutput
		$statusMain.Text = 'Select a tax option.'
		return
	}

	$selectedTax = @($currentTaxOptions)[$comboTax.SelectedIndex]
	$totals = Get-AuctionTotals -BidAmount $bid -TaxRatePercent ([decimal]$selectedTax.Rate)

	$valueLabels['Winning Bid'].Text = Format-Currency -Value $totals.BidAmount
	$valueLabels["Buyer's Premium (17.25%)"].Text = Format-Currency -Value $totals.BuyerPremium
	$valueLabels['Freight Fee'].Text = Format-Currency -Value $totals.FreightFee
	$valueLabels['Taxable Subtotal'].Text = Format-Currency -Value $totals.TaxableSubtotal
	$valueLabels['Tax Amount'].Text = ('{0} at {1}' -f (Format-Currency -Value $totals.TaxAmount), (Format-Percent -Value $totals.TaxRatePercent))
	$valueLabels['Grand Total'].Text = Format-Currency -Value $totals.GrandTotal
}

$textBid.Add_TextChanged({ & $updateTotals })
$comboTax.Add_SelectedIndexChanged({ & $updateTotals })

$form.Add_HelpButtonClicked({
		param($helpSender, $helpEventArgs)
		Show-AboutMessage
		$helpEventArgs.Cancel = $true
	})

$buttonOptions.Add_Click({
		$result = Show-TaxOptionsDialog -Owner $form -ConfigPath $configPath -TaxOptions $currentTaxOptions -CurrentDefaultTaxName $defaultTaxName -CurrentThemeMode $themeMode -CurrentShowSplash $showSplash

		$currentTaxOptions = @(Resolve-TaxOptions -TaxOptions $result.Options)
		$defaultTaxName = Resolve-DefaultTaxName -TaxOptions $currentTaxOptions -DefaultTaxName ([string]$result.DefaultTaxName)
		$themeMode = Resolve-ThemeMode -ThemeMode ([string]$result.ThemeMode)
		$showSplash = Resolve-ShowSplash -ShowSplash $result.ShowSplash
		$selectedIndex = Get-TaxOptionIndex -TaxOptions $currentTaxOptions -TaxName $defaultTaxName
		& $refreshTaxDropdown -PreferredIndex $selectedIndex

		if ($result.RestartRequested) {
			if (Restart-AppInstance -ScriptPath $PSCommandPath) {
				$form.Close()
				return
			}

			$statusMain.ForeColor = [System.Drawing.Color]::Firebrick
			$statusMain.Text = 'Theme saved, but restart could not be started automatically.'
		}

		if ($result.Changed) {
			if ($result.ThemeChanged) {
				$statusMain.ForeColor = [System.Drawing.Color]::ForestGreen
				$statusMain.Text = ('Theme saved (' + $themeMode + '). Restart app to apply.')
			}
			else {
				$statusMain.ForeColor = [System.Drawing.Color]::ForestGreen
				$statusMain.Text = 'Tax options updated.'
			}
		}

		& $updateTotals
	})

[System.Windows.Forms.Application]::add_ThreadException({
		param($threadSender, $threadEventArgs)
		[System.Windows.Forms.MessageBox]::Show(
			('Unexpected error: ' + $threadEventArgs.Exception.Message),
			'Application Error',
			[System.Windows.Forms.MessageBoxButtons]::OK,
			[System.Windows.Forms.MessageBoxIcon]::Error
		) | Out-Null
	})

$launchIndex = Get-TaxOptionIndex -TaxOptions $currentTaxOptions -TaxName $defaultTaxName
& $refreshTaxDropdown -PreferredIndex $launchIndex
$palette = Get-ThemePalette -ThemeMode $themeMode
Set-ThemeOnControl -Control $form -Palette $palette -PreserveStatusColor
& $updateTotals

if ($showSplash) {
	Show-StartupSplash -ThemeMode $themeMode
}

[void]$form.ShowDialog()
