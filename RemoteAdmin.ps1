Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Drawing

Import-Module -Name PSReadLine
Import-module WindowsCompatibility

# ICONE

$iconpath = "C:\temp\RemoteMonitoring\Icon\icone-dadministration.ico"
$icon = New-Object System.Drawing.Icon($iconpath)

###FENETRE###

$folderForm = New-Object System.Windows.Forms.Form
$folderform.text = "Remote Admin"
$folderForm.Size = New-Object Drawing.Point 1140,950
$folderForm.AutoScroll = $true
$largeurA = '1124'
$hauteurA = $folderform.height
$folderform.Icon = $icon

### KEYBOARD SHORTCUTS ###

$folderform.Add_KeyDown({

    param($sender, $e)

    if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::C) {

        $Connection.PerformClick()

    }

    if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::S) {

        $System.PerformClick()

    }

    if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::P) {

        $Processes.PerformClick()

    }

    if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {

        $Applications.PerformClick()

    }

    if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::J) {

        $Services.PerformClick()

    }

    if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::L) {

        $EventLogs.PerformClick()

    }

    if ($e.alt -and $e.KeyCode -eq [System.Windows.Forms.Keys]::F) {

        $folders.PerformClick()

    }

    if ($e.alt -and $e.KeyCode -eq [System.Windows.Forms.Keys]::C) {

        $CMD.PerformClick()

    }


})

$folderform.KeyPreview = $true

###LABEL###

# HEADERS --------------------------------------
$Title = New-object System.Windows.Forms.Label
$Title.text = 'REMOTE ADMIN'
$Title.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 18, [System.Drawing.FontStyle]::Bold)
$Title.BackColor = "#2D3250"
$Title.width = '600'
$Title.Height = '50'
$Title.ForeColor = 'white'
$Title.location = '430,30'
$folderForm.controls.Add($Title)
$Title.bringtofront()

$header = New-object System.Windows.Forms.Label
$header.BackColor = "#2D3250"
$header.ForeColor = "white"
$header.location = '0,0'
$header.Width = $largeurA
$header.Height = "90"
$folderForm.controls.Add($header)
$Header.sendtoback()

$lightColor = [System.Drawing.ColorTranslator]::FromHtml('#9290C3')

$header.add_Paint({

    param($sender, $e)
    $graphics = $e.Graphics
    $rect = $sender.ClientRectangle

    # Couleurs pour l'effet de relief
    $penLight = New-Object System.Drawing.Pen($lightColor)
    $darkColor = [System.Drawing.Color]::black

    # Dessiner le relief
    $penLight = New-Object System.Drawing.Pen($lightColor)
    $penDark = New-Object System.Drawing.Pen($darkColor)

    # Dessiner les côtés clairs
    $graphics.DrawLine($penLight, 0, 0, $rect.Width - 1, 0) # Haut
    $graphics.DrawLine($penLight, 0, 0, 0, $rect.Height - 1) # Gauche

    # Dessiner les côtés sombres
    $graphics.DrawLine($penDark, 0, $rect.Height - 1, $rect.Width - 1, $rect.Height - 1) # Bas
    $graphics.DrawLine($penDark, $rect.Width - 1, 0, $rect.Width - 1, $rect.Height - 1) # Droite

    # Nettoyer
    $penLight.Dispose()
    $penDark.Dispose()

})

$header2 = New-object System.Windows.Forms.Label
$header2.BackColor = "#424769"
$header2.ForeColor = "white"
$header2.location = '0,90'
$header2.Width = $largeurA
$header2.Height = "50"
$folderForm.controls.Add($header2)
$header2.sendtoback()

$lightColor = [System.Drawing.ColorTranslator]::FromHtml('#9290C3')

$header2.add_Paint({

    param($sender, $e)
    $graphics = $e.Graphics
    $rect = $sender.ClientRectangle

    # Couleurs pour l'effet de relief
    $penLight = New-Object System.Drawing.Pen($lightColor)
    $darkColor = [System.Drawing.Color]::black

    # Dessiner le relief
    $penLight = New-Object System.Drawing.Pen($lightColor)
    $penDark = New-Object System.Drawing.Pen($darkColor)

    # Dessiner les côtés clairs
    $graphics.DrawLine($penLight, 0, 0, $rect.Width - 1, 0) # Haut
    $graphics.DrawLine($penLight, 0, 0, 0, $rect.Height - 1) # Gauche

    # Dessiner les côtés sombres
    $graphics.DrawLine($penDark, 0, $rect.Height - 1, $rect.Width - 1, $rect.Height - 1) # Bas
    $graphics.DrawLine($penDark, $rect.Width - 1, 0, $rect.Width - 1, $rect.Height - 1) # Droite

    # Nettoyer
    $penLight.Dispose()
    $penDark.Dispose()

})

$Header2Frame1 = New-object System.Windows.Forms.Label
$Header2Frame1.BackColor = "#424769"
$Header2Frame1.ForeColor = "white"
$Header2Frame1.location = '210,90'
$Header2Frame1.Height = "50"
$Header2Frame1.width = "100"
$folderForm.controls.Add($Header2Frame1)
$Header2Frame1.BringToFront()

$lightColor = [System.Drawing.ColorTranslator]::FromHtml('#9290C3')

$Header2Frame1.add_Paint({

    param($sender, $e)
    $graphics = $e.Graphics
    $rect = $sender.ClientRectangle

    # Couleurs pour l'effet de relief
    $penLight = New-Object System.Drawing.Pen($lightColor)
    $darkColor = [System.Drawing.Color]::black

    # Dessiner le relief
    $penLight = New-Object System.Drawing.Pen($lightColor)
    $penDark = New-Object System.Drawing.Pen($darkColor)

    # Dessiner les côtés clairs
    $graphics.DrawLine($penLight, 0, 0, $rect.Width - 1, 0) # Haut
    $graphics.DrawLine($penLight, 0, 0, 0, $rect.Height - 1) # Gauche

    # Dessiner les côtés sombres
    $graphics.DrawLine($penDark, 0, $rect.Height - 1, $rect.Width - 1, $rect.Height - 1) # Bas
    $graphics.DrawLine($penDark, $rect.Width - 1, 0, $rect.Width - 1, $rect.Height - 1) # Droite

    # Nettoyer
    $penLight.Dispose()
    $penDark.Dispose()

})

$Header2Frame2 = New-object System.Windows.Forms.Label
$Header2Frame2.BackColor = "#424769"
$Header2Frame2.ForeColor = "white"
$Header2Frame2.location = '310,90'
$Header2Frame2.Height = "50"
$Header2Frame2.width = "420"
$folderForm.controls.Add($Header2Frame2)
$Header2Frame2.BringToFront()

$lightColor = [System.Drawing.ColorTranslator]::FromHtml('#9290C3')

$Header2Frame2.add_Paint({

    param($sender, $e)
    $graphics = $e.Graphics
    $rect = $sender.ClientRectangle

    # Couleurs pour l'effet de relief
    $penLight = New-Object System.Drawing.Pen($lightColor)
    $darkColor = [System.Drawing.Color]::black

    # Dessiner le relief
    $penLight = New-Object System.Drawing.Pen($lightColor)
    $penDark = New-Object System.Drawing.Pen($darkColor)

    # Dessiner les côtés clairs
    $graphics.DrawLine($penLight, 0, 0, $rect.Width - 1, 0) # Haut
    $graphics.DrawLine($penLight, 0, 0, 0, $rect.Height - 1) # Gauche

    # Dessiner les côtés sombres
    $graphics.DrawLine($penDark, 0, $rect.Height - 1, $rect.Width - 1, $rect.Height - 1) # Bas
    $graphics.DrawLine($penDark, $rect.Width - 1, 0, $rect.Width - 1, $rect.Height - 1) # Droite

    # Nettoyer
    $penLight.Dispose()
    $penDark.Dispose()

})

$Header2Frame3 = New-object System.Windows.Forms.Label
$Header2Frame3.BackColor = "#424769"
$Header2Frame3.ForeColor = "white"
$Header2Frame3.location = '730,90'
$Header2Frame3.Height = "50"
$Header2Frame3.width = "180"
$folderForm.controls.Add($Header2Frame3)
$Header2Frame3.BringToFront()

$lightColor = [System.Drawing.ColorTranslator]::FromHtml('#9290C3')

$Header2Frame3.add_Paint({

    param($sender, $e)
    $graphics = $e.Graphics
    $rect = $sender.ClientRectangle

    # Couleurs pour l'effet de relief
    $penLight = New-Object System.Drawing.Pen($lightColor)
    $darkColor = [System.Drawing.Color]::black

    # Dessiner le relief
    $penLight = New-Object System.Drawing.Pen($lightColor)
    $penDark = New-Object System.Drawing.Pen($darkColor)

    # Dessiner les côtés clairs
    $graphics.DrawLine($penLight, 0, 0, $rect.Width - 1, 0) # Haut
    $graphics.DrawLine($penLight, 0, 0, 0, $rect.Height - 1) # Gauche

    # Dessiner les côtés sombres
    $graphics.DrawLine($penDark, 0, $rect.Height - 1, $rect.Width - 1, $rect.Height - 1) # Bas
    $graphics.DrawLine($penDark, $rect.Width - 1, 0, $rect.Width - 1, $rect.Height - 1) # Droite

    # Nettoyer
    $penLight.Dispose()
    $penDark.Dispose()

})

#$Header3Backcolor = "ececec"
$header3 = New-object System.Windows.Forms.Label
$header3.ForeColor = "white"
#$Header3.BackColor = "$Header3Backcolor"
$header3.location = '0,140'
$header3.Width = $largeurA
$header3.visible = $true
$header3.Height = "30"
$folderForm.controls.Add($header3)

$lightColor = [System.Drawing.ColorTranslator]::FromHtml('#9290C3')

$header3.add_Paint({

    param($sender, $e)
    $graphics = $e.Graphics
    $rect = $sender.ClientRectangle

    # Couleurs pour l'effet de relief
    $penLight = New-Object System.Drawing.Pen($lightColor)
    $darkColor = [System.Drawing.Color]::black

    # Dessiner le relief
    $penLight = New-Object System.Drawing.Pen($lightColor)
    $penDark = New-Object System.Drawing.Pen($darkColor)

    # Dessiner les côtés clairs
    $graphics.DrawLine($penLight, 0, 0, $rect.Width - 1, 0) # Haut
    $graphics.DrawLine($penLight, 0, 0, 0, $rect.Height - 1) # Gauche

    # Dessiner les côtés sombres
    $graphics.DrawLine($penDark, 0, $rect.Height - 1, $rect.Width - 1, $rect.Height - 1) # Bas
    $graphics.DrawLine($penDark, $rect.Width - 1, 0, $rect.Width - 1, $rect.Height - 1) # Droite

    # Nettoyer
    $penLight.Dispose()
    $penDark.Dispose()

})

$Header3Frame1 = New-object System.Windows.Forms.Label
$Header3Frame1.ForeColor = "white"
$Header3Frame1.BackColor = "$Header3Backcolor"
$Header3Frame1.location = '0,140'
$Header3Frame1.Height = "30"
$Header3Frame1.width = "310"
$folderForm.controls.Add($Header3Frame1)
$Header3Frame1.BringToFront()

$lightColor = [System.Drawing.ColorTranslator]::FromHtml('#dadada')

$Header3Frame1.add_Paint({

    param($sender, $e)
    $graphics = $e.Graphics
    $rect = $sender.ClientRectangle

    # Couleurs pour l'effet de relief
    $penLight = New-Object System.Drawing.Pen($lightColor)
    $darkColor = [System.Drawing.Color]::black

    # Dessiner le relief
    $penLight = New-Object System.Drawing.Pen($lightColor)
    $penDark = New-Object System.Drawing.Pen($darkColor)

    # Dessiner les côtés clairs
    $graphics.DrawLine($penLight, 0, 0, $rect.Width - 1, 0) # Haut
    $graphics.DrawLine($penLight, 0, 0, 0, $rect.Height - 1) # Gauche

    # Dessiner les côtés sombres
    $graphics.DrawLine($penDark, 0, $rect.Height - 1, $rect.Width - 1, $rect.Height - 1) # Bas
    $graphics.DrawLine($penDark, $rect.Width - 1, 0, $rect.Width - 1, $rect.Height - 1) # Droite

    # Nettoyer
    $penLight.Dispose()
    $penDark.Dispose()

})

$Header3Frame2 = New-object System.Windows.Forms.Label
$Header3Frame2.ForeColor = "white"
$Header3Frame2.BackColor = "$Header3Backcolor"
$Header3Frame2.location = '310,140'
$Header3Frame2.Height = "30"
$Header3Frame2.width = "420"
$folderForm.controls.Add($Header3Frame2)
$Header3Frame2.BringToFront()

$lightColor = [System.Drawing.ColorTranslator]::FromHtml('#dadada')

$Header3Frame2.add_Paint({

    param($sender, $e)
    $graphics = $e.Graphics
    $rect = $sender.ClientRectangle

    # Couleurs pour l'effet de relief
    $penLight = New-Object System.Drawing.Pen($lightColor)
    $darkColor = [System.Drawing.Color]::black

    # Dessiner le relief
    $penLight = New-Object System.Drawing.Pen($lightColor)
    $penDark = New-Object System.Drawing.Pen($darkColor)

    # Dessiner les côtés clairs
    $graphics.DrawLine($penLight, 0, 0, $rect.Width - 1, 0) # Haut
    $graphics.DrawLine($penLight, 0, 0, 0, $rect.Height - 1) # Gauche

    # Dessiner les côtés sombres
    $graphics.DrawLine($penDark, 0, $rect.Height - 1, $rect.Width - 1, $rect.Height - 1) # Bas
    $graphics.DrawLine($penDark, $rect.Width - 1, 0, $rect.Width - 1, $rect.Height - 1) # Droite

    # Nettoyer
    $penLight.Dispose()
    $penDark.Dispose()

})

$Header3Frame3 = New-object System.Windows.Forms.Label
$Header3Frame3.ForeColor = "white"
$Header3Frame3.BackColor = "$Header3Backcolor"
$Header3Frame3.location = '730,140'
$Header3Frame3.Height = "30"
$Header3Frame3.width = "180"
$folderForm.controls.Add($Header3Frame3)
$Header3Frame3.BringToFront()

$lightColor = [System.Drawing.ColorTranslator]::FromHtml('#dadada')

$Header3Frame3.add_Paint({

    param($sender, $e)
    $graphics = $e.Graphics
    $rect = $sender.ClientRectangle

    # Couleurs pour l'effet de relief
    $penLight = New-Object System.Drawing.Pen($lightColor)
    $darkColor = [System.Drawing.Color]::black

    # Dessiner le relief
    $penLight = New-Object System.Drawing.Pen($lightColor)
    $penDark = New-Object System.Drawing.Pen($darkColor)

    # Dessiner les côtés clairs
    $graphics.DrawLine($penLight, 0, 0, $rect.Width - 1, 0) # Haut
    $graphics.DrawLine($penLight, 0, 0, 0, $rect.Height - 1) # Gauche

    # Dessiner les côtés sombres
    $graphics.DrawLine($penDark, 0, $rect.Height - 1, $rect.Width - 1, $rect.Height - 1) # Bas
    $graphics.DrawLine($penDark, $rect.Width - 1, 0, $rect.Width - 1, $rect.Height - 1) # Droite

    # Nettoyer
    $penLight.Dispose()
    $penDark.Dispose()

})

#$Header3Backcolor = "ececec"
$header4 = New-object System.Windows.Forms.Label
$header4.ForeColor = "white"
#$header4.BackColor = "$header4Backcolor"
$header4.location = '0,169'
$header4.Width = $largeurA
$header4.visible = $false
$header4.Height = "30"
$folderForm.controls.Add($header4)

$lightColor = [System.Drawing.ColorTranslator]::FromHtml('#9290C3')

$header4.add_Paint({

    param($sender, $e)
    $graphics = $e.Graphics
    $rect = $sender.ClientRectangle

    # Couleurs pour l'effet de relief
    $penLight = New-Object System.Drawing.Pen($lightColor)
    $darkColor = [System.Drawing.Color]::black

    # Dessiner le relief
    $penLight = New-Object System.Drawing.Pen($lightColor)
    $penDark = New-Object System.Drawing.Pen($darkColor)

    # Dessiner les côtés clairs
    $graphics.DrawLine($penLight, 0, 0, $rect.Width - 1, 0) # Haut
    $graphics.DrawLine($penLight, 0, 0, 0, $rect.Height - 1) # Gauche

    # Dessiner les côtés sombres
    $graphics.DrawLine($penDark, 0, $rect.Height - 1, $rect.Width - 1, $rect.Height - 1) # Bas
    $graphics.DrawLine($penDark, $rect.Width - 1, 0, $rect.Width - 1, $rect.Height - 1) # Droite

    # Nettoyer
    $penLight.Dispose()
    $penDark.Dispose()

})

$IndcationsSystem = New-object System.Windows.Forms.Label
$IndcationsSystem.text = ''
$IndicationsSystem.font = 'Arial, 7'
$IndcationsSystem.Width = '310'
$IndcationsSystem.height = '15'
$IndcationsSystem.FlatStyle = 'standard'
$IndcationsSystem.ForeColor = 'black'
$IndcationsSystem.location = '15,177'
$folderForm.controls.Add($IndcationsSystem)
$IndcationsSystem.bringtofront()

# FRAMES --------------------------------------
$ControlsLabel = New-object System.Windows.Forms.Label
$ControlsLabel.text = 'Controls'
$ControlsLabel.width = '46'
$ControlsLabel.visible = $false
$ControlsLabel.Height = '12'
$ControlsLabel.ForeColor = 'black'
$ControlsLabel.location = '1012,243'
$folderForm.controls.Add($ControlsLabel)
$ControlsLabel.bringtofront()

$Frame = New-object System.Windows.Forms.Label
$Frame.ForeColor = "white"
$Frame.location = '975,250'
$Frame.Width = $largeurA
$Frame.Height = "130"
$Frame.Width = "120"
$frame.Visible = $false
$folderForm.controls.Add($Frame)

$frame.add_Paint({

    param ($sender, $e)
    $borderColor = [System.Drawing.Color]::black
    $borderWidth = 1 # Largeur du cadre
    $rectangle = [System.Drawing.Rectangle]::new(0, 0, $frame.Width - 1, $frame.Height - 1)
    $pen = [System.Drawing.Pen]::new($borderColor, $borderWidth)
    $e.Graphics.DrawRectangle($pen, $rectangle)
    $pen.Dispose()

})

$FrameBoutons = New-object System.Windows.Forms.Label
$FrameBoutons.ForeColor = "white"
$FrameBoutons.location = '975,400'
$FrameBoutons.Width = $largeurA
$FrameBoutons.Height = "130"
$FrameBoutons.Width = "120"
$FrameBoutons.Visible = $false
$folderForm.controls.Add($FrameBoutons)
$frameBoutons.BringToFront()

$frameBoutons.add_Paint({

    param ($sender, $e)
    $borderColor = [System.Drawing.Color]::black
    $borderWidth = 1 # Largeur du cadre
    $rectangle = [System.Drawing.Rectangle]::new(0, 0, $frame.Width - 1, $frame.Height - 1)
    $pen = [System.Drawing.Pen]::new($borderColor, $borderWidth)
    $e.Graphics.DrawRectangle($pen, $rectangle)
    $pen.Dispose()

})

$frame2 = New-object System.Windows.Forms.Label
$frame2.ForeColor = "white"
$frame2.location = '40,229'
$frame2.Width = $largeurA
$frame2.Height = "601"
$frame2.Width = "902"
$frame2.Visible = $false
$folderForm.controls.Add($frame2)

$frame2.add_Paint({

    param ($sender, $e)
    $borderColor = [System.Drawing.Color]::black
    $borderWidth = 1 # Largeur du cadre
    $rectangle = [System.Drawing.Rectangle]::new(0, 0, $frame2.Width - 1, $frame2.Height - 1)
    $pen = [System.Drawing.Pen]::new($borderColor, $borderWidth)
    $e.Graphics.DrawRectangle($pen, $rectangle)
    $pen.Dispose()

})

$frame3 = New-object System.Windows.Forms.Label
$frame3.ForeColor = "white"
$frame3.location = '40,835'
$frame3.Width = $largeurA
$frame3.Height = "60"
$frame3.Width = "902"
$frame3.Visible = $false
$folderForm.controls.Add($frame3)
$frame3.bringtofront()

$frame3.add_Paint({

    param ($sender, $e)
    $borderColor = [System.Drawing.Color]::black
    $borderWidth = 1 # Largeur du cadre
    $rectangle = [System.Drawing.Rectangle]::new(0, 0, $frame3.Width - 1, $frame3.Height - 1)
    $pen = [System.Drawing.Pen]::new($borderColor, $borderWidth)
    $e.Graphics.DrawRectangle($pen, $rectangle)
    $pen.Dispose()

})

# Labels
$ComputerNameLabel = New-object System.Windows.Forms.Label
$ComputerNameLabel.font = New-Object Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$ComputerNameLabel.text = 'Computer:'
$ComputerNameLabel.Width = '70'
$computerNameLabel.BackColor = '#424769'
$computerNameLabel.FlatStyle = 'standard'
$ComputerNameLabel.ForeColor = 'White'
$ComputerNameLabel.location = '5,107'
$folderForm.controls.Add($ComputerNameLabel)
$ComputerNameLabel.bringtofront()

$NodenameLabel = New-object System.Windows.Forms.Label
$NodenameLabel.text = 'PC:'
$NodenameLabel.font = 'Arial, 8'
$NodenameLabel.BackColor = "$Header3Backcolor"
$NodenameLabel.Height = '15'
$NodenameLabel.Width = '180'
$NodenameLabel.Visible = $true
$NodenameLabel.ForeColor = 'black'
$NodenameLabel.location = '5,148'
$folderForm.controls.Add($NodenameLabel)
$NodenameLabel.bringtofront()

$OwnerLabel = New-object System.Windows.Forms.Label
$OwnerLabel.text = 'User:'
$OwnerLabel.font = 'Arial, 8'
$OwnerLabel.Height = '15'
$OwnerLabel.Width = '100'
$OwnerLabel.Visible = $true
$OwnerLabel.ForeColor = 'black'
$OwnerLabel.BackColor = "$Header3Backcolor"
$OwnerLabel.location = '110,148'
$folderForm.controls.Add($OwnerLabel)
$OwnerLabel.bringtofront()

$StatusLabel = New-object System.Windows.Forms.Label
$StatusLabel.text = 'Status:'
$StatusLabel.Height = '15'
$StatusLabel.Width = '95'
$StatusLabel.Visible = $true
$StatusLabel.ForeColor = 'black'
$StatusLabel.BackColor = "$Header3Backcolor"
$StatusLabel.location = '205,148'
$folderForm.controls.Add($StatusLabel)
$StatusLabel.bringtofront()

$PingMSLabel = New-object System.Windows.Forms.Label
$PingMSLabel.text = 'Ping: ms'
$PingMSLabel.Height = '15'
$PingMSLabel.Width = '95'
$PingMSLabel.Visible = $true
$PingMSLabel.ForeColor = 'black'
$PingMSLabel.BackColor = "$Header3Backcolor"
$PingMSLabel.location = '325,148'
$folderForm.controls.Add($PingMSLabel)
$PingMSLabel.bringtofront()

$AdapterLabel = New-object System.Windows.Forms.Label
$AdapterLabel.text = "Adapter: $adapter"
$AdapterLabel.Height = '15'
$AdapterLabel.Width = '125'
$AdapterLabel.Visible = $true
$AdapterLabel.ForeColor = 'black'
$AdapterLabel.BackColor = "$Header3Backcolor"
$AdapterLabel.location = '395,149'
$folderForm.controls.Add($AdapterLabel)
$AdapterLabel.bringtofront()

$IPLabel = New-object System.Windows.Forms.Label
$IPLabel.text = "IP: $IP"
$IPLabel.Height = '15'
$IPLabel.Width = '95'
$IPLabel.Visible = $true
$IPLabel.ForeColor = 'black'
$IPLabel.BackColor = "$Header3Backcolor"
$IPLabel.location = '530,148'
$folderForm.controls.Add($IPLabel)
$IPLabel.bringtofront()

$Uptimelabel = New-object System.Windows.Forms.Label
$Uptimelabel.text = "Uptime:"
$Uptimelabel.Height = '15'
$Uptimelabel.Width = '90'
$Uptimelabel.Visible = $true
$Uptimelabel.ForeColor = 'black'
$Uptimelabel.BackColor = "$Header3Backcolor"
$Uptimelabel.location = '625,148'
$folderForm.controls.Add($Uptimelabel)
$Uptimelabel.bringtofront()

$Searchlabel = New-object System.Windows.Forms.Label
$Searchlabel.text = "Search:"
$Searchlabel.Height = '15'
$Searchlabel.Width = '90'
$Searchlabel.Visible = $true
$Searchlabel.ForeColor = 'black'
$Searchlabel.BackColor = "$Header3Backcolor"
$Searchlabel.location = '745,148'
$folderForm.controls.Add($Searchlabel)
$Searchlabel.bringtofront()

$Titre1 = New-object System.Windows.Forms.Label
$Titre1.font = New-Object Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$Titre1.text = ''
$Titre1.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 16, [System.Drawing.FontStyle]::Bold)
$Titre1.height = '20'
$Titre1.width = '200'
$Titre1.ForeColor = '#2D3250'
$Titre1.location = '35,200'
$folderForm.controls.Add($Titre1)
$Titre1.bringtofront()

$Titre2 = New-object System.Windows.Forms.Label
$Titre2.font = New-Object Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$Titre2.text = ''
$Titre2.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 16, [System.Drawing.FontStyle]::Bold)
$Titre2.autosize = $true
$Titre2.ForeColor = '#2D3250'
$Titre2.location = '35,330'
$folderForm.controls.Add($Titre2)
$Titre2.bringtofront()

$Titre3 = New-object System.Windows.Forms.Label
$Titre3.font = New-Object Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$Titre3.text = ''
$Titre3.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 16, [System.Drawing.FontStyle]::Bold)
$Titre3.autosize = $true
$Titre3.ForeColor = '#2D3250'
$Titre3.location = '35,460'
$folderForm.controls.Add($Titre3)
$Titre3.bringtofront()

$Titre4 = New-object System.Windows.Forms.Label
$Titre4.font = New-Object Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$Titre4.text = ''
$Titre4.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 16, [System.Drawing.FontStyle]::Bold)
$Titre4.autosize = $true
$Titre4.ForeColor = '#2D3250'
$Titre4.location = '35,590'
$folderForm.controls.Add($Titre4)
$Titre4.bringtofront()

$Titre5 = New-object System.Windows.Forms.Label
$Titre5.font = New-Object Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$Titre5.text = ''
$Titre5.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 16, [System.Drawing.FontStyle]::Bold)
$Titre5.autosize = $true
$Titre5.ForeColor = '#2D3250'
$Titre5.location = '35,720'
$folderForm.controls.Add($Titre5)
$Titre5.bringtofront()

$CountLabel = New-object System.Windows.Forms.Label
$CountLabel.text = ''
$CountLabel.font = 'Arial,8'
$CountLabel.autosize = $true
$CountLabel.ForeColor = '#2D3250'
$CountLabel.location = '42,850'
$folderForm.controls.Add($CountLabel)
$CountLabel.bringtofront()

$CountLabel2 = New-object System.Windows.Forms.Label
$CountLabel2.text = ''
$CountLabel2.font = 'Arial,8'
$CountLabel2.autosize = $true
$CountLabel2.ForeColor = '#2D3250'
$CountLabel2.location = '42,870'
$folderForm.controls.Add($CountLabel2)
$CountLabel2.bringtofront()

$LoadingLabel = New-object System.Windows.Forms.Label
$LoadingLabel.text = ''
$LoadingLabel.font = 'Arial, 9'
$LoadingLabel.autosize = $true
$LoadingLabel.ForeColor = '#2D3250'
$LoadingLabel.location = '1020,148'
$folderForm.controls.Add($LoadingLabel)
$LoadingLabel.bringtofront()

###TEXTBOXES###

$computerName = New-Object System.Windows.Forms.TextBox
$ComputerName.size = New-Object System.Drawing.Size(113,20)
$ComputerName.location = '80,105'
$folderform.controls.Add($ComputerName)
$computername.bringtofront()

$Search = New-Object System.Windows.Forms.TextBox
$Search.size = New-Object System.Drawing.Size(100,20)
$Search.location = '795,145'
$folderform.controls.Add($Search)
$Search.bringtofront()

### BOUTONS ###

$Connection = New-Object System.Windows.Forms.Button
$Connection.Text = 'Connection'
$Connection.Font = 'Arial, 7'
$Connection.width = 70
$Connection.height = 20
$Connection.FlatStyle = 'standard'
$Connection.BackColor = 'white'
$Connection.ForeColor = 'Black'
$Connection.Location = '225,105'
$InfoConnection = New-Object System.Windows.Forms.ToolTip
$Infoconnection.SetToolTip($Connection, 'Shortcut "Ctrl + C"')
$folderForm.Controls.Add($Connection)
$Connection.bringtofront()

# tableau system --------------------------------------
$SystemForDisplaying = New-Object System.Windows.Forms.Button
$SystemForDisplaying.Text = 'System'
$SystemForDisplaying.Font = 'Arial, 7'
$SystemForDisplaying.width = 70
$SystemForDisplaying.height = 20
$SystemForDisplaying.FlatStyle = 'standard'
$SystemForDisplaying.BackColor = 'White'
$SystemForDisplaying.ForeColor = 'Black'
$SystemForDisplaying.Location = '325,105'
$SystemForDisplaying.Visible = $true
$Infosystem = New-Object System.Windows.Forms.ToolTip
$Infosystem.SetToolTip($SystemForDisplaying, 'Shortcut "Ctrl + S" `n Display system submenues')
$folderForm.Controls.Add($SystemForDisplaying)
$SystemForDisplaying.bringtofront()

$Delete = New-Object System.Windows.Forms.Button
$Delete.Text = 'Del. cache'
$Delete.Font = 'Arial, 7'
$Delete.width = 70
$Delete.height = 20
$Delete.FlatStyle = 'standard'
$Delete.BackColor = 'White'
$Delete.ForeColor = 'Black'
$Delete.Location = '1000,275'
$Delete.Visible = $false
$folderForm.Controls.Add($Delete)
$Delete.bringtofront()

$system = New-Object System.Windows.Forms.Button
$system.Text = 'System'
$system.Font = 'Arial, 7'
$system.width = 70
$system.height = 20
$system.FlatStyle = 'flat'
$system.BackColor = 'white'
$system.ForeColor = 'Black'
$system.Location = '325,174'
$system.Visible = $false
$Infosystem = New-Object System.Windows.Forms.ToolTip
$Infosystem.SetToolTip($system, 'Shortcut "Ctrl + S"')
$folderForm.Controls.Add($system)
$system.bringtofront()

$Drivers = New-Object System.Windows.Forms.Button
$Drivers.Text = 'Drivers'
$Drivers.Font = 'Arial, 7'
$Drivers.width = 70
$Drivers.height = 20
$Drivers.FlatStyle = 'flat'
$Drivers.BackColor = 'White'
$Drivers.ForeColor = 'Black'
$Drivers.Location = '405,174'
$Drivers.Visible = $false
$folderForm.Controls.Add($Drivers)
$Drivers.bringtofront()

$Peripherals = New-Object System.Windows.Forms.Button
$Peripherals.Text = 'Peripherals'
$Peripherals.Font = 'Arial, 7'
$Peripherals.width = 70
$Peripherals.height = 20
$Peripherals.FlatStyle = 'flat'
$Peripherals.BackColor = 'White'
$Peripherals.ForeColor = 'Black'
$Peripherals.Location = '485,174'
$Peripherals.Visible = $false
$folderForm.Controls.Add($Peripherals)
$Peripherals.bringtofront()


# tableau processes --------------------------------------
$Processes = New-Object System.Windows.Forms.Button
$Processes.Text = 'Processes'
$Processes.Font = 'Arial, 7'
$Processes.width = 70
$Processes.height = 20
$Processes.FlatStyle = 'standard'
$Processes.BackColor = 'White'
$Processes.ForeColor = 'Black'
$Processes.Location = '405,105'
$Processes.Visible = $true
$infoProcesses = New-Object System.Windows.Forms.ToolTip
$infoProcesses.SetToolTip($Processes, 'Shortcut "Ctrl + P"')
$folderForm.Controls.Add($Processes)
$Processes.bringtofront()

$StopProcess = New-Object System.Windows.Forms.Button
$StopProcess.Text = 'Stop'
$StopProcess.Font = 'Arial, 7'
$StopProcess.width = 70
$StopProcess.height = 20
$StopProcess.FlatStyle = 'standard'
$StopProcess.BackColor = 'White'
$StopProcess.ForeColor = 'Black'
$StopProcess.Location = '1000,275'
$StopProcess.visible = $false
$folderForm.Controls.Add($StopProcess)
$StopProcess.bringtofront()

$RestartProcess = New-Object System.Windows.Forms.Button
$RestartProcess.Text = 'Restart'
$RestartProcess.Font = 'Arial, 7'
$RestartProcess.width = 70
$RestartProcess.height = 20
$RestartProcess.FlatStyle = 'standard'
$RestartProcess.BackColor = 'White'
$RestartProcess.ForeColor = 'Black'
$RestartProcess.Location = '1000,305'
$RestartProcess.visible = $false
$folderForm.Controls.Add($RestartProcess)
$RestartProcess.bringtofront()

$RefreshProcess = New-Object System.Windows.Forms.Button
$RefreshProcess.Text = 'Refresh'
$RefreshProcess.Font = 'Arial, 7'
$RefreshProcess.width = 70
$RefreshProcess.height = 20
$RefreshProcess.FlatStyle = 'standard'
$RefreshProcess.BackColor = 'White'
$RefreshProcess.ForeColor = 'Black'
$RefreshProcess.Location = '1000,335'
$RefreshProcess.visible = $false
$folderForm.Controls.Add($RefreshProcess)
$RefreshProcess.bringtofront()

# Tableau applications --------------------------------------
$Applications = New-Object System.Windows.Forms.Button
$Applications.Text = 'Applications'
$Applications.Font = 'Arial, 7'
$Applications.width = 70
$Applications.height = 20
$Applications.FlatStyle = 'standard'
$Applications.BackColor = 'White'
$Applications.ForeColor = 'Black'
$Applications.Location = '485,105'
$Applications.Visible = $true
$infoApplications = New-Object System.Windows.Forms.ToolTip
$infoApplications.SetToolTip($Applications, 'Shortcut "Ctrl + A"')
$folderForm.Controls.Add($Applications)
$Applications.bringtofront()

$ApplicationsUninstall = New-Object System.Windows.Forms.Button
$ApplicationsUninstall.Text = 'Uninstall'
$ApplicationsUninstall.Font = 'Arial, 7'
$ApplicationsUninstall.width = 70
$ApplicationsUninstall.height = 20
$ApplicationsUninstall.FlatStyle = 'standard'
$ApplicationsUninstall.BackColor = 'White'
$ApplicationsUninstall.ForeColor = 'Black'
$ApplicationsUninstall.Location = '1000,275'
$ApplicationsUninstall.Visible = $false
$folderForm.Controls.Add($ApplicationsUninstall)
$ApplicationsUninstall.bringtofront()

$ApplicationRepair = New-Object System.Windows.Forms.Button
$ApplicationRepair.Text = 'Repair'
$ApplicationRepair.Font = 'Arial, 7'
$ApplicationRepair.width = 70
$ApplicationRepair.height = 20
$ApplicationRepair.FlatStyle = 'standard'
$ApplicationRepair.BackColor = 'White'
$ApplicationRepair.ForeColor = 'Black'
$ApplicationRepair.Location = '1000,305'
$ApplicationRepair.Visible = $false
$folderForm.Controls.Add($ApplicationRepair)
$ApplicationRepair.bringtofront()

$Applicationrefresh = New-Object System.Windows.Forms.Button
$Applicationrefresh.Text = 'Refresh'
$Applicationrefresh.Font = 'Arial, 7'
$Applicationrefresh.width = 70
$Applicationrefresh.height = 20
$Applicationrefresh.FlatStyle = 'standard'
$Applicationrefresh.BackColor = 'White'
$Applicationrefresh.ForeColor = 'Black'
$Applicationrefresh.Location = '1000,335'
$Applicationrefresh.Visible = $false
$folderForm.Controls.Add($Applicationrefresh)
$Applicationrefresh.bringtofront()

# tableau services --------------------------------------
$Services = New-Object System.Windows.Forms.Button
$Services.Text = 'Services'
$Services.Font = 'Arial, 7'
$Services.width = 70
$Services.height = 20
$Services.FlatStyle = 'standard'
$Services.BackColor = 'White'
$Services.ForeColor = 'Black'
$Services.Location = '565,105'
$Services.Visible = $true
$infoServices = New-Object System.Windows.Forms.ToolTip
$infoServices.SetToolTip($Services, 'Shortcut "Ctrl + J"')
$folderForm.Controls.Add($Services)
$Services.bringtofront()

$StopService = New-Object System.Windows.Forms.Button
$StopService.Text = 'Stop'
$StopService.Font = 'Arial, 7'
$StopService.width = 70
$StopService.height = 20
$StopService.FlatStyle = 'standard'
$StopService.BackColor = 'White'
$StopService.ForeColor = 'Black'
$StopService.Location = '1000,275'
$StopService.visible = $false
$folderForm.Controls.Add($StopService)
$StopService.bringtofront()

$Startservice = New-Object System.Windows.Forms.Button
$Startservice.Text = 'Start'
$Startservice.Font = 'Arial, 7'
$Startservice.width = 70
$Startservice.height = 20
$Startservice.FlatStyle = 'standard'
$Startservice.BackColor = 'White'
$Startservice.ForeColor = 'Black'
$Startservice.Location = '1000,305'
$Startservice.visible = $false
$folderForm.Controls.Add($Startservice)
$Startservice.bringtofront()

$Restartservice = New-Object System.Windows.Forms.Button
$Restartservice.Text = 'Restart'
$Restartservice.Font = 'Arial, 7'
$Restartservice.width = 70
$Restartservice.height = 20
$Restartservice.FlatStyle = 'standard'
$Restartservice.BackColor = 'White'
$Restartservice.ForeColor = 'Black'
$Restartservice.Location = '1000,335'
$Restartservice.visible = $false
$folderForm.Controls.Add($Restartservice)
$Restartservice.bringtofront()

# Logs --------------------------------------
$EventLogs = New-Object System.Windows.Forms.Button
$EventLogs.Text = 'Event Logs'
$EventLogs.Font = 'Arial, 7'
$EventLogs.width = 70
$EventLogs.height = 20
$EventLogs.FlatStyle = 'standard'
$EventLogs.BackColor = 'White'
$EventLogs.ForeColor = 'Black'
$EventLogs.Location = '645,105'
$EventLogs.Visible = $true
$infoEventLogs = New-Object System.Windows.Forms.ToolTip
$infoEventLogs.SetToolTip($EventLogs, 'Shortcut "Ctrl + L"')
$folderForm.Controls.Add($EventLogs)
$EventLogs.bringtofront()

$ApplicationLogs = New-Object System.Windows.Forms.Button
$ApplicationLogs.Text = 'Application Logs'
$ApplicationLogs.Font = 'Arial, 7'
$ApplicationLogs.width = 70
$ApplicationLogs.height = 20
$ApplicationLogs.FlatStyle = 'standard'
$ApplicationLogs.BackColor = 'White'
$ApplicationLogs.ForeColor = 'Black'
$ApplicationLogs.Location = '1000,275'
$ApplicationLogs.Visible = $false
$infoEventLogs = New-Object System.Windows.Forms.ToolTip
$infoEventLogs.SetToolTip($ApplicationLogs, 'Shortcut "Ctrl + L"')
$folderForm.Controls.Add($ApplicationLogs)
$ApplicationLogs.bringtofront()

$SecurityLogs = New-Object System.Windows.Forms.Button
$SecurityLogs.Text = 'Security Log'
$SecurityLogs.Font = 'Arial, 7'
$SecurityLogs.width = 70
$SecurityLogs.height = 20
$SecurityLogs.FlatStyle = 'standard'
$SecurityLogs.BackColor = 'White'
$SecurityLogs.ForeColor = 'Black'
$SecurityLogs.Location = '1000,305'
$SecurityLogs.Visible = $false
$infoEventLogs = New-Object System.Windows.Forms.ToolTip
$infoEventLogs.SetToolTip($SecurityLogs, 'Shortcut "Ctrl + L"')
$folderForm.Controls.Add($SecurityLogs)
$SecurityLogs.bringtofront()

$Logextract = New-Object System.Windows.Forms.Button
$Logextract.Text = 'Extract log'
$Logextract.Font = 'Arial, 7'
$Logextract.width = 70
$Logextract.height = 20
$Logextract.FlatStyle = 'standard'
$Logextract.BackColor = 'White'
$Logextract.ForeColor = 'Black'
$Logextract.Location = '1000,335'
$Logextract.Visible = $false
$infoEventLogs = New-Object System.Windows.Forms.ToolTip
$infoEventLogs.SetToolTip($Logextract, 'Shortcut "Ctrl + L"')
$folderForm.Controls.Add($Logextract)
$Logextract.bringtofront()

# explorateur de fichiers --------------------------------------
$folders = New-Object System.Windows.Forms.Button
$folders.Text = 'File explorer'
$folders.Font = 'Arial, 7'
$folders.width = 70
$folders.height = 20
$Folders.FlatStyle = 'standard'
$Folders.BackColor = 'White'
$Folders.ForeColor = 'Black'
$folders.Location = '745,105'
$Folders.Visible = $true
$infoFolders = New-Object System.Windows.Forms.ToolTip
$infoFolders.SetToolTip($folders, 'Shortcut "Alt + F"')
$folderForm.Controls.Add($folders)
$folders.bringtofront()

# CMD --------------------------------------
$CMD = New-Object System.Windows.Forms.Button
$CMD.Text = 'cmd'
$CMD.Font = 'Arial, 7'
$CMD.width = 70
$CMD.height = 20
$CMD.FlatStyle = 'standard'
$CMD.BackColor = 'White'
$CMD.ForeColor = 'Black'
$CMD.Location = '825,105'
$CMD.Visible = $true
$infoCMD = New-Object System.Windows.Forms.ToolTip
$infoCMD.SetToolTip($CMD, 'Shortcut "Alt + C"')
$folderForm.Controls.Add($CMD)
$CMD.bringtofront()

###GRID###

$gridA1 = New-Object System.Windows.Forms.DataGridView
$gridA1.Location = '40,230'
$gridA1.Size = New-Object Drawing.Point (900,100)
$gridA1.ScrollBars = "Vertical"  
$gridA1.ReadOnly = $false
$gridA1.AutoSize = $false
$gridA1.GridColor = "Black"
$gridA1.AllowUserToResizeRows = $true
$gridA1.ScrollBars = "Both"
$gridA1.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$gridA1.BackgroundColor = [System.Drawing.Color]::FromArgb(255,240,240,240)
$gridA1.DefaultCellStyle.SelectionBackColor = "#2D3250"
$gridA1.SelectionMode = "FullRowSelect"
$gridA1.MultiSelect = $true
$gridA1.ReadOnly = $true
$grida1.Visible = $true
$gridA1.RowHeadersVisible    = $false
$gridA1.ColumnHeadersVisible = $true
$gridA1.AutoSizeColumnsMode =  'Fill'
$gridA1.EnableHeadersVisualStyles = $false
$gridA1.ColumnHeadersDefaultCellStyle.font = [System.Drawing.Font]::new("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$gridA1.ColumnHeadersDefaultCellStyle.BackColor = "#424769"
$gridA1.ColumnHeadersDefaultCellStyle.ForeColor = "White"
$gridA1.ColumnHeadersDefaultCellStyle.Alignment = "MiddleCenter"
$gridA1.ColumnHeadersBorderStyle = "Single"
$Folderform.Controls.Add($gridA1)

$gridA2 = New-Object System.Windows.Forms.DataGridView
$gridA2.Location = '40,360'
$gridA2.Size = New-Object Drawing.Point (900,100)
$gridA2.ScrollBars = "Vertical"  
$gridA2.ReadOnly = $false
$gridA2.AutoSize = $false
$gridA2.GridColor = "Black"
$gridA2.AllowUserToResizeRows = $true
$gridA2.ScrollBars = "Both"
$gridA2.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$gridA2.BackgroundColor = [System.Drawing.Color]::FromArgb(255,240,240,240)
$gridA2.DefaultCellStyle.SelectionBackColor = "#2D3250"
$gridA2.SelectionMode = "FullRowSelect"
$gridA2.MultiSelect = $true
$gridA2.ReadOnly = $true
$gridA2.Visible = $true
$gridA2.RowHeadersVisible    = $false
$gridA2.ColumnHeadersVisible = $true
$gridA2.AutoSizeColumnsMode =  'Fill'
$gridA2.EnableHeadersVisualStyles = $false
$gridA2.ColumnHeadersDefaultCellStyle.font = [System.Drawing.Font]::new("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$gridA2.ColumnHeadersDefaultCellStyle.BackColor = "#424769"
$gridA2.ColumnHeadersDefaultCellStyle.ForeColor = "White"
$gridA2.ColumnHeadersDefaultCellStyle.Alignment = "MiddleCenter"
$gridA2.ColumnHeadersBorderStyle = "Single"
$Folderform.Controls.Add($gridA2)

$gridA3 = New-Object System.Windows.Forms.DataGridView
$gridA3.Location = '40,490'
$gridA3.Size = New-Object Drawing.Point (900,100)
$gridA3.ScrollBars = "Vertical"  
$gridA3.ReadOnly = $false
$gridA3.AutoSize = $false
$gridA3.GridColor = "Black"
$gridA3.AllowUserToResizeRows = $true
$gridA3.ScrollBars = "Both"
$gridA3.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$gridA3.BackgroundColor = [System.Drawing.Color]::FromArgb(255,240,240,240)
$gridA3.DefaultCellStyle.SelectionBackColor = "#2D3250"
$gridA3.SelectionMode = "FullRowSelect"
$gridA3.MultiSelect = $true
$gridA3.ReadOnly = $true
$gridA3.Visible = $true
$gridA3.RowHeadersVisible    = $false
$gridA3.ColumnHeadersVisible = $true
$gridA3.AutoSizeColumnsMode =  'Fill'
$gridA3.EnableHeadersVisualStyles = $false
$gridA3.ColumnHeadersDefaultCellStyle.font = [System.Drawing.Font]::new("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$gridA3.ColumnHeadersDefaultCellStyle.BackColor = "#424769"
$gridA3.ColumnHeadersDefaultCellStyle.ForeColor = "White"
$gridA3.ColumnHeadersDefaultCellStyle.Alignment = "MiddleCenter"
$gridA3.ColumnHeadersBorderStyle = "Single"
$Folderform.Controls.Add($gridA3)

$gridA4 = New-Object System.Windows.Forms.DataGridView
$gridA4.Location = '40,620'
$gridA4.Size = New-Object Drawing.Point (900,100)
$gridA4.ScrollBars = "Vertical"  
$gridA4.ReadOnly = $false
$gridA4.AutoSize = $false
$gridA4.GridColor = "Black"
$gridA4.AllowUserToResizeRows = $true
$gridA4.ScrollBars = "Both"
$gridA4.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$gridA4.BackgroundColor = [System.Drawing.Color]::FromArgb(255,240,240,240)
$gridA4.DefaultCellStyle.SelectionBackColor = "#2D3250"
$gridA4.SelectionMode = "FullRowSelect"
$gridA4.MultiSelect = $true
$gridA4.ReadOnly = $true
$gridA4.Visible = $true
$gridA4.RowHeadersVisible    = $false
$gridA4.ColumnHeadersVisible = $true
$gridA4.AutoSizeColumnsMode =  'Fill'
$gridA4.EnableHeadersVisualStyles = $false
$gridA4.ColumnHeadersDefaultCellStyle.font = [System.Drawing.Font]::new("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$gridA4.ColumnHeadersDefaultCellStyle.BackColor = "#424769"
$gridA4.ColumnHeadersDefaultCellStyle.ForeColor = "White"
$gridA4.ColumnHeadersDefaultCellStyle.Alignment = "MiddleCenter"
$gridA4.ColumnHeadersBorderStyle = "Single"
$Folderform.Controls.Add($gridA4)

$gridA5 = New-Object System.Windows.Forms.DataGridView
$gridA5.Location = '40,750'
$gridA5.Size = New-Object Drawing.Point (900,100)
$gridA5.ScrollBars = "Vertical"  
$gridA5.ReadOnly = $false
$gridA5.AutoSize = $false
$gridA5.GridColor = "Black"
$gridA5.AllowUserToResizeRows = $true
$gridA5.ScrollBars = "Both"
$gridA5.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$gridA5.BackgroundColor = [System.Drawing.Color]::FromArgb(255,240,240,240)
$gridA5.DefaultCellStyle.SelectionBackColor = "#2D3250"
$gridA5.SelectionMode = "FullRowSelect"
$gridA5.MultiSelect = $true
$gridA5.ReadOnly = $true
$gridA5.Visible = $true
$gridA5.RowHeadersVisible    = $false
$gridA5.ColumnHeadersVisible = $true
$gridA5.AutoSizeColumnsMode =  'Fill'
$gridA5.EnableHeadersVisualStyles = $false
$gridA5.ColumnHeadersDefaultCellStyle.font = [System.Drawing.Font]::new("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$gridA5.ColumnHeadersDefaultCellStyle.BackColor = "#424769"
$gridA5.ColumnHeadersDefaultCellStyle.ForeColor = "White"
$gridA5.ColumnHeadersDefaultCellStyle.Alignment = "MiddleCenter"
$gridA5.ColumnHeadersBorderStyle = "Single"
$Folderform.Controls.Add($gridA5)

###########################################################################################################################################

function Out-DataTable
{
<#
.SYNOPSIS
    Creates a DataTable for an object

.DESCRIPTION
    Creates a DataTable based on an object's properties.

.PARAMETER InputObject
    One or more objects to convert into a DataTable

.PARAMETER NonNullable
    A list of columns to set disable AllowDBNull on

.INPUTS
    Object
        Any object can be piped to Out-DataTable

.OUTPUTS
   System.Data.DataTable

.EXAMPLE
    $dt = Get-psdrive | Out-DataTable
    
    # This example creates a DataTable from the properties of Get-psdrive and assigns output to $dt variable

.EXAMPLE
    Get-Process | Select Name, CPU | Out-DataTable | Invoke-SQLBulkCopy -ServerInstance $SQLInstance -Database $Database -Table $SQLTable -force -verbose

    # Get a list of processes and their CPU, create a datatable, bulk import that data

.NOTES
    Adapted from script by Marc van Orsouw and function from Chad Miller
    Version History
    v1.0 - Chad Miller - Initial Release
    v1.1 - Chad Miller - Fixed Issue with Properties
    v1.2 - Chad Miller - Added setting column datatype by property as suggested by emp0
    v1.3 - Chad Miller - Corrected issue with setting datatype on empty properties
    v1.4 - Chad Miller - Corrected issue with DBNull
    v1.5 - Chad Miller - Updated example
    v1.6 - Chad Miller - Added column datatype logic with default to string
    v1.7 - Chad Miller - Fixed issue with IsArray
    v1.8 - ramblingcookiemonster - Removed if($Value) logic. This would not catch empty strings, zero, $false and other non-null items
                                  - Added perhaps pointless error handling

.LINK
    https://github.com/RamblingCookieMonster/PowerShell

.LINK
    Invoke-SQLBulkCopy

.LINK
    Invoke-Sqlcmd2

.LINK
    New-SQLConnection

.FUNCTIONALITY
    SQL
#>
    [CmdletBinding()]
    [OutputType([System.Data.DataTable])]
    param(
        [Parameter( Position=0,
                    Mandatory=$true,
                    ValueFromPipeline = $true)]
        [PSObject[]]$InputObject,

        [string[]]$NonNullable = @()
    )

    Begin
    {
        $dt = New-Object Data.datatable  
        $First = $true 

        function Get-ODTType
        {
            param($type)

            $types = @(
                'System.Boolean',
                'System.Byte[]',
                'System.Byte',
                'System.Char',
                'System.Datetime',
                'System.Decimal',
                'System.Double',
                'System.Guid',
                'System.Int16',
                'System.Int32',
                'System.Int64',
                'System.Single',
                'System.UInt16',
                'System.UInt32',
                'System.UInt64')

            if ( $types -contains $type ) {
                Write-Output "$type"
            }
            else {
                Write-Output 'System.String'
            }
        } #Get-Type
    }
    Process
    {
        foreach ($Object in $InputObject)
        {
            $DR = $DT.NewRow()  
            foreach ($Property in $Object.PsObject.Properties)
            {
                $Name = $Property.Name
                $Value = $Property.Value
                
                #RCM: what if the first property is not reflective of all the properties? Unlikely, but...
                if ($First)
                {
                    $Col = New-Object Data.DataColumn  
                    $Col.ColumnName = $Name  
                    
                    #If it's not DBNull or Null, get the type
                    if ($Value -isnot [System.DBNull] -and $Value -ne $null)
                    {
                        $Col.DataType = [System.Type]::GetType( $(Get-ODTType $property.TypeNameOfValue) )
                    }
                    
                    #Set it to nonnullable if specified
                    if ($NonNullable -contains $Name )
                    {
                        $col.AllowDBNull = $false
                    }

                    try
                    {
                        $DT.Columns.Add($Col)
                    }
                    catch
                    {
                        Write-Error "Could not add column $($Col | Out-String) for property '$Name' with value '$Value' and type '$($Value.GetType().FullName)':`n$_"
                    }
                }  
                
                Try
                {
                    #Handle arrays and nulls
                    if ($property.GetType().IsArray)
                    {
                        $DR.Item($Name) = $Value | ConvertTo-XML -As String -NoTypeInformation -Depth 1
                    }
                    elseif($Value -eq $null)
                    {
                        $DR.Item($Name) = [DBNull]::Value
                    }
                    else
                    {
                        $DR.Item($Name) = $Value
                    }
                }
                Catch
                {
                    Write-Error "Could not add property '$Name' with value '$Value' and type '$($Value.GetType().FullName)'"
                    continue
                }

                #Did we get a null or dbnull for a non-nullable item? let the user know.
                if($NonNullable -contains $Name -and ($Value -is [System.DBNull] -or $Value -eq $null))
                {
                    write-verbose "NonNullable property '$Name' with null value found: $($object | out-string)"
                }

            } 

            Try
            {
                $DT.Rows.Add($DR)
            }
            Catch
            {
                Write-Error "Failed to add row '$($DR | Out-String)':`n$_"
            }

            $First = $false
        }
    } 
     
    End
    {
        Write-Output @(,$dt)
    }

} 

function Invoke-WithLoadingAnimation {

    param (
        [Parameter(Mandatory=$true)]
        [scriptblock]$ScriptBlock,
        
        [Parameter(Mandatory=$false)]
        [object[]]$ArgumentList,
        
        [Parameter(Mandatory=$false)]
        [string]$LoadingMessage = "WORK IN PROGRESS"
    )

    $job = Start-Job -ScriptBlock $ScriptBlock -ArgumentList $ArgumentList

    $spinner = @('|', '/', '-', '\')
    $index = 0

    while ($job.State -eq 'Running') {
        $Loadinglabel.text = "`r$LoadingMessage $($spinner[$index])"
        $index = ($index + 1) % 4
        Start-Sleep -Milliseconds 100
    }

    $result = Receive-Job $job
    Remove-Job $job

    $loadinglabel.text = ''

    return $result

} 

function ping{

    $computer = $computerName.text

    $result = Invoke-WithLoadingAnimation -ScriptBlock {

        param($computer)

        test-connection $computer -count 2

    } -ArgumentList $computer -LoadingMessage "CONNECTION"

    if($result){

        $StatusLabel.text = "Status: connected"
        $StatusLabel.forecolor = 'green'

    }

    Else{

        $StatusLabel.text = "Status: disconnected"
        $StatusLabel.forecolor = 'red'

    }

}

function refresh{

    # boutons
    $StopProcess.visible = $false
    $RefreshProcess.visible = $false
    $stopservice.visible = $false
    $startservice.visible = $false
    $restartservice.visible = $false
    $RestartProcess.visible = $false
    $ApplicationsUninstall.visible = $false
    $Applicationrefresh.Visible = $false
    $ApplicationRepair.visible = $false
    $delete.Visible = $false
    $Logextract.Visible = $false
    $ApplicationLogs.Visible = $false
    $SecurityLogs.Visible = $false
    $Drivers.Visible = $false
    $system.visible = $false
    $peripherals.visible = $false

    # grids
    $gridA1.datasource = $null
    $gridA2.datasource = $null
    $gridA3.datasource = $null
    $gridA4.datasource = $null
    $gridA5.datasource = $null

    # labels
    $controlslabel.visible = $false
    $frame.visible = $false
    $frame2.visible = $false
    $header4.visible = $false
    $Titre1.text = '' 
    $Titre2.text = ''
    $Titre3.text = ''
    $Titre4.text = ''
    $titre5.text = ''
    $CountLabel.Text = ''
    $CountLabel2.Text = ''
    $IndcationsSystem.text = ''

}

function refreshApplications {

    refresh
    start-sleep -Seconds 1
    $gridA1.Size = New-Object Drawing.Point (900,599)
    $gridA1.BringToFront()
    $Titre1.text = "APPLICATIONS"
    $ApplicationsUninstall.visible = $true
    $controlslabel.visible = $true
    $frame2.visible = $true
    $frame.visible = $true
    $ApplicationRefresh.visible = $true
    $ApplicationRepair.visible = $true
    
    $computer = $computername.Text

    $App = get-wmiobject Win32_Product -ComputerName "$computer" -ea silentlycontinue
    $Appcount = $App.Count
    $App.size

    $gridA1.DataSource = $App | select-object Name,Version | sort-object Name | Out-DataTable

    $CountLabel.Location = (35,1000)
    $CountLabel.text = "Total applications: $Appcount"

}

function RefreshProcesses {

    $Stopprocess.visible = $true
    $RefreshProcess.visible = $true
    $Controlslabel.visible = $true
    $frame.visible = $true
    $RestartProcess.visible = $true
    $frame2.Visible = $true

    $computer = $computername.text

    if($computer -like $env:COMPUTERNAME){

            $processes = get-process

        }

    else{

            $processes = invoke-command -ComputerName $computer -scriptblock{

                Get-Process

            } 

    }

    $global:TotalProcess = $processes.count
    $global:TotalRAMusage = [math]::round(($processes | Measure-Object -Property WorkingSet64 -Sum).Sum / 1MB, 2)

    # Définit le tableau out-datatable
    $gridA1.DataSource = $processes | select-object name,ID,@{Name="Memory (MB)";Expression={[math]::Round($_.WS / 1MB, 2)}},@{Name="CPU"; Expression={[math]::Round($_.CPU, 2)}},StartTime | Out-DataTable

}

function RefreshServices{

    $stopservice.visible = $true
    $startservice.visible = $true
    $restartservice.visible = $true
    $Controlslabel.visible = $true
    $frame.visible = $true
    $frame2.Visible = $true
    $gridA1.bringtofront()

    $services = Get-WmiObject -Class Win32_Service -ComputerName $computer

    $gridA1.DataSource = $services | select-object name,processID,state,description  | Out-DataTable

    $runingservices = $services | where-object {$_.state -eq "running"}
    $runingservices = $runingservices.count

    $CountLabel.text = "Total services: $($services.count)"
    $CountLabel2.text = "Total running services: $runingservices"

}

$script:isPinging = $false

function pingms {

    if($script:isPinging){
        
        $ping = Test-Connection -ComputerName $computername.Text -Count 1 -ErrorAction Stop
        $PingMSLabel.Text = "Ping: $($ping.ResponseTime) ms"
   
    }

}

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 5000
$timer.Add_Tick({ pingms })

###########################################################################################################################################

$Connection.add_click({
    
    If($computername.text -ne $null){

        refresh
        $computer = $computerName.Text
        ping

        $pingms = test-connection -ComputerName $computer -count 1
        $pingms = $pingms.ResponseTime
        $pingmslabel.text = "Ping: $pingms ms"
        $ComputerNetwork = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName "$computer" -ea SilentlyContinue | Where-Object { $_.IPEnabled -eq $true }
        $IP = $Computernetwork.IPAddress[0]
        $IPLabel.text = "IP: $IP"
        $Adaptertype =  Get-WmiObject Win32_NetworkAdapter -computername $computer | Where-Object { $_.NetEnabled -eq $true }
        $Adaptertype = $Adaptertype.AdapterType
        $AdapterLabel.text = "Adapter: $adaptertype"

        $OS = Get-WmiObject win32_operatingsystem -ComputerName "$computer" -ea silentlycontinue
        $LastBootUp = $os.LastBootUpTime
        $DayOfStart = [Management.ManagementDateTimeConverter]::ToDateTime($LastBootUp)
        $StartedDate = [Management.ManagementDateTimeConverter]::ToDateTime($LastBootUp)
        $Date = $os.ConvertTodateTime($os.LocalDateTime)
        $Up = $date - $StartedDate
        $UpTime = "$($up.days)d $($up.Hours)h $($up.minutes)m"
        $Uptimelabel.text = "Uptime: $uptime"

        # Variables
        $ComputerSystem = Get-WmiObject Win32_ComputerSystem -ComputerName "$computer" -ea SilentlyContinue
        
        # Définit les objets
        $Nodename = $ComputerSystem.Name
        $Owner = $computersystem.UserName
        $Owner = $Owner -split "\\"
        $owner = $Owner[1]

        # MàJ labels
        $NodenameLabel.text = "PC: $($Nodename)"
        $OwnerLabel.text = "User: $($Owner)"

        $script:isPinging = !$script:isPinging

        if($script:ispinging -and $statuslabel.text -eq "Status: connected"){

            $connection.text = "Disconnection"
            $connection.font = "arial,6"
            $computername.enabled = $false
            $timer.Start()

        }

        else{

            $connection.text = "Connection"
            $NodenameLabel.text = 'PC:'
            $OwnerLabel.text = 'User:'
            $Statuslabel.text = "Status:"
            $StatusLabel.ForeColor = "black"
            $connection.font = "arial,7"
            $PingMSLabel.Text = "Ping: ms"
            $AdapterLabel.text = "Adapter:"
            $IPlabel.text = "IP:"
            $Uptimelabel.text = "Uptime:"
            $computername.enabled = $true
            $timer.Stop()
            $timer.Dispose()

        }

    }

})

$SystemForDisplaying.add_click({

    refresh

    $header4.visible = $true
    $drivers.visible = $true
    $peripherals.visible = $true
    $system.Visible = $true
    $IndcationsSystem.text = '::: ACTION REQUIRED :::     Select informations to display :'

})

$system.add_click({

    $computer = $computerName.Text
    $gridA1.Size = New-Object Drawing.Point (900,100)
    $gridA2.Size = New-Object Drawing.Point (900,100)
    $gridA3.Size = New-Object Drawing.Point (900,100)
    $gridA4.Size = New-Object Drawing.Point (900,100)
    $gridA5.Size = New-Object Drawing.Point (900,100)
    $Frame.visible = $true
    $delete.Visible = $true
    $ControlsLabel.visible = $true
    $Drivers.Visible = $true
    $frame2.visible = $false
   

    # Variables
    $InfosPC = invoke-command -ComputerName $Computer -ScriptBlock({Get-ComputerInfo})
    $OS = Get-WmiObject win32_operatingsystem -ComputerName "$computer" -ea silentlycontinue
    $ComputerSystem = Get-WmiObject Win32_ComputerSystem -ComputerName "$computer" -ea SilentlyContinue
    $ComputerNetwork = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName "$computer" -ea SilentlyContinue | Where-Object { $_.IPEnabled -eq $true }
    $RAM = Get-WmiObject Win32_PhysicalMemory -ComputerName "$computer" -ea SilentlyContinue
    $Processor = Get-WmiObject Win32_Processor -ComputerName "$computer" -ea SilentlyContinue
    $LogicalDisk = Get-WmiObject win32_logicaldisk -ComputerName "$computer" -ea SilentlyContinue
    $Adaptertype =  Get-WmiObject Win32_NetworkAdapter -computername $computer | Where-Object { $_.NetEnabled -eq $true }

    # Définit les objets à réupérer
    # Tableau 1 Hardware
    $Nodename = $ComputerSystem.Name
    $Model = $ComputerSystem.model
    $TotalRAMBytes = ($RAM | Measure-Object -Property Capacity -Sum).Sum
    $TotalRAMGB = [math]::round($TotalRAMBytes / 1GB, 2)
    $ProcessorName = $Processor.name
    $DiskName = $LogicalDisk.DeviceID
    $DiskFreeSpace = [math]::round($LogicalDisk.FreeSpace / 1GB, 2)

    #Tableau 2 Mapped drives
    $results = Invoke-Command -ComputerName $computer -ScriptBlock {

        $userSessions = Get-WmiObject -Class Win32_ComputerSystem | Select-Object -ExpandProperty UserName
        $userSID = (New-Object System.Security.Principal.NTAccount($userSessions)).Translate([System.Security.Principal.SecurityIdentifier]).Value
        $registryPath = "Registry::HKU\$userSID\Network\*"
    
            $mappedDrives = @()
            Get-ItemProperty -Path $registryPath | foreach-object {
    
                $mappedDrives += [PSCustomObject]@{
                letter = $_.PSchildname
                path = $_.Remotepath
    
            }
        
        }
    
        return $mappedDrives

    }

    # Tableau 2 OS
    $Name = $OS.Name
    $Name = $Name.Split('|')[0]
    $Version = $InfosPC.osVersion
    $Build = $OS.BuildNumber
    $DisplayVersion = $InfosPC.OSDisplayversion
    $Deploymentdate = ($InfosPC).WindowsInstallDateFromRegistry
    # Tableau 3 Network
    $Adaptertype = $Adaptertype.AdapterType
    $MAC = ($adaptertype).MACAddress
    $IP = $Computernetwork.IPAddress[0]
    # Tableau 4 Uptime
    $LastBootUp = $os.LastBootUpTime
    $DayOfStart = [Management.ManagementDateTimeConverter]::ToDateTime($LastBootUp)
    $StartedDate = [Management.ManagementDateTimeConverter]::ToDateTime($LastBootUp)
    $Date = $os.ConvertTodateTime($os.LocalDateTime)
    $Up = $date - $StartedDate
    $UpTime = "$($up.days)d $($up.Hours)h $($up.minutes)m"

    # Définit le tableau 1 Out-datatable
    $Titre1.text = "HARDWARE"
    $Hardware = new-object psobject
    $Hardware | Add-Member noteproperty "Model" $Model
    $Hardware | Add-Member noteproperty "Processor" $ProcessorName
    $Hardware | Add-Member NoteProperty "Ram" "$TotalRAMGB Go"
    $Hardware | Add-Member noteproperty "Disk" "$Diskname $DiskFreeSpace Go (free)"
    $gridA1.DataSource = $Hardware | Select-Object Model,Processor,Ram,Disk  | Out-datatable

    # Définit le tableau 2 Out-datatable
    $Titre2.text = "OPERATING SYSTEM"
    $osVersion = new-object psobject
    $osVersion | Add-Member noteproperty "Operating System" $Name
    $osVersion | Add-Member noteproperty Version $Version
    $osVersion | Add-Member noteproperty Build $Build
    $osVersion | Add-Member noteproperty "Display Version" $DisplayVersion
    $osVersion | Add-Member noteproperty "Deployment Date" $DeploymentDate
    $gridA2.DataSource = $osVersion | Select-Object "Operating System",version,"Display Version",Build,"Deployment date" | Out-datatable
    
    # Définit le tableau 3 Out-datatable
    $Titre3.text = "NETWORK"
    $Network = new-object psobject
    $Network | Add-Member noteproperty Adapter $AdapterType
    $Network | Add-Member noteproperty IP $IP
    $Network | Add-Member noteproperty MAC $MAC
    $gridA3.DataSource = $Network | Select-Object Adapter,IP,MAC | Out-datatable

    # Définit le tableau 4 Out-datatable
    $Titre4.text = "UPTIME"
    $UptimeGrid = new-object psobject
    $UptimeGrid | Add-Member noteproperty LastBootUp $DayOfStart
    $UptimeGrid | Add-Member noteproperty Uptime $Uptime
    $gridA4.DataSource = $UptimeGrid | Select-Object LastBootUp,Uptime | Out-datatable

    # définit le tableau 5 mapdrives
    $Titre5.text = "MAP DRIVES"
    $gridA5.DataSource = $results | Select-Object Letter,Path | Out-datatable
    
})

$delete.add_click({

    $computer = $computerName.text

    #affiche une pop-up qui demande une confirmation:
    $msgbox = [System.Windows.Forms.MessageBox]::Show("Cette action supprimera les fichiers temp. Voulez-vous poursuivre ?","confirmation","yesnocancel")

        #bloc de reverseshell qui exécute la suppression des fichiers contenus dans les dossiers de cache du pc distant si le bouton 'yes' est pressé:
        if ($msgbox -eq 'yes'){

            Invoke-Command -ComputerName $computer -ScriptBlock {

                #création du fichier de log:
                param($username)
                
                #suppression des fichiers:
                Remove-Item -path "C:\Windows\Temp\*" -r -force 2>$null 
                Remove-Item -path "C:\Windows\ccmcache\*" -r -force 2>$null
                Remove-item -path "C:\Users\$username\AppData\Local\Temp\*" -r -force 2>$null -path

            } -ArgumentList $username

        }

        else{

                $popup = New-Object -ComObject Wscript.Shell
                $popup.Popup("Opération annulée.",0,"État",0x1)

        }

})

$Drivers.add_click({

    $computer = $computername.text
    $Frame.visible = $true
    $gridA1.datasource = ''
    $gridA2.datasource = ''
    $gridA3.datasource = ''
    $gridA4.datasource = ''
    $gridA5.datasource = ''
    $gridA1.BringToFront()
    $gridA1.Size = New-Object Drawing.Point (900,599)
    $frame2.Visible = $true
    $delete.Visible = $false
    
    $Titre1.text = 'DRIVERS'
    $Titre2.text = ''
    $Titre3.text = ''
    $Titre4.text = ''
    $titre5.text = ''
    $countlabel.text = ''

    $drivers = Get-WmiObject Win32_PnPSignedDriver -computername $computer | select-object FriendlyName,DriverProviderName,Description,DriverVersion 
    $gridA1.DataSource = $drivers | select-object FriendlyName,DriverProviderName,Description,DriverVersion | Out-DataTable

})

$Peripherals.add_click({

    $computer = $computername.Text
    $gridA1.Size = New-Object Drawing.Point (900,599)
    $Frame.visible = $true
    $frame2.Visible = $true
    $delete.Visible = $false

    $Titre1.text = 'PERIPHERALS'
    $Titre2.text = ''
    $Titre3.text = ''
    $Titre4.text = ''
    $titre5.text = ''
    $countlabel.text = ''
    $GridA1.bringtofront()

    $totalperipherals = Get-WmiObject -Class Win32_PnPEntity -computername $computer | select-object caption
    $totalperipherals = $totalperipherals.count

    $peripherals1 = Get-WmiObject -Class Win32_PnPEntity -computername $computer | select-object name,pnpclass,description,caption,manufacturer
    $gridA1.DataSource = $peripherals1 | select-object name,pnpclass,description,caption,manufacturer | Out-DataTable

    $CountLabel.text = "Total peripherals: $totalperipherals"

    # #$deviceId = ""
    # $device = Get-WmiObject -Query "SELECT * FROM Win32_PnPEntity WHERE name='$deviceId'"
    # $device.disable()
    # start-sleep -seconds 2
    # $device.enable()

})

$Applications.add_click({

    refreshApplications

})

$ApplicationsUninstall.add_click({

    $computer = $computername.Text

    # Détermine la case sélectionnée dans le gridview des process
    $row = $gridA1.SelectedCells[0].RowIndex
    $RowTotalNumber = ($gridA1.Rowcount -1)

    if($Row -ne $RowTotalNumber -and $Row -ne $null -or $row -ne $empty){

        $IDapp = $GridA1.SelectedCells[0].Value

    }

    if($computer -eq $env:COMPUTERNAME){

        $App = Get-WmiObject -Class Win32_Product | where-object {$_.name -like "*$IDapp"}
        $App.uninstall()

    }

    else{

        Invoke-Command -ComputerName $computer -ScriptBlock {
        
            param($IDapp)
        
            $App = Get-WmiObject -Class Win32_Product | where-object {$_.name -like "*$IDapp"}
            $App.uninstall()
        
        } -ArgumentList $IDapp

    } 


})

$ApplicationRepair.add_click({

    $computer = $computername.Text

    # Détermine la case sélectionnée dans le gridview des process
    $row = $gridA1.SelectedCells[0].RowIndex
    $RowTotalNumber = ($gridA1.Rowcount -1)

    if($Row -ne $RowTotalNumber -and $Row -ne $null -or $row -ne $empty){

        $IDapp = $GridA1.SelectedCells[0].Value

    }

    if($computer -eq $env:COMPUTERNAME){

        $App = Get-WmiObject -Class Win32_Product | where-object {$_.name -like "*$IDapp"}
        $App.repair()

    }

    else{

        Invoke-Command -ComputerName $computer -ScriptBlock {
        
            param($IDapp)
        
            $App = Get-WmiObject -Class Win32_Product | where-object {$_.name -like "*$IDapp"}
            $App.repair()
        
        } -ArgumentList $IDapp

    }

})

$ApplicationRefresh.add_click({

    refreshApplications

})

$Processes.add_click({
    
    refresh
    start-sleep -seconds 1
    $gridA1.Size = New-Object Drawing.Point (900,599)
    $Titre1.text = "PROCESSES"
    $Stopprocess.visible = $true
    $GridA1.bringtofront()

    refreshProcesses
    $CountLabel.text = "Total processes: $global:TotalProcess"
    $CountLabel2.text = "Total RAM usage: $global:TotalRAMusage MB"
    
})

$RefreshProcess.add_click({

    refreshprocesses
    $CountLabel.text = "Total processes: $global:TotalProcess"
    $CountLabel2.text = "Total RAM usage: $global:TotalRAMusage MB"

})

$StopProcess.add_click({
    
    $computer = $computername.text

    # Détermine la case sélectionnée dans le gridview des process
    $row = $gridA1.SelectedCells[0].RowIndex
    $RowTotalNumber = ($gridA1.Rowcount -1)

    if($Row -ne $RowTotalNumber -and $Row -ne $null -or $row -ne $empty){

        $ID = $GridA1.SelectedCells[0].Value

    }

    if($computer -eq $ENV:computername){

        stop-process -name $ID -force

    }

    else{
    
        invoke-command -ComputerName $computer -ScriptBlock({

            param($ID)

            stop-process -name $ID -Force

        }) -ArgumentList $ID
    }

    start-sleep -Milliseconds 500
    RefreshProcesses
    start-sleep -Milliseconds 500
    $CountLabel.text = "Total processes: $global:TotalProcess"
    $CountLabel2.text = "Total RAM usage: $global:TotalRAMusage MB"

})

$restartprocess.add_click({
    
    $computer = $computername.Text

    # D�termine la case s�lectionn�e dans le GridView des processus
    $row = $gridA1.SelectedCells[0].RowIndex
    $RowTotalNumber = $gridA1.RowCount - 1

    if ($row -ne $null -and $row -le $RowTotalNumber) {
        $ID = $gridA1.SelectedCells[0].Value

        if ($computer -eq $ENV:COMPUTERNAME) {
            # R�cup�rer tous les processus dont le nom contient $ID
            $processes = Get-Process | Where-Object { $_.Name -like "*$ID*" }

            # Stocker les noms et chemins des processus avant de les arr�ter
            $processInfo = @()
            foreach ($process in $processes) {
                $processInfo += [PSCustomObject]@{
                    Name = $process.Name
                    Path = $process.Path
                }
            }

            # Arr�ter chaque processus trouv�
            foreach ($process in $processes) {

                Stop-Process -Id $process.Id -Force

            }

            Start-Sleep -Seconds 1
            
            foreach ($info in $processInfo) {

                Start-Process -FilePath $info.Path

            }

        }

        if ($computer -ne $ENV:COMPUTERNAME) {

            $user = (Get-WmiObject -Class Win32_ComputerSystem).UserName -replace '.*\\(.{6})', '$1'
            $batchpath = $Batchpath = "c:\users\$user\restartapp.bat"
            
            Invoke-Command -ComputerName $computer -ScriptBlock {
                
                param ($ID, $computer, $batchpath, $user)
            
                # Utilisation de Get-WmiObject pour obtenir les processus et leur chemin
                $processes = Get-WmiObject Win32_Process | Where-Object { $_.Name -like "*$ID*" }
            
                # Parcourir chaque processus et obtenir son chemin d'exécutable
                foreach ($process in $processes) {

                    $processPath = $process.Path
                    
                    Stop-Process -name $ID -Force
                
                    Start-Sleep -Seconds 1
                
                    # Vérifier si le chemin d'exécutable n'est pas nul
                    if ($processPath) {
                        
                        # Créer le contenu du bat
                        $batchContent = "start """" `"$processPath`""
                        Set-Content -Path $batchpath -Value $batchContent
            
                        $tacheDescription = "à supprimer" 
                        $tacheAction = New-ScheduledTaskAction -Execute 'cmd.exe' -Argument "/c $batchpath" 
                        $taskDeclencheur = New-ScheduledTaskTrigger -Once -at (get-date).AddSeconds(3)
                        
                        # Création et exécution de la tâche planifiée
                        Register-ScheduledTask -TaskName "RelaunchApp" -Action $tacheAction -Trigger $taskDeclencheur -Description $tacheDescription -User $user -TaskPath "\"
            
                        Start-ScheduledTask -TaskName "RelaunchApp"
            
                        # Suppression de la tâche planifiée et du fichier batch
                        Unregister-ScheduledTask -TaskName "RelaunchApp" -Confirm:$false
                        #Remove-Item -Path $batchpath -Force
                    }
                }
            
            } -ArgumentList $ID, $computer, $batchpath, $user
        }
        
    }

    RefreshProcesses

})

$Services.add_click({

    refresh
    start-sleep -seconds 1
    $gridA1.Size = New-Object Drawing.Point (900,599)
    $computer = $computerName.Text
    $Titre1.text = "SERVICES"
    RefreshServices

})

$StopService.add_click({
    
    $computer = $computername.text

    # Détermine la case sélectionnée dans le gridview des process
    $row = $gridA1.SelectedCells[0].RowIndex
    $RowTotalNumber = ($gridA1.Rowcount -1)

    if($Row -ne $RowTotalNumber -and $Row -ne $null -or $row -ne $empty){

        $ID = $GridA1.SelectedCells[0].Value

    }

    if($computer -eq $ENV:computername){

        stop-service -name $ID -force

    }

    else{
    
        invoke-command -ComputerName $computer -ScriptBlock({

            param($ID)

            stop-service -name $ID

        }) -ArgumentList $ID
    }

    Refreshservices

})

$StartService.add_click({
    
    $computer = $computername.text

    # Détermine la case sélectionnée dans le gridview des process
    $row = $gridA1.SelectedCells[0].RowIndex
    $RowTotalNumber = ($gridA1.Rowcount -1)

    if($Row -ne $RowTotalNumber -and $Row -ne $null -or $row -ne $empty){

        $ID = $GridA1.SelectedCells[0].Value

    }

    if($computer -eq $ENV:computername){

        start-service -name $ID

    }

    else{
    
        invoke-command -ComputerName $computer -ScriptBlock({

            param($ID)

            start-service -name $ID

        }) -ArgumentList $ID
    }

    Refreshservices
     

})

$RestartService.add_click({
    
    $computer = $computername.text

    # Détermine la case sélectionnée dans le gridview des process
    $row = $gridA1.SelectedCells[0].RowIndex
    $RowTotalNumber = ($gridA1.Rowcount -1)

    if($Row -ne $RowTotalNumber -and $Row -ne $null -or $row -ne $empty){

        $ID = $GridA1.SelectedCells[0].Value

    }

    if($computer -eq $ENV:computername){

        restart-service -name $ID

    }

    else{
    
        invoke-command -ComputerName $computer -ScriptBlock({

            param($ID)

            restart-service -name $ID

        }) -ArgumentList $ID
    }

    start-sleep -milliseconds 500
    Refreshservices
    start-sleep -milliseconds 500

})

$EventLogs.add_click({

    refresh
    start-sleep -Seconds 1
    $gridA1.Size = New-Object Drawing.Point (900,599)
    $Titre1.text = "EVENT LOGS"
    $computer = $ComputerName.Text
    $Logextract.Visible = $true
    $Controlslabel.visible = $true
    $frame.visible = $true
    $Frame2.visible = $true
    $securityLogs.Visible = $true
    $ApplicationLogs.Visible = $true
    $gridA1.bringtofront()

    $script:eventstate = 1
    $events = Get-WmiObject -Query "SELECT * FROM Win32_NTLogEvent WHERE Logfile = 'system'" -computername $computer | select-object -first 300
    $gridA1.DataSource = $events | select-object sourcename,message | Out-DataTable

})

$ApplicationLogs.add_click({

    refresh
    start-sleep -Seconds 1
    $gridA1.Size = New-Object Drawing.Point (900,599)
    $Titre1.text = "EVENT LOGS"
    $computer = $ComputerName.Text
    $Logextract.Visible = $true
    $Controlslabel.visible = $true
    $frame.visible = $true
    $Frame2.visible = $true
    $securityLogs.Visible = $true
    $ApplicationLogs.Visible = $true
    $gridA1.bringtofront()

    $script:eventstate = 2
    $eventsApplication = Get-WmiObject -Query "SELECT * FROM Win32_NTLogEvent WHERE Logfile = 'Application'" -computername $computer | select-object -first 300
    $gridA1.DataSource = $eventsApplication | select-object sourcename,message | Out-DataTable


})

$SecurityLogs.add_click({

    refresh
    $gridA1.datasouce = ''
    start-sleep -Seconds 1
    $gridA1.Size = New-Object Drawing.Point (900,599)
    $Titre1.text = "EVENT LOGS"
    $computer = $ComputerName.Text
    $Logextract.Visible = $true
    $Controlslabel.visible = $true
    $frame.visible = $true
    $Frame2.visible = $true
    $securityLogs.Visible = $true
    $ApplicationLogs.Visible = $true
    $gridA1.bringtofront()

    $script:eventstate = 3
    $eventsSecurity = Get-WmiObject -Query "SELECT * FROM Win32_NTLogEvent WHERE Logfile = 'Security'" -computername $computer | select-object -first 300
    $gridA1.DataSource = $eventsSecurity | select-object sourcename,message | Out-DataTable

})

$Logextract.add_click({

    $computer = $ComputerName.Text

    if($script:eventstate -eq 1){

        $events = Get-WmiObject -Query "SELECT * FROM Win32_NTLogEvent WHERE Logfile = 'system'" -computername $computer | select-object -first 300
        $events | out-file "c:\temp\SystemLog.txt"

    }

    if($script:eventstate -eq 2){

        $events = Get-WmiObject -Query "SELECT * FROM Win32_NTLogEvent WHERE Logfile = 'Application'" -computername $computer | select-object -first 300
        $events | out-file "c:\temp\ApplicationLog.txt"

    }

    if($script:eventstate -eq 3){

        $events = Get-WmiObject -Query "SELECT * FROM Win32_NTLogEvent WHERE Logfile = 'Security'" -computername $computer | select-object -first 300
        $events | out-file "c:\temp\SecurityLog.txt"

    }

})

$Folders.add_click({

    $computer = $ComputerName.Text

    $initialDirectory = "\\$computer\c$"
    $folderBrowserDialog = New-Object System.Windows.Forms.OpenFileDialog
    $folderBrowserDialog.InitialDirectory = $initialDirectory
    $folderBrowserDialog.FileName = "Select a Folder"

    if ($folderBrowserDialog.ShowDialog() -eq "OK"){

        $selectedPath = [System.IO.Path]::GetDirectoryName($fileBrowserDialog.FileName)
        [System.Windows.Forms.MessageBox]::Show("Selected folder: " + $selectedPath)

    }

})

$CMD.add_click({

    $computer = $computername.Text

    if($computername.text -ne $null -and $statuslabel.text -eq "Status: connected"){

        $path = "C:\temp"
        $command = "psexec \\$computer cmd /K hostname"
        $argumentlist = "/c cd $path && $command"

        Start-Process cmd.exe -ArgumentList $argumentlist

    }

})

#N�cessaire pour l'affichage de la fen�tre (GUI):
$folderform.ShowDialog()