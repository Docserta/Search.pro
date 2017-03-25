########################################################################
# Script Création Search.pro 
# Cree le 13/03/2017
# Par Christian Frigoult
# Version 1 : Initial
# 
########################################################################

#Fonction Création du fichier Search.pro
#Fonctionnement
#Ecris dans un fichier texte la liste( recursive) de tout les répertoires
#      du repertoire sélectionné
# Permet de selectionner un fichier search.pro existant pour y ajouter une liste de répertoire
# détecte les chemins contenant des espaces 


#=========== VARIABLES
$nFicSearch = "search.pro"
$nLogoeXcent = "excentGroupeSmallBlanc.png"
#===========

#ajout de la list des fichiers au search.pro
Function Create_Searchpro ([string]$PathFicSearch, [string]$PathtoAdd, [bool]$fileexist)
{
# $PathFicSearch = Répertoire du fichier searche.pro çà créer ou mettre à jour
# $PathtoAdd = répertoire à lister et à ajouter au search.pro
# $fileexist = Le fichier search.pro est déjà existant ou non
$SpaceDetect = $false
$MaTable =Get-ChildItem -path $PathtoAdd -Recurse | Where-Object { $_.PSIsContainer } | Select-Object Name,Fullname
if (-not ($fileexist ))
    {
    $FicSearch = New-Item -ItemType file -path $PathFicSearch -Force
    }
    else
    {
    $FicSearch = $PathFicSearch
    }
#Ajout du username et de la date
$strName = $env:username
$strDate = get-date 
ADD-content -path $FicSearch -value "!------ Lignes ajoutées par $strName le $strDate ---#"
#Ajout des lignes
foreach ($MySubFolder in $MaTable){
    #Recheche d'espaces dans le nom du dossier
    $str =$MySubFolder.fullname -replace " ","_"  
    $strligne = $MySubFolder.fullname
    if (-not ($str -eq  $MySubFolder.fullname)) 
        {
        $SpaceDetect = $true
        $strligne = "!---"+ $strligne
        }
    ADD-content -path $FicSearch -value $strligne
}
return $SpaceDetect
}

# MSGBOX
function Show-MessageBox([string] $Message,[string] $Titre="",  [String] $IconType="Information",[String] $BtnType="Ok")
{  #Affiche une boîte de dialogue fenêtrée
   #Prérequis [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

  trap 
  {
   switch ($_.Exception.GetType().FullName) 
   {
    "System.Management.Automation.PSInvalidCastException" {$_.message;Break}
    "System.Management.Automation.RuntimeException" {
     Throw "Assurez-vous que l'assembly [System.Windows.Forms] est bien chargé."}
    default {Throw $_}
    }#switch
  }#trap
  return [Windows.Forms.MessageBox]::Show($Message,$Titre, $BtnType,$IconType)
  
  # Exemples d'appel :
  #   Show-MessageBox "Message" 
  #   Show-MessageBox "Message" "Titre"  
  #   [void](Show-MessageBox "Message" "Titre" )
  #   Show-MessageBox "Message" "Titre"  "Question" "YesNo"
  #   Show-MessageBox "Message" "Titre"  "Error" "AbortRitryIgnore"
  #   Show-MessageBox "Message" "Titre" "Erreur" "AbortRetryIgnore"
  #   Show-MessageBox "Message" "Titre"  "Error" "AbortRetryIgnore"
}

# Select-Folder (Selection d'un répertoire)
function Select-Folder($message='Selectionner un répertoire', $path = 0)
{
$object = New-Object -comObject Shell.Application
$folder = $object.BrowseForFolder(0, $message, 0, $path)
if ($folder -ne $null)
	{
	$folder.self.Path
	}
}

# select-search (selection d'un fichier search.pro existant
function select-search($patch)
{
$fd = New-Object system.windows.forms.openfiledialog
$fd.InitialDirectory = $patch
$fd.MultiSelect = $true
$fd.Filter ="PRO (*.pro)|*.pro"
$fd.showdialog() | out-Null
}

#Vérification de l'éxistence du fichier search.pro
function verif_Search($xfilesearch)
{
If (Test-Path $xfilesearch) {return $true} else {return $false}
}

# Recupération du path du script
function Get-ScriptDirectory
{
$Invocation = (Get-Variable MyInvocation -Scope 1).Value
write-host ("-" * 40)
Split-Path $Invocation.MyCommand.Path
}

#FONCTION FDE GESTION INTERFACE GRAPHIQUE
# Clic sur bouton Quitter 
function ActionBtQuitter {
$form1.Close();
}

#Fonction de selection du répertoire a ajouter au search.pro
#propose de localiser le fichier search.pro dans le répertoire parent du dossier sélectionné
function ActionBt1_SelectFolder {
$xfolderpath=Select-Folder 'Selectionner un répertoire'
$textBox1.Text =$xfolderpath
$textBox3.text =(split-path -path $xfolderpath)+ "\" + $nFicSearch
}

#Fonction de délection d'un fichier Search.pro existant
function ActionBt2_SelectSearch {
$xfilesearch=select-search 'Selectionner le fichier search.pro'
$textBox3.Text =$xfilesearch.filename
}

# Clic sur bouton Creer Search.pro
Function ActionBtCreerSearchpro {
if ($textBox1.text.length -eq 0 )
	{
	Show-MessageBox "Vous n'avez pas choisi de dossier"
	}
Else{
    if ($textBox3.text.length -eq 0)
        {
	    #si pas de fichier search.pro documenté -> on force la localisation du fichier
        $textBox3.text ="c:\temp\search.pro"
	    }
	$xfolderpath = $textBox1.text
    $xNomSearchPro = $textBox3.text
    $fileSearchExist = verif_Search $xNomSearchPro
	if(Create_Searchpro $xNomSearchPro $xfolderpath $fileSearchExist)
        {
        Show-MessageBox "Des espaces ont été détectés dans les chemins. Vérifier le fichier Search.pro"
        Invoke-Expression "notepad.exe $xNomSearchPro" 
        }
        else
        {
        Show-MessageBox "Création du fichier search.pro effectuée"
	    }
    }
}

#Generated Form Function
function GenerateForm {
########################################################################
# Code Generated By: SAPIEN Technologies PrimalForms (Community Edition) v1.0.10.0
# Generated On: 27/03/2015 17:24
# Generated By: stephane.vasselon
########################################################################

#region Import the Assemblies
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
#endregion

#region Generated Form Objects
$form1 = New-Object System.Windows.Forms.Form
$label5 = New-Object System.Windows.Forms.Label
$label1 = New-Object System.Windows.Forms.Label
$textBox3 = New-Object System.Windows.Forms.TextBox
$label2 = New-Object System.Windows.Forms.Label
$btCreerAffaire = New-Object System.Windows.Forms.Button
$BtSelectFolder = New-Object System.Windows.Forms.Button
$BtSelectSearch = New-Object System.Windows.Forms.Button
$textBox1 = New-Object System.Windows.Forms.TextBox
$textBox2 = New-Object System.Windows.Forms.TextBox
$label3 = New-Object System.Windows.Forms.Label
$label4 = New-Object System.Windows.Forms.Label
$btQuitter = New-Object System.Windows.Forms.Button
$label1_Titre = New-Object System.Windows.Forms.Label
$pictureBox1 = New-Object System.Windows.Forms.PictureBox
$LabelFondBlanc = New-Object System.Windows.Forms.Label
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
#endregion Generated Form Objects

#----------------------------------------------
#Generated Event Script Blocks
#----------------------------------------------
#Provide Custom Code for events specified in PrimalForms.
$handler_pictureBox1_Click= 
{
#TODO: Place custom script here

}

$handler_button1_Click= 
{
#TODO: Place custom script here
ActionBt1_SelectFolder
}

$handler_button2_Click= 
{
#TODO: Place custom script here
ActionBt2_SelectSearch
}

$handler_btCreerSearchpro_Click= 
{
#TODO: Place custom script here
ActionBtCreerSearchpro
}

$handler_form1_Load= 
{
#TODO: Place custom script here

}

$handler_label4_Click= 
{
#TODO: Place custom script here

}

$handler_textBox1_TextChanged= 
{
#TODO: Place custom script here

}

$btQuitter_OnClick= 
{
#TODO: Place custom script here
ActionBtQuitter
}

$OnLoadForm_StateCorrection=
{#Correct the initial state of the form to prevent the .Net maximized form issue
	$form1.WindowState = $InitialFormWindowState
}

#----------------------------------------------
#region Generated Form Code
$form1.BackColor = [System.Drawing.Color]::FromArgb(255,105,105,105)
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 510
$System_Drawing_Size.Width = 690
$form1.ClientSize = $System_Drawing_Size
$form1.DataBindings.DefaultDataSourceUpdateMode = 0
$form1.FormBorderStyle = 5
$form1.Name = "form1"
$form1.Text = "Create Search.pro"
$form1.add_Load($handler_form1_Load)

#--- LABEL FICHIER SEARCH.PRO
$label1.BackColor = [System.Drawing.Color]::FromArgb(255,255,255,255)
$label1.DataBindings.DefaultDataSourceUpdateMode = 0
$label1.Font = New-Object System.Drawing.Font("Segoe UI",15.75,0,3,1)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 25
$System_Drawing_Point.Y = 280
$label1.Location = $System_Drawing_Point
$label1.Name = "label1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 40
$System_Drawing_Size.Width = 455
$label1.Size = $System_Drawing_Size
$label1.TabIndex = 8
$label1.Text = "Localisation du fichier Search.pro"
$form1.Controls.Add($label1)

#--- BOX CHEMIN+NOM FICHIER SEARCH.PRO
$textBox3.BackColor = [System.Drawing.Color]::FromArgb(255,255,255,255)
$textBox3.BorderStyle = 1
$textBox3.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 25
$System_Drawing_Point.Y = 320
$textBox3.Location = $System_Drawing_Point
$textBox3.Name = "textBox3"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 455
$textBox3.Size = $System_Drawing_Size
$textBox3.TabIndex = 7
$form1.Controls.Add($textBox3)

#--- LABEL REPRTOIRES A AJOUTER
$label2.AutoSize = $True
$label2.BackColor = [System.Drawing.Color]::FromArgb(255,255,255,255)
$label2.DataBindings.DefaultDataSourceUpdateMode = 0
$label2.Font = New-Object System.Drawing.Font("Segoe UI",15.75,0,3,1)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 25
$System_Drawing_Point.Y = 170
$label2.Location = $System_Drawing_Point
$label2.Name = "label2"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 40
$System_Drawing_Size.Width = 455
$label2.Size = $System_Drawing_Size
$label2.TabIndex = 1
$label2.Text = "Répertoire à ajouter au fichier search.pro"
$label2.add_Click($handler_label2_Click)
$form1.Controls.Add($label2)

#--- BOUTON AJOUTER LES REPERTOIRES
$btCreerAffaire.BackColor = [System.Drawing.Color]::FromArgb(255,255,165,0)

$btCreerAffaire.DataBindings.DefaultDataSourceUpdateMode = 0
$btCreerAffaire.FlatAppearance.BorderSize = 0
$btCreerAffaire.FlatStyle = 0
$btCreerAffaire.Font = New-Object System.Drawing.Font("Segoe UI",12,0,3,1)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 20
$System_Drawing_Point.Y = 380
$btCreerAffaire.Location = $System_Drawing_Point
$btCreerAffaire.Name = "btCreerAffaire"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 35
$System_Drawing_Size.Width = 640
$btCreerAffaire.Size = $System_Drawing_Size
$btCreerAffaire.TabIndex = 6
$btCreerAffaire.Text = "Ajouter les répertoires au fichier Search.pro"
$btCreerAffaire.UseVisualStyleBackColor = $False
$btCreerAffaire.add_Click($handler_btCreerSearchpro_Click)
$form1.Controls.Add($btCreerAffaire)

#--- BOUTON SELECTION DOSSIER
$BtSelectFolder.BackColor = [System.Drawing.Color]::FromArgb(255,143,188,139)

$BtSelectFolder.DataBindings.DefaultDataSourceUpdateMode = 0
$BtSelectFolder.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(255,143,188,139)
$BtSelectFolder.FlatAppearance.BorderSize = 0
$BtSelectFolder.FlatStyle = 0
$BtSelectFolder.Font = New-Object System.Drawing.Font("Segoe UI",12,0,3,1)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 500
$System_Drawing_Point.Y = 200
$BtSelectFolder.Location = $System_Drawing_Point
$BtSelectFolder.Name = "BtSelectFolder"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 30
$System_Drawing_Size.Width = 150
$BtSelectFolder.Size = $System_Drawing_Size
$BtSelectFolder.TabIndex = 2
$BtSelectFolder.Text = "Sélection Dossier"
$BtSelectFolder.UseVisualStyleBackColor = $False
$BtSelectFolder.add_Click($handler_button1_Click)
$form1.Controls.Add($BtSelectFolder)

#--- BOUTON SELECTION SEARCH.PRO
$BtSelectSearch.BackColor = [System.Drawing.Color]::FromArgb(255,143,188,139)

$BtSelectSearch.DataBindings.DefaultDataSourceUpdateMode = 0
$BtSelectSearch.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(255,143,188,139)
$BtSelectSearch.FlatAppearance.BorderSize = 0
$BtSelectSearch.FlatStyle = 0
$BtSelectSearch.Font = New-Object System.Drawing.Font("Segoe UI",12,0,3,1)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 500
$System_Drawing_Point.Y = 310
$BtSelectSearch.Location = $System_Drawing_Point
$BtSelectSearch.Name = "BtSelectSearch"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 30
$System_Drawing_Size.Width = 150
$BtSelectSearch.Size = $System_Drawing_Size
$BtSelectSearch.TabIndex = 2
$BtSelectSearch.Text = "Sélection Search.pro"
$BtSelectSearch.UseVisualStyleBackColor = $False
$BtSelectSearch.add_Click($handler_button2_Click)
$form1.Controls.Add($BtSelectSearch)

#--- BOX REPERTOIRE A AJOUTER AU SEARCH.PRO
$textBox1.BackColor = [System.Drawing.Color]::FromArgb(255,255,255,255)
$textBox1.BorderStyle = 1
$textBox1.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 25
$System_Drawing_Point.Y = 210
$textBox1.Location = $System_Drawing_Point
$textBox1.Name = "textBox1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 455
$textBox1.Size = $System_Drawing_Size
$textBox1.TabIndex = 0
$textBox1.add_TextChanged($handler_textBox1_TextChanged)
$form1.Controls.Add($textBox1)

#--- BOUTON QUITTER ---
$btQuitter.BackColor = [System.Drawing.Color]::FromArgb(255,180,180,180)

$btQuitter.DataBindings.DefaultDataSourceUpdateMode = 0
$btQuitter.FlatAppearance.BorderSize = 0
$btQuitter.FlatStyle = 0
$btQuitter.Font = New-Object System.Drawing.Font("Segoe UI",12,0,3,1)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 15
$System_Drawing_Point.Y = 460
$btQuitter.Location = $System_Drawing_Point
$btQuitter.Name = "btQuitter"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 35
$System_Drawing_Size.Width = 660
$btQuitter.Size = $System_Drawing_Size
$btQuitter.TabIndex = 3
$btQuitter.Text = "Quitter"
$btQuitter.UseVisualStyleBackColor = $False
$btQuitter.add_Click($btQuitter_OnClick)
$form1.Controls.Add($btQuitter)

#--- TEXTE TITRE
$label1_Titre.DataBindings.DefaultDataSourceUpdateMode = 0
$label1_Titre.Font = New-Object System.Drawing.Font("Segoe UI",27.75,0,3,1)
$label1_Titre.ForeColor = [System.Drawing.Color]::FromArgb(255,255,255,255)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 185
$System_Drawing_Point.Y = 72
$label1_Titre.Location = $System_Drawing_Point
$label1_Titre.Name = "label1_Titre"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 56
$System_Drawing_Size.Width = 505
$label1_Titre.Size = $System_Drawing_Size
$label1_Titre.TabIndex = 2
$label1_Titre.Text = "Création Fichier Search.pro"
$label1_Titre.add_Click($handler_label1_Click)
$form1.Controls.Add($label1_Titre)

#--- IMAGE LOGO EXCENT
$pictureBox1.DataBindings.DefaultDataSourceUpdateMode = 0

$pictureBox1.Image = [System.Drawing.Image]::FromFile($imgLogoeXcent)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 12
$pictureBox1.Location = $System_Drawing_Point
$pictureBox1.Name = "pictureBox1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 109
$System_Drawing_Size.Width = 167
$pictureBox1.Size = $System_Drawing_Size
$pictureBox1.TabIndex = 1
$pictureBox1.TabStop = $False
$pictureBox1.add_Click($handler_pictureBox1_Click)
$form1.Controls.Add($pictureBox1)

#--- CADRE FOND BLANC
$LabelFondBlanc.BackColor = [System.Drawing.Color]::FromArgb(255,255,255,255)
$LabelFondBlanc.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 15
$System_Drawing_Point.Y = 145
$LabelFondBlanc.Location = $System_Drawing_Point
$LabelFondBlanc.Name = "LabelFondBlanc"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 300
$System_Drawing_Size.Width = 660
$LabelFondBlanc.Size = $System_Drawing_Size
$LabelFondBlanc.TabIndex = 6
$LabelFondBlanc.UseCompatibleTextRendering = $True
$form1.Controls.Add($LabelFondBlanc)

#---
#endregion Generated Form Code

#Save the initial state of the form
$InitialFormWindowState = $form1.WindowState
#Init the OnLoad event to correct the initial state of the form
$form1.add_Load($OnLoadForm_StateCorrection)
#Show the Form
$form1.ShowDialog()| Out-Null

} #End Function


$ScriptPath = Get-ScriptDirectory
#Recup img
$imgLogoeXcent = $ScriptPath+"\"+$nLogoeXcent

GenerateForm