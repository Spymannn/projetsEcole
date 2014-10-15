Option Explicit
Dim MonShell,MonEnvironment,MonProc,MonReseau,listingReseau, res
Dim Drives,Printers,i, fichierTxt
Dim colItems, objWMIService, strComputer, objItem
strComputer = "." 
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colItems= objWMIService.ExecQuery("SELECT * FROM Win32_DiskDrive",,48)

Set MonShell=WScript.CreateObject("WScript.Shell")
Set MonEnvironment=MonShell.Environment("SYSTEM")
Set MonProc = MonShell.Environment("PROCESS")
Set MonReseau=WScript.CreateObject("Wscript.Network")
Set listingReseau = MonReseau.EnumNetworkDrives

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set fichierTxt = objFSO.CreateTextFile("reportingSamirHanini.txt",true,true)


Dim objItem2, colItems2

Dim Word, Tasks, Task


Set colItems2 = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")

fichierTxt.WriteLine "================================================"
fichierTxt.WriteLine "||            Projet Reporting                ||"
fichierTxt.WriteLine "||            Samir Hanini 2IA                ||"
fichierTxt.WriteLine "||       Reseau et systeme d'exploitation     ||"
fichierTxt.WriteLine "||               Juin 2014                    ||"
fichierTxt.WriteLine "================================================"

'================================================
'Infos générale
fichierTxt.WriteLine VBCRLF & "=================================================="
fichierTxt.WriteLine "||    Quelques informations generales           ||"
fichierTxt.WriteLine "=================================================="
fichierTxt.WriteLine "Le nom de l'utilisateur est: " & MonReseau.UserName
For Each objItem2 in colItems2
	fichierTxt.WriteLine "Nom de la machine : " & objItem2.CSName 
	fichierTxt.WriteLine "Fabricant: " & objItem2.Manufacturer 
	fichierTxt.WriteLine "Systeme d'exploitation: " & objItem2.Caption 
	fichierTxt.WriteLine "Version: " & objItem2.Version 
	fichierTxt.WriteLine "Service Pack: " & objItem2.CSDVersion 
	fichierTxt.WriteLine "CodeSet: " & objItem2.CodeSet 
	fichierTxt.WriteLine "Code du pays: " & objItem2.CountryCode 
	fichierTxt.WriteLine "Langage du systeme d'exploitation: " & objItem2.OSLanguage 
	fichierTxt.WriteLine "Zone d'heure local: " & objItem2.CurrentTimeZone 
	fichierTxt.WriteLine "Locale: " & objItem2.Locale 
	fichierTxt.WriteLine "Numero de serie: " & objItem2.SerialNumber 
	fichierTxt.WriteLine "Disque systeme: " & objItem2.SystemDrive 
	fichierTxt.WriteLine "Repertoire Windows: " & objItem2.WindowsDirectory 
Next
fichierTxt.WriteLine "Nombre de processeur " & MonProc("NUMBER_OF_PROCESSORS") &  " Type d'architecture " & MonProc("PROCESSOR_ARCHITECTURE")
fichierTxt.WriteLine "Identifiant du processeur de l'ordinateur : " & MonProc("PROCESSOR_IDENTIFIER")
fichierTxt.WriteLine "lettre de la partition principal : " & MonProc("HOMEDRIVE")
fichierTxt.WriteLine "Repertoire par defaut de l'utilisateur : " & MonProc("HOMEPATH")
fichierTxt.WriteLine "Chemin vers la racine du repertoire de l'OS : " & MonProc("WINDIR")
'=========================
'Disque durs
fichierTxt.WriteLine VBCRLF & "============================================"
fichierTxt.WriteLine "||               Lecteurs                 ||" 
fichierTxt.WriteLine "============================================"
Dim  colDrives, objDrive, DriveType
Set colDrives=objFSO.Drives
For Each objDrive in colDrives
	fichierTxt.WriteLine VBCRLF & "Lecteur " & objDrive.DriveLetter & ": " & VBCRLF & "----------"
	Select Case objDrive.DriveType
		Case 1 DriveType="Périphérique Amovible "
		Case 2 DriveType="Disque Dur "
		Case 3 DriveType="Disque reseau "
		Case 4 DriveType="CD/DVD "
		Case 5 DriveType="Disque RAM "
	End Select
	fichierTxt.WriteLine DriveType
	If objDrive.IsReady Then
		fichierTxt.WriteLine "Nom du volume : " & objDrive.VolumeName
		fichierTxt.WriteLine "Type de systeme de fichier : " & objDrive.FileSystem
		fichierTxt.WriteLine "Chemin d'acces : " & objDrive.Path
		fichierTxt.WriteLine Int(objDrive.TotalSize/(1024*1024)) & " Mo"
		fichierTxt.WriteLine "espace libre (Mo) " & Int(objDrive.FreeSpace/(1024*1024)) & "Mo"
		fichierTxt.WriteLine "Numero de serie : " & objDrive.SerialNumber
		fichierTxt.WriteLine "Nom du partage assigne a ce disque " & objDrive.ShareName
	End If
	
Next
'=========================================
'Informations sur le disque dur principal 
For Each objItem in colItems 
   IF objItem.Size <>0 Then 
		fichierTxt.WriteLine VBCRLF & "============================================================="
		fichierTxt.WriteLine "||        Informations sur le disque dur principal         ||"
		fichierTxt.WriteLine "============================================================="
		   fichierTxt.WriteLine "Bytes par secteur: " &  objitem.BytesPerSector
		   fichierTxt.WriteLine "Capacites : " &  Join(objItem.Capabilities, ",")
		   fichierTxt.WriteLine "Legende : " & objitem.Caption
		   fichierTxt.WriteLine "Identifiant de l'appareil : " & objitem.DeviceID
		   fichierTxt.WriteLine "Description : " & objitem.Description
		   fichierTxt.WriteLine "Fabricant: " & objitem.Manufacturer
		   fichierTxt.WriteLine "Type de media : " & objitem.MediaType
		   fichierTxt.WriteLine "Modele : " & objitem.Model
		   fichierTxt.WriteLine "Nom : " & objitem.Name
		   fichierTxt.WriteLine "Partitions  " & objitem.Partitions
		   fichierTxt.WriteLine "SCSIPort: " & objitem.SCSIPort
		   fichierTxt.WriteLine "Statuts : " & objitem.status
		   fichierTxt.WriteLine "Nom du systeme : " & objitem.Systemname
		   fichierTxt.WriteLine "Type d'interface : " & objItem.InterfaceType
  End If
Next
'==================================
'Listing des imprimantes
fichierTxt.WriteLine VBCRLF & "====================================="
fichierTxt.WriteLine "||           Imprimantes           ||"
fichierTxt.WriteLine "====================================="
Set Printers = MonReseau.EnumPrinterConnections()
For i = 0 to Printers.Count - 1 step 2
	fichierTxt.WriteLine "Imprimante " & Printers.item(i) & " = " & Printers.item(i+1)
Next
'==================================
'Listing des réseaux
fichierTxt.WriteLine VBCRLF & "=========================================="
fichierTxt.WriteLine "||         Listing des reseaux          ||"
fichierTxt.WriteLine "=========================================="
For Each res in listingReseau
	fichierTxt.WriteLine "Reseau : " & res
next

'=========================================================
'Informations BIOS
Dim colitemsBios, objItemBios
Set colItemsBios = objWMIService.ExecQuery( "Select * from Win32_BIOS where PrimaryBIOS = true", , 48 )
fichierTxt.WriteLine VBCRLF & "==========================================="
fichierTxt.WriteLine "||            Resume BIOS                || " & VBCRLF &"==========================================="
For Each objItemBios in colItemsBios
	fichierTxt.WriteLine "Nom BIOS : " & objItemBios.Name
	fichierTxt.WriteLine "Version : " & objItemBios.Version
	fichierTxt.WriteLine "Fabricant : " & objItemBios.Manufacturer
	fichierTxt.WriteLine "SMBIOS Version  :  " & objItemBios.SMBIOSBIOSVersion
Next
'=================================================
'Liste des groupes locaux utilisant le WMI 
Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_Group  Where LocalAccount = True")
	
fichierTxt.WriteLine VBCRLF & "===================================================================="
fichierTxt.WriteLine "||             Liste des groupes locaux utilisant WMI             ||" & VBCRLF & "===================================================================="
For Each objItem in colItems
    fichierTxt.WriteLine VBCRLF & "Groupe local : " & objItem.LocalAccount 
    fichierTxt.WriteLine "Nom: " & objItem.Name 
    fichierTxt.WriteLine "SID: " & objItem.SID 
    fichierTxt.WriteLine "Type de SID: " & objItem.SIDType 
    fichierTxt.WriteLine "Statut: " & objItem.Status 
Next

'===========================================
'Liste des applications en cours
Set Word = CreateObject("Word.Application")
Set Tasks = Word.Tasks

fichierTxt.WriteLine VBCRLF & "==============================================================="
fichierTxt.WriteLine "||          Liste des applications en cours                  ||" & VBCRLF & "==============================================================="
For Each Task in Tasks
	If Task.Visible Then
		fichierTxt.WriteLine Task.Name
	End if
Next

'================================================
'Liste des logiciels lancés au démarrage
Dim colStartupCommands, objStartupCommand
Set colStartupCommands = objWMIService.ExecQuery ("Select * from Win32_StartupCommand")

fichierTxt.WriteLine VBCRLF & "======================================================================"
fichierTxt.WriteLine "||     Liste des logiciels lances au demarrage de la machine        ||"
fichierTxt.WriteLine "======================================================================"
For Each objStartupCommand in colStartupCommands
    fichierTxt.WriteLine VBCRLF & "Commande : " & objStartupCommand.Command 
    fichierTxt.WriteLine "Description : " & objStartupCommand.Description 
    fichierTxt.WriteLine "Chemin d'acces : " & objStartupCommand.Location 
    fichierTxt.WriteLine "Nom: " & objStartupCommand.Name 
    fichierTxt.WriteLine "Utilisateur : " & objStartupCommand.User
Next
'===========================================
'Liste des applications installées sur la machine
Dim colSoftware, objSoftware
Set colSoftware = objWMIService.ExecQuery("SELECT * FROM Win32_Product")   

fichierTxt.WriteLine VBCRLF & "======================================================"
fichierTxt.WriteLine "||        Liste des applications installees         ||" 
fichierTxt.WriteLine "======================================================" & VBCRLF & VBCRLF

If colSoftware.Count > 0 Then
    For Each objSoftware in colSoftware
        fichierTxt.WriteLine VBCRLF & "Nom : " & objSoftware.Caption 
		fichierTxt.WriteLine "Version : " & objSoftware.Version
		fichierTxt.WriteLine "Date d'installation : " & objSoftware.InstallDate
		fichierTxt.WriteLine "Chemin d'acces : " & objSoftware.InstallSource
		fichierTxt.WriteLine "Package locale : " & objSoftware.LocalPackage
    Next
End If
'================================================
'Version VB script utilisé
fichierTxt.WriteLine VBCRLF & "=============================================="
fichierTxt.WriteLine "||         Version script et WSH            ||"
fichierTxt.WriteLine "=============================================="
fichierTxt.WriteLine "Version de " & ScriptEngine & " : " & ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion 
fichierTxt.WriteLine "(build " & ScriptEngineBuildVersion & ")" 
fichierTxt.WriteLine "Version de WSH  : " & WScript.Version
'==================================================
'partie qui permet de lancer directement le fichier texte avec notepad
Dim shell , Chemin 
Set shell = CreateObject("WScript.Shell")
Chemin = "reportingSamirHanini.txt"
shell.Run "Notepad " & Chemin
WScript.Quit


