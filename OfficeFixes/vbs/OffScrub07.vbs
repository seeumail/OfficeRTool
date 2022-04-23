'=======================================================================================================
' Name: OffScrub07.vbs
' Author: Microsoft Customer Support Services
' Copyright (c) 2008, Microsoft Corporation
' Script to remove (scrub) Office 2007 products
'=======================================================================================================
Option Explicit

Const VERSION       = "1.16"
Const HKCR          = &H80000000
Const HKCU          = &H80000001
Const HKLM          = &H80000002
Const HKU           = &H80000003
Const FOR_WRITING   = 2
Const PRODLEN       = 13
Const OFFICEID      = "0000000FF1CE}"
Const COMPPERMANENT = "00000000000000000000000000000000"
Const REG_ARP       = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"

Dim oFso, oMsi, oReg, oWShell, oWmiLocal
Dim ComputerItem, Item, LogStream, TmpKey
Dim arrInstalledSKUs, arrRemoveSKUs, arrKeepSKUs, arrTmpSKUs
Dim arrDeleteFiles, arrDeleteFolders, arrMseFolders
Dim b64
Dim sErr, sSkuInstalledList, sSkuRemoveList, sDefault, sWinDir, sMode, sApps
Dim sAppData, sTemp, sScrubDir, sProgramFiles, sProgramFilesX86, sCommonProgramFiles, sAllusersProfile

'=======================================================================================================
'Main
'=======================================================================================================
'Configure defaults
Dim sLogDir : sLogDir = ""
Dim sMoveMessage: sMoveMessage = ""
Dim bRemoveOSE      : bRemoveOSE = False
Dim bRemoveAll      : bRemoveAll = False
Dim bSimulate       : bSimulate = False
Dim bQuiet          : bQuiet = False
Dim bNoCancel       : bNoCancel = False
'CAUTION! -> "bForce" will kill running applications which can result in data loss! <- CAUTION
Dim bForce          : bForce = False
'CAUTION! -> "bForce" will kill running applications which can result in data loss! <- CAUTION
Dim bLogInitialized : bLogInitialized = False
Dim bBypass_Stage1  : bBypass_Stage1 = False 'Component Detection
Dim bBypass_Stage2  : bBypass_Stage2 = False 'Setup
Dim bBypass_Stage3  : bBypass_Stage3 = False 'Msiexec
Dim bBypass_Stage4  : bBypass_Stage4 = False 'CleanUp

sApps = "\communicator.exe"

'Create required objects
Set oWmiLocal   = GetObject("winmgmts:\\.\root\cimv2")
Set oWShell     = CreateObject("Wscript.Shell")
Set oFso        = CreateObject("Scripting.FileSystemObject")
Set oMsi        = CreateObject("WindowsInstaller.Installer")
Set oReg        = GetObject("winmgmts:\\.\root\default:StdRegProv")

'Ensure CScript as engine
If Not UCase(Mid(Wscript.FullName, Len(Wscript.Path) + 2, 1)) = "C" Then RelaunchAsCScript

'Get environment path info
sAppData            = oWShell.ExpandEnvironmentStrings("%appdata%")
sTemp               = oWShell.ExpandEnvironmentStrings("%temp%")
sAllUsersProfile    = oWShell.ExpandEnvironmentStrings("%allusersprofile%")
sProgramFiles       = oWShell.ExpandEnvironmentStrings("%programfiles%")
sCommonProgramFiles = oWShell.ExpandEnvironmentStrings("%commonprogramfiles%")
sWinDir             = oWShell.ExpandEnvironmentStrings("%windir%")
sScrubDir           = sTemp & "\OffScrub07"

'Create the temp folder
If Not oFso.FolderExists(sScrubDir) Then oFso.CreateFolder sScrubDir

'Set the default logging directory
sLogDir = sScrubDir

'Detect if we're running on a 64 bit OS
Set ComputerItem = oWmiLocal.ExecQuery("Select * from Win32_ComputerSystem")
For Each Item In ComputerItem
    b64 = Instr(Left(Item.SystemType,3),"64") > 0
    'Log "64 bit OS: " & b64 & " -> " & Item.SystemType
Next
If b64 Then sProgramFilesX86 = oWShell.ExpandEnvironmentStrings("%programfiles(x86)%")

'Call the command line parser
ParseCmdLine

If Not CheckRegPermissions Then
    Log "Insufficient registry access permissions - exiting"
    'Undo temporary entries created in ARP
    TmpKeyCleanUp
    wscript.quit 
End If

'-------------------
'Stage # 0 - Basics |
'-------------------
'Build a list with installed/registered Office 2007 products
Log vbCrLf & Now & " - Stage # 0 " & chr(34) & "Basics" & chr(34)
FindInstalledO12Products
If Len(sSkuInstalledList)>0 Then Log "Found registered product(s): " & Left(sSkuInstalledList,Len(sSkuInstalledList)-1)

'Validate the list of products we got from the command line if applicable
ValidateRemoveSkuList
sMode = "Selected Office 2007 products"
If Not IsArray(arrRemoveSKUs) Then sMode = "Orphaned Office 2007 products"
If bRemoveAll Then sMode = "All Office 2007 products"
Log "Final removal mode: " & sMode
If IsArray(arrRemoveSKUs) Then Log "List of configuration product(s) to remove: " & Join(arrRemoveSKUs,",")
Log "Remove OSE service: " & bRemoveOSE
If bSimulate Then Log "*************************************************************************"
If bSimulate Then Log "*                          PREVIEW MODE                                 *"  
If bSimulate Then Log "* All uninstall and delete operations will only be logged not executed! *"
If bSimulate Then Log "*************************************************************************"

'Cache .msi files
If IsArray(arrRemoveSKUs) Then CacheMsiFiles

'--------------------------------
'Stage # 1 - Component Detection |
'--------------------------------
If Not bBypass_Stage1 Then
    Log vbCrLf & Now & " - Stage # 1 " & chr(34) & "Component Detection" & chr(34)

    'Build a list with files which are installed/registered to a product that's going to be removed
    Log vbCrLf & "Prepare for CleanUp stages."
    Log "Searching for removable files. This can take several minutes."
    BuildFileList : Log "Done"
Else
    Log vbCrLf & Now & " - Skipping Stage # 1 " & chr(34) & "Component Detection" & chr(34) & " because bypass was requested."
End If

'Kill all running Office applications
If bForce Then KillApps

'----------------------
'Stage # 2 - Setup.exe |
'----------------------
If Not bBypass_Stage2 Then
    Log vbCrLf & Now & " - Stage # 2 " & chr(34) & "Setup.exe" & chr(34)
    SetupExeRemoval
Else
    Log vbCrLf & Now & " - Skipping Stage # 2 " & chr(34) & "Setup.exe" & chr(34) & " because bypass was requested."
End If

'------------------------
'Stage # 3 - Msiexec.exe |
'------------------------
If Not bBypass_Stage3 Then
    Log vbCrLf & Now & " - Stage # 3 " & chr(34) & "Msiexec.exe" & chr(34)
    MsiexecRemoval
Else
    Log vbCrLf & Now & " - Skipping Stage # 3 " & chr(34) & "Msiexec.exe" & chr(34) & " because bypass was requested."
End If

'--------------------
'Stage # 4 - CleanUp |
'--------------------
'Removal of files and registry settings
If Not bBypass_Stage4 Then
    Log vbCrLf & Now & " - Stage # 4 " & chr(34) & "CleanUp" & chr(34)
    'Office Source Engine
    If bRemoveOSE Then RemoveOSE

    'Local Installation Source (MSOCache)
    WipeLIS
    
    'Obsolete files
    If bRemoveAll Then 
        FileWipeAll 
    Else 
        FileWipeIndividual
    End If
    
    'Empty Folders
    DeleteEmptyFolders
    
    'Restore Explorer if needed
    If bForce Then RestoreExplorer
    
    'Registry data
    RegWipe
    
    'Wipe orphaned files from Windows Installer cache
    MsiClearOrphanedFiles
    
    'Temporary .msi files in scrubcache
    DeleteMsiScrubCache
    
    'Temporary files from file move operations
    DelScrubTmp
    
Else
    Log vbCrLf & Now & " - Skipping Stage # 5 " & chr(34) & "CleanUp" & chr(34) & " because bypass was requested."
End If

If Not sMoveMessage = "" Then Log vbCrLf & "Please remove this folder after next reboot: " & sMoveMessage

'THE END
Log vbCrLf & "End removal: " & Now & vbCrLf
'=======================================================================================================
'=======================================================================================================

'Stage 0 - 4 Subroutines
'=======================================================================================================

'Office 2007 configuration products are listed with their configuration product name in the "Uninstall" key
'To identify an Office 2007 configuration product all of these condiditions have to be met:
' - "SystemComponent" entry exists with a value of "0" 
' - "PackageIds" entry exists and is not empty

' - "DisplayVersion" exists and the 3 leftmost digits are "12."
Sub FindInstalledO12Products
    Dim ArpItem
    Dim hDefKey, sSubKeyName, sCurKey, sName, sValue, sConfigName, sLcid
    Dim arrKeys, arrValues, arrMultiSzValues
    Dim bSystemComponent0, bPackageIDs, bDisplayVersion

    If Len(sSkuInstalledList) > 0 Then Exit Sub 'Already done from InputBox prompt
    sSubKeyName = REG_ARP
    
    'Locate standalone Office 2007 products that have no configuration product entry and create a
    'temporary configuration entry
    ReDim arrTmpSKUs(-1)
    If RegEnumKey(HKLM,sSubKeyName,arrKeys) Then
        For Each ArpItem in arrKeys
            If UCase(Right(ArpItem,PRODLEN))=OFFICEID AND Mid(ArpItem,4,2)="12" Then
                sCurKey = sSubKeyName & ArpItem & "\"
                bSystemComponent0 = RegReadValue(HKLM,sCurKey,"SystemComponent",sValue,"REG_DWORD") AND sValue = "0"
                If bSystemComponent0 OR Not RegReadValue(HKLM,sCurKey,"SystemComponent",sValue,"REG_DWORD") Then
                    RegReadValue HKLM,sCurKey,"DisplayVersion",sValue,"REG_SZ"
                    Redim arrMultiSzValues(0)
                    sConfigName = GetProductID(Mid(ArpItem,11,4)) & "_" & CInt("&h" & Mid(ArpItem,16,4))
                    ReDim Preserve arrTmpSKUs(UBound(arrTmpSKUs)+1)
                    arrTmpSKUs(UBound(arrTmpSKUs)) = sConfigName
                    oReg.CreateKey HKLM,REG_ARP & "\" & sConfigName
                    arrMultiSzValues(0) = sConfigName
                    oReg.SetMultiStringValue HKLM,REG_ARP & "\" & sConfigName,"PackageIds",arrMultiSzValues
                    arrMultiSzValues(0) = ArpItem
                    oReg.SetMultiStringValue HKLM,REG_ARP & "\" & sConfigName,"ProductCodes",arrMultiSzValues
                    oReg.SetStringValue HKLM,REG_ARP & "\" & sConfigName,"DisplayVersion",sValue
                    oReg.SetDWordValue HKLM,REG_ARP & "\" & sConfigName,"SystemComponent",0
                End If
            End If
        Next 'ArpItem
    End If 'RegEnumKey
    
    'Find the configuration products
    If RegEnumKey(HKLM,sSubKeyName,arrKeys)Then
        For Each ArpItem in arrKeys
            sCurKey = sSubKeyName & ArpItem & "\"
            bSystemComponent0 = RegReadValue(HKLM,sCurKey,"SystemComponent",sValue,"REG_DWORD") AND sValue = "0"
            bPackageIDs = RegReadValue(HKLM,sCurKey,"PackageIds",sValue,"REG_MULTI_SZ")
            bDisplayVersion = RegReadValue(HKLM,sCurKey,"DisplayVersion",sValue,"REG_SZ") AND (Left(sValue,3) = "12.")
            If (bSystemComponent0 AND bPackageIDs AND bDisplayVersion) Then _
                sSkuInstalledList = sSkuInstalledList & UCase(ArpItem) & ","
        Next 'ArpItem
    End If 'RegEnumKey
    If Len(sSkuInstalledList) > 0 Then arrInstalledSKUs = Split(Left(sSkuInstalledList,Len(sSkuInstalledList)-1),",")
End Sub 'FindInstalledO12Products
'=======================================================================================================

'Create clean list of Products to remove.
'Strip of bad & empty contents
Sub ValidateRemoveSkuList
    Dim Sku, sSkuKeepList
    Dim iPos
    
    If bRemoveAll Then
        If Len(sSkuInstalledList) > 0 Then 
            sSkuInstalledList = Left(sSkuInstalledList,Len(sSkuInstalledList)-1)
            arrRemoveSKUs = Split(sSkuInstalledList,",")
            sSkuInstalledList = sSkuInstalledList & ","
        Else
            Set arrRemoveSKUs = Nothing
        End If
    Else
        'Ensure to have a string with no unexpected contents
        sSkuRemoveList = Replace(sSkuRemoveList," ","")
        sSkuRemoveList = Replace(sSkuRemoveList,Chr(34),"")
        While InStr(sSkuRemoveList,",,")>0
            sSkuRemoveList = Replace(sSkuRemoveList,",,",",")
        Wend
        arrRemoveSKUs = Split(UCase(sSkuRemoveList),",")
        sSkuKeepList = "," & sSkuInstalledList & "OSE,"
        sSkuRemoveList = ""
        'Compare the list from the Cmd against the actually installed list of Office 2007 products
        For Each Sku in arrRemoveSKUs
            If Len(Sku)>0 AND InStr(sSkuKeepList,"," & Sku & ",") > 0 Then
                sSkuKeepList = Replace(sSkuKeepList,Sku & ",","")
                sSkuRemoveList = sSkuRemoveList & Sku & ","
            End If 'iPos > 0
        Next 'Sku
        If Right(sSkuKeepList,4)="OSE," Then sSkuKeepList = Left(sSkuKeepList,Len(sSkuKeepList)-4)
        sSkuKeepList = Right(sSkuKeepList,Len(sSkuKeepList)-1)
        bRemoveAll = (sSkuKeepList = "")
        If Not bRemoveAll Then arrKeepSKUs = Split(Mid(sSkuKeepList,1,Len(sSkuKeepList)-1),",")
        If Len(sSkuRemoveList) > 0 Then 
            sSkuRemoveList = "," & sSkuRemoveList
            If InStr(sSkuRemoveList,",OSE,")>0 Then 
                sSkuRemoveList = Replace(sSkuRemoveList,",OSE,",",")
                bRemoveOSE = True
            End If
            sSkuRemoveList = Right(sSkuRemoveList,Len(sSkuRemoveList)-1)
            'Recheck if there are products to remove in the list after the OSE chcek
            If Len(sSkuRemoveList) > 0 Then
                arrRemoveSKUs = Split(Left(sSkuRemoveList,Len(sSkuRemoveList)-1),",")
            Else
                arrRemoveSKus = Nothing
            End If
        Else
            Set arrRemoveSKus = Nothing
        End If
    End If 'bRemoveAll
    If bRemoveAll AND Not bRemoveOSE Then CheckRemoveOSE
End Sub 'ValidateRemoveSkuList
'=======================================================================================================

'Check if OSE service can be scrubbed
Sub CheckRemoveOSE
    Const O11 = "6000-11D3-8CFE-0150048383C9}"
    Dim Product
    
    For Each Product in oMsi.Products
        If UCase(Right(Product,28)) = O11 Then 
            bRemoveOSE = False
            Exit Sub
        End If
    Next 'Product
    If UCase(Right(Product,PRODLEN))=OFFICEID AND Mid(Product,4,2)="14" Then
        'Found Office 14 Product. Set flag to not remove the OSE service
        bRemoveOSE = False
        Exit Sub
    End If
    bRemoveOSE = True
End Sub 'CheckRemoveOSE
'=======================================================================================================

'Cache .msi files for products that will be removed in case they are needed for later file detection
Sub CacheMsiFiles
    Dim Product
    Dim sMsiFile
    
    On Error Resume Next
    Log "Cache .msi files to temporary Scrub folder:"
    'Cache the files
    For Each Product in oMsi.Products
        If (Right(Product,PRODLEN) = OFFICEID  AND Mid(Product,4,2)="12") AND (bRemoveAll OR CheckDelete(Product))Then
            CheckError "CacheMsiFiles"
            sMsiFile = oMsi.ProductInfo(Product,"LocalPackage") : CheckError "CacheMsiFiles"
            Log "File backup: " & Product & ".msi"
            If oFso.FileExists(sMsiFile) Then oFso.CopyFile sMsiFile,sScrubDir & "\" & Product & ".msi",True
            CheckError "CacheMsiFiles"
        End If 'Right(Product,PRODLEN) = OFFICEID ...
    Next 'Product
End Sub 'CacheMsiFiles
'=======================================================================================================

'Build a list of all files that will be deleted
Sub BuildFileList
    Const MSIOPENDATABASEREADONLY   = 0

    Dim FileList, oDic, oFolderDic, ComponentID, CompClient, Record, qView, MsiDb
    Dim sQuery, sSubKeyName, sPath, sFile, sMsiFile, sCompClient, sComponent
    Dim bRemoveComponent
    Dim i, iProgress, iCompCnt
    
    'Logfile
    Set FileList = oFso.OpenTextFile(sScrubDir & "\FileList.txt",FOR_WRITING,True,True)
    
    On Error Resume Next 'Not optional here. Required for inline error handler
    Set oDic = CreateObject("Scripting.Dictionary")
    Set oFolderDic = CreateObject("Scripting.Dictionary")
    iCompCnt = oMsi.Components.Count
    'Enum all Components
    For Each ComponentID In oMsi.Components
        'Progress bar
        i = i + 1
        If iProgress < (i / iCompCnt) * 100 Then 
            wscript.stdout.write "." : LogStream.Write "."
            iProgress = iProgress + 1
        End If
        bRemoveComponent = False
        'Check if all ComponentClients will be removed
        For Each CompClient In oMsi.ComponentClients(ComponentID)
            bRemoveComponent = Right(CompClient,PRODLEN)=OFFICEID AND Mid(CompClient,4,2)="12" AND CheckDelete(CompClient)
            If Not bRemoveComponent Then Exit For
            'In "force" mode all components will be removed regardless of msidbComponentAttributesPermanent flag.
            'Default is to honour the msidbComponentAttributesPermanent attribute and keep the files
            If Not bForce Then
                sSubKeyName = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\"
                If RegValExists(HKLM,sSubKeyName & GetCompressedGuid(CompClient),COMPPERMANENT) Then
                    bRemoveComponent = False
                    Exit For
                End If
            End If 'bForce
            sCompClient = CompClient
        Next 'CompClient

        If bRemoveComponent Then
            Err.Clear
            'Get the component path
            sPath = oMsi.ComponentPath(sCompClient,ComponentID)
            If oFso.FileExists(sPath) Then
                sPath = Left(sPath,InStrRev(sPath,"\")-1)
                If Not oFolderDic.Exists(sPath) Then oFolderDic.Add sPath,sPath
                'Get the .msi file
                If oFso.FileExists(sScrubDir & "\" & sCompClient & ".msi") Then
                    sMsiFile = sScrubDir & "\" & sCompClient & ".msi"
                Else
                    sMsiFile = oMsi.ProductInfo(sCompClient,"LocalPackage")
                End If
                Set MsiDb = oMsi.OpenDatabase(sMsiFile,MSIOPENDATABASEREADONLY)
                
                If Err = 0 Then
                    'Get the component name from the 'Component' table
                    sQuery = "SELECT `Component`,`ComponentId` FROM Component WHERE `ComponentId` = '" & ComponentID &"'"
                    Set qView = MsiDb.OpenView(sQuery) : qView.Execute
                    Set Record = qView.Fetch()
                    If Not Record Is Nothing Then sComponent = Record.Stringdata(1)

                    'Get filenames from the 'File' table
                    sQuery = "SELECT `Component_`,`FileName` FROM File WHERE `Component_` = '" & sComponent &"'"
                    Set qView = MsiDb.OpenView(sQuery) : qView.Execute
                    Set Record = qView.Fetch()
                    Do Until Record Is Nothing
                        'Read the filename
                        sFile = Record.StringData(2)
                        If InStr(sFile,"|") > 0 Then sFile = Mid(sFile,InStr(sFile,"|")+1,Len(sFile))
                        sFile = sPath & "\" & sFile
                        If Not oDic.Exists(sFile) Then 
                            oDic.Add sFile,sFile
                            FileList.WriteLine sFile
                            If LCase(Right(sFile,4))=".exe" Then sApps = sApps & ";" & sFile
                        End If
                        Set Record = qView.Fetch()
                    Loop
                    Set Record = Nothing
                    qView.Close
                    Set qView = Nothing
                End If 'Err = 0
            End If 'FileExists(sPath)
        End If 'bRemoveComponent
    Next 'ComponentID
    sApps = sApps & ";"
    If Not oFolderDic.Count = 0 Then arrDeleteFolders = oFolderDic.Keys Else Set arrDeleteFolders = Nothing
    If Not oDic.Count = 0 Then arrDeleteFiles = oDic.Keys Else Set arrDeleteFiles = Nothing
End Sub 'BuildFileList
'=======================================================================================================

'Try to remove the products by calling setup.exe
Sub SetupExeRemoval
    Dim OseService, Service, TextStream
    Dim iSetupCnt, RetVal
    Dim Sku, sConfigFile, sUninstallCmd, sCatalyst, sDll, sDisplayLevel, sNoCancel

    iSetupCnt = 0
    If Not IsArray(arrRemoveSKUs) Then 
        Log "Nothing to remove for setup."
        Exit Sub
    End If
    
    'Ensure that the OSE service is *installed, *not disabled, *running under System context.
    'If validation fails exit out of this sub.
    Set OseService = oWmiLocal.Execquery("Select * From Win32_Service Where Name='ose'")
    If OseService.Count = 0 Then Exit Sub
    For Each Service in OseService
        If (Service.StartMode = "Disabled") AND (Not Service.ChangeStartMode("Manual")=0) Then Exit Sub
        If (Not Service.StartName = "LocalSystem") AND (Service.Change( , , , , , , "LocalSystem", "")) Then Exit Sub
    Next 'Service
    
    For Each Sku in arrRemoveSKUs
        'Create an "unattended" config.xml file for uninstall
        If bQuiet Then sDisplayLevel = "None" Else sDisplayLevel="Basic"
        If bNoCancel Then sNoCancel="Yes" Else sNoCancel="No"
        Set TextStream = oFso.OpenTextFile(sScrubDir & "\config.xml",FOR_WRITING,True,True)
        TextStream.Writeline "<Configuration Product=""" & Sku & """>"
        TextStream.Writeline "<Display Level=""" & sDisplayLevel & """ CompletionNotice=""No"" SuppressModal=""Yes"" NoCancel=""" & sNoCancel & """ AcceptEula=""Yes"" />"
        TextStream.Writeline "<Logging Type=""Verbose"" Path=""" & sLogDir & """ Template=""Microsoft Office " & Sku & " Setup(*).txt"" />"
        TextStream.Writeline "<Setting Id=""SETUP_REBOOT"" Value=""Never"" />"
        TextStream.Writeline "</Configuration>"
        TextStream.Close
        Set TextStream = Nothing
        
        'Ensure path to setup.exe is valid to prevent errors
        RetVal = RegReadValue(HKLM,REG_ARP & Sku,"UninstallString",sCatalyst,"REG_SZ")
        If RetVal Then
            If InStr(LCase(sCatalyst),"/dll")>0 Then sDll = Right(sCatalyst,Len(sCatalyst)-InStr(LCase(sCatalyst),"/dll")+2)
            sCatalyst = Trim(Replace(Left(sCatalyst,InStr(sCatalyst,"/")-2),Chr(34),""))
            If oFso.FileExists(sCatalyst) Then
                sUninstallCmd = Chr(34) & sCatalyst & Chr(34) & " /uninstall " & Sku & " /config " & Chr(34) & sScrubDir & "\config.xml" & Chr(34) & sDll 
                iSetupCnt = iSetupCnt + 1
                Log "Calling setup.exe to remove " & Sku '& vbCrLf & sUninstallCmd 
                On Error Resume Next
                If Not bSimulate Then RetVal = oWShell.Run(sUninstallCmd,0,True) : CheckError "CacheMsiFiles"
                On Error Goto 0
                Log "Removal of " & Sku & " returned: " & SetupExeRetVal(Retval) & " (" & RetVal & ")"
            Else
                Log "Error: Office 2007 setup.exe appears to be missing"
            End If 'RetVal = 0) AND oFso.FileExists
        End If 'RetVal
    Next 'Sku
    If iSetupCnt = 0 Then Log "Nothing to remove for setup."
End Sub 'SetupExeRemoval
'=======================================================================================================

'Invoke msiexec to remove individual .MSI packages
Sub MsiexecRemoval
    Const MSIINSTALLSTATEABSENT = 2
    
    Const MSIUILEVELNONE = 2
    Const MSIUILEVELBASIC = 3 'Simple progress and error handling. 
    Const MSIUILEVELHIDECANCEL = 32 ' shows progress dialog boxes but does not display a Cancel button
    Const MSIUILEVELPROGRESSONLY = 64 'displays progress dialog boxes but does not display any modal dialog boxes or error dialog boxes. 

    Dim Product
    Dim i
    
    'Check registered products
    'O12 does only support per machine installation so it's sufficient to use Installer.Products
    i = 0
    If bQuiet Then
        oMsi.UILevel = MSIUILEVELNONE
    Else
        If bNoCancel Then
            oMsi.UILevel = MSIUILEVELBASIC + MSIUILEVELHIDECANCEL + MSIUILEVELPROGRESSONLY
        Else
            oMsi.UILevel = MSIUILEVELBASIC + MSIUILEVELPROGRESSONLY
        End If
    End If
    For Each Product in oMsi.Products
        If (Right(Product,PRODLEN) = OFFICEID AND Mid(Product,4,2)="12") AND (bRemoveAll OR CheckDelete(Product))Then
            i = i + 1 
            Log "Calling msiexec.exe to remove " & Product
            oMsi.EnableLog "voicewarmup", sLogDir & "\Uninstall_" & Product & ".log"
            On Error Resume Next
            If Not bSimulate Then oMsi.ConfigureProduct Product,0,MSIINSTALLSTATEABSENT
            On Error Goto 0
        End If 'Right(Product,PRODLEN) = OFFICEID
    Next 'Product
    If i = 0 Then Log "Nothing to remove for msiexec"
End Sub 'MsiexecRemoval
'=======================================================================================================

Sub RemoveOSE
    Dim OseService, Service, Processes, Process
    
    On Error Resume Next
    
    Log "OSE CleanUp:"
    'Invoke the subroutine to delete a service
    DeleteService ("ose")
    'Delete the folder
    DeleteFolder sCommonProgramFiles & "\Microsoft Shared\Source Engine"
    'Delete the registration
    RegDeleteKey HKLM,"SYSTEM\CurrentControlSet\Services\ose"

End Sub 'RemoveOSE
'=======================================================================================================

'File cleanup operations for the Local Installation Source (MSOCache)
Sub WipeLIS
    Const LISROOT = "MSOCache\All Users\"
    Dim LogicalDisks, Disk, Folder, SubFolder, MseFolder, File, Files
    Dim arrSubFolders
    Dim sFolder
    Dim bRemoveFolder
    
    'On Error Resume Next
    Log "LIS CleanUp:"
    'Search all hard disks
    Set LogicalDisks = oWmiLocal.ExecQuery("Select * from Win32_LogicalDisk")
    For Each Disk in LogicalDisks
        If Disk.DriveType = 3 Then
            If oFso.FolderExists(Disk.DeviceID & "\" & LISROOT)Then
                If Err <> 0 Then 
                    CheckError  "WipeLIS"
                    Exit Sub
                End If
                Set Folder = oFso.GetFolder(Disk.DeviceID & "\" & LISROOT)
                For Each Subfolder in Folder.Subfolders
                    If bRemoveAll Then 
                        If  (Mid(Subfolder.Name,26,PRODLEN) = OFFICEID AND Mid(SubFolder.Name,4,2)="12") OR _
                            LCase(Right(Subfolder.Name,7)) = "12.data" Then DeleteFolder Subfolder.Path
                    Else
                        If  (Mid(Subfolder.Name,26,PRODLEN) = OFFICEID AND Mid(SubFolder.Name,4,2)="12") AND _
                            CheckDelete(UCase(Left(Subfolder.Name,38))) AND _
                            UCase(Right(Subfolder,1))= UCase(Left(Disk.DeviceID,1))Then DeleteFolder Subfolder.Path
                    End If
                Next 'Subfolder
                If (Folder.Subfolders.Count = 0) AND (Folder.Files.Count = 0) Then 
                    sFolder = Folder.Path
                    Set Folder = Nothing
                    SmartDeleteFolder sFolder
                End If
            End If 'oFso.FolderExists
        End If 'Disk.DriveType = 3
    Next 'Disk
    
    'MSECache
    If EnumFolders(sProgramFiles,arrSubFolders) Then
        For Each SubFolder in arrSubFolders
            If UCase(Right(SubFolder,9))="\MSECACHE" Then
                ReDim arrMseFolders(-1)
                Set Folder = oFso.GetFolder(SubFolder)
                GetFolderStructure Folder
                For Each MseFolder in arrMseFolders
                    If oFso.FolderExists(MseFolder) Then
                        bRemoveFolder = False
                        Set Folder = oFso.GetFolder(MseFolder)
                        Set Files = Folder.Files
                        For Each File in Files
                            If (LCase(Right(File.Name,4))=".msi") Then
                                If CheckDelete(ProductCode(File.Path)) Then 
                                    bRemoveFolder = True
                                    Exit For
                                End If 'CheckDelete
                            End If
                        Next 'File
                        Set Files = Nothing
                        Set Folder = Nothing
                        If bRemoveFolder Then SmartDeleteFolder MseFolder
                    End If 'oFso.FolderExists(MseFolder)
                Next 'MseFolder
            End If
        Next 'SubFolder
    End If 'oFso.FolderExists
End Sub 'WipeLis
'=======================================================================================================

'Wipe files and folders as documented in KB 928218
Sub FileWipeAll
    Dim sFile, sAppdata, sFolder
    Dim File, Files, Folder, Subfolder, OSPPsvc, Service
    
    'On Error Resume Next
    
    'Run the individual filewipe first
    FileWipeIndividual
    DeleteFolder sCommonProgramFiles & "\Microsoft Shared\Office12"
    DeleteFolder sProgramFiles & "\Microsoft Office\Office12"
    DeleteFile sAllUsersProfile & "\Application Data\Microsoft\Office\Data\opa12.dat"
    'Delete files that should be backed up before deleting them
    CopyAndDeleteFile sAppdata & "\Microsoft\Templates\Normal.dotm"
    CopyAndDeleteFile sAppdata & "\Microsoft\Templates\Normalemail.dotm"
    sFolder = sAppdata & "\microsoft\document building blocks"
    If oFso.FolderExists(sFolder) Then 
        Set Folder = oFso.GetFolder(sFolder)
        For Each Subfolder In Folder.Subfolders
            If oFso.FileExists(Subfolder & "\blocks.dotx") Then CopyAndDeleteFile Subfolder & "\blocks.dotx"
        Next 'Subfolder
        Set Folder = Nothing
    End If
    'Cleanup %temp% folder
    'Don't run this if the current log folder points to %temp%
    If Not sLogDir = sTemp Then
        Set Folder = oFso.GetFolder(sTemp)
        Set Files = Folder.Files
        For Each File in Files
            DeleteFile File.Path
        Next 'File
        For Each Subfolder in Folder.Subfolders
            If Not LCase(Subfolder.Name) = "offscrub07" Then DeleteFolder Subfolder.Path
        Next 'Subfolder
    End If 'Not sLogDir = sTemp
End Sub 'FileWipeAll
'=======================================================================================================

'Wipe individual files & folders related to SKU's that are no longer installed
Sub FileWipeIndividual
    Dim LogicalDisks, Disk
    Dim File, Files, XmlFile, scFiles, oFile, Folder, SubFolder, Processes, Process
    Dim sFile, sFolder, sPath, sConfigName, sContents, sProductCode, sLocalDrives
    Dim arrSubfolders
    Dim bKeepFolder, bDeleteSC
    
    Log "File CleanUp:"
    'On Error Resume Next
    If IsArray(arrDeleteFiles) Then
        If bForce Then
            Log "Doing Action: KillOSE"
            Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process")
            For Each Process in Processes
                Log "Running process : " & Process.Name
                If LCase(Left(Process.Name,3))="ose" Then 
                    Log "Terminating process: " & Process.Name
                    Process.Terminate
                End If
            Next 'Process
            Log "End Action: KillOSE"
            KillApps
        End If
        'Wipe individual files detected earlier
        For Each sFile in arrDeleteFiles
            If oFso.FileExists(sFile) Then DeleteFile sFile
        Next 'File
    End If 'IsArray
    'Wipe Catalyst in commonfiles
    sFolder = sCommonProgramFiles & "\microsoft shared\OFFICE12\Office Setup Controller\"
    If EnumFolderNames(sFolder,arrSubFolders) Then
        For Each SubFolder in arrSubFolders
            sPath = sFolder & SubFolder
            If InStr(SubFolder,".")>0 Then sConfigName = UCase(Left(SubFolder,InStr(SubFolder,".")-1))Else sConfigName = UCase(Subfolder)
            If GetFolderPath(sPath) Then
                Set Folder = oFso.GetFolder(sPath)
                Set Files = Folder.Files
                bKeepFolder = False
                For Each File In Files
                    If (LCase(Right(File.Name,4))=".xml") AND (UCase(Left(File.Name,Len(sConfigName)))=sConfigName) Then
                        Set XmlFile = oFso.OpenTextFile(File,1)
                        sContents = XmlFile.ReadAll
                        Set XmlFile = Nothing
                        sProductCode = Mid(sContents,InStr(sContents,"ProductCode=")+Len("ProductCode=")+1,38)
                        If CheckDelete(sProductCode) Then DeleteFile File.Path Else bKeepFolder = True
                    End If
                Next 'File
                Set Files = Nothing
                Set Folder = Nothing
                If Not bKeepFolder Then DeleteFolder sPath
            End If 'GetFolderPath
        Next 'SubFolder
    End If 'EnumFolderNames
    
    'Wipe Shortcuts
    'Find local hard disks
    On Error Resume Next
    Set LogicalDisks = oWmiLocal.ExecQuery("Select * from Win32_LogicalDisk")
    For Each Disk in LogicalDisks
        If Disk.DriveType = 3 Then
            sLocalDrives = sLocalDrives & UCase(Disk.DeviceID) & "\;"
        End If
    Next
    On Error Goto 0
    
    'Query all shortcuts 
    Log "Searching for shortcuts. This can take a few moments ..."
    Set scFiles = oWmiLocal.ExecQuery("Select * From Win32_ShortcutFile")
    For Each File in scFiles
        bDeleteSC = False
        'Ensure to keep the scope to local HD
        If InStr(sLocalDrives,UCase(Left(File.Description,2)))>0 Then
            'Compare if the shortcut target is in the list of executables that will be removed
            If Len(File.Target)>0 AND InStr(sApps,";" & File.Target & ";")>0 Then bDeleteSC = True
            'Handle Windows Installer shortcuts
            If InStr(File.Target,"{")>0 Then
                If Len(File.Target)>=InStr(File.Target,"{")+37 Then
                    If CheckDelete(Mid(File.Target,InStr(File.Target,"{"),38)) Then bDeleteSC = True
                End If
            End If
            If bDeleteSC Then 
                If Not IsArray(arrDeleteFolders) Then ReDim arrDeleteFolders(0)
                sFolder = Left(File.Description,InStrRev(File.Description,"\")-1)
                If Not arrDeleteFolders(UBound(arrDeleteFolders)) = sFolder Then
                    ReDim Preserve arrDeleteFolders(UBound(arrDeleteFolders)+1)
                    arrDeleteFolders(UBound(arrDeleteFolders)) = sFolder
                End If
                DeleteFile File.Description
            End If 'bDeleteSC
        End If 'InStr(sLocalDrives,UCase(Left(File.Description,2)))>0
    Next 'scFile
End Sub 'FileWipeIndividual
'=======================================================================================================

Sub DelScrubTmp
    Dim LogicalDisks, Disk
    Dim sFolder
    
    'Search all hard disks
    Set LogicalDisks = oWmiLocal.ExecQuery("Select * from Win32_LogicalDisk")
    For Each Disk in LogicalDisks
        If Disk.DriveType = 3 Then
            If oFso.FolderExists(Disk.DeviceID & "\ScrubTmp") Then 
                On Error Resume Next
                oFso.DeleteFolder Disk.DeviceID & "\ScrubTmp",True
                On Error Goto 0
            End If
        End If
    Next 'Disk
End Sub 'DelScrubTmp
'=======================================================================================================

'Ensure there are no unexpected .msi files in the scrub folder
Sub DeleteMsiScrubCache
    Dim Folder, File, Files
    
    Log "ScrubCache CleanUp:"
    Set Folder = oFso.GetFolder(sScrubDir) : CheckError "DeleteMsiScrubCache"
    Set Files = Folder.Files
    For Each File in Files
        CheckError "DeleteMsiScrubCache"
        If LCase(Right(File.Name,4))=".msi" Then
            CheckError "DeleteMsiScrubCache"
            DeleteFile File.Path : CheckError "DeleteMsiScrubCache"
        End If
    Next 'File
End Sub 'DeleteMsiScrubCache
'=======================================================================================================

Sub MsiClearOrphanedFiles
    Const USERSIDEVERYONE = "s-1-1-0"
    Const MSIINSTALLCONTEXT_ALL = 7
    Const MSIPATCHSTATE_ALL = 15

    On Error Resume Next

    Dim Patch, AllPatches, Product, AllProducts
    Dim File, Files, Folder
    Dim sFName, sLocalMsp, sLocalMsi, sPatchList, sMsiList

    Set Folder = oFso.GetFolder(sWinDir & "\Installer")
    Set Files = Folder.Files

    Log "Windows Installer cache CleanUp:"
    'Get a complete list of patches
    Err.Clear
    Set AllPatches = oMsi.PatchesEx("",USERSIDEVERYONE,MSIINSTALLCONTEXT_ALL,MSIPATCHSTATE_ALL)
    If Err <> 0 Then
        CheckError "MsiClearOrphanedFiles (msp)"
    Else
        'Fill a comma separated stringlist with all .msp patchfiles
        For Each Patch in AllPatches
            sLocalMsp = "" : sLocalMsp = LCase(Patch.Patchproperty("LocalPackage")) : CheckError "MsiClearOrphanedFiles (msp)"
            sPatchList = sPatchList & sLocalMsp & ","
        Next 'Patch

        'Delete all non referenced .msp files from %windir%\installer
        For Each File in Files
            sFName = "" : sFName = LCase(File.Path)
            If LCase(Right(sFName,4)) = ".msp" Then
                If Not InStr(sPatchList,sFName) > 0 Then
                    'While this is an orphaned file keep the scope of Office only
                    If InStr(UCase(MapTargets(File.Path)),OFFICEID)>0 Then DeleteFile File.Path
                End If
            End If 'LCase(Right(sFName,4))
        Next 'File
    End If 'Err=0

    'Get a complete list products
    Err.Clear
    Set AllProducts = oMsi.ProductsEx("",USERSIDEVERYONE,MSIINSTALLCONTEXT_ALL)
    If Err <> 0 Then
        CheckError "MsiClearOrphanedFiles (msi)"
    Else
        'Fill a comma separated stringlist with all .msi files
        For Each Product in AllProducts
            sLocalMsi = "" : sLocalMsi = LCase(Product.InstallProperty("LocalPackage")) : CheckError "MsiClearOrphanedFiles (msi)"
            sMsiList = sMsiList & sLocalMsi & ","
        Next 'Product

        'Delete all non referenced .msi files from %windir%\installer
        For Each File in Files
            sFName = "" : sFName = LCase(File.Path)
            If LCase(Right(sFName,4)) = ".msi" Then
                If Not InStr(sMsiList,sFName) > 0 Then
                    'While this is an orphaned file keep the scope of Office only
                    If UCase(Right(ProductCode(File.Path),PRODLEN))=OFFICEID Then DeleteFile File.Path
                End If
            End If 'LCase(Right(sFName,4)) = ".msi"
        Next 'File
    End If 'Err=0

End Sub 'MsiClearOrphanedFiles
'=======================================================================================================

Sub RegWipe
    Dim Item, Name, Sku
    Dim hDefKey, sSubKeyName, sCurKey, sValue, sGuid
    Dim bKeep, bSystemComponent0, bPackageIDs, bDisplayVersion
    Dim arrKeys, arrNames, arrTypes
    Dim iLoopCnt
    
    Log "Registry CleanUp:"
    'Wipe registry data
    If bRemoveAll Then
        RegDeleteKey HKCU,"Software\Microsoft\Office\12.0"
        RegDeleteKey HKCU,"Software\Policies\Microsoft\Office\12.0"
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Office\12.0"
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Terminal Server\Install\Software\Microsoft\Office\12.0"
        'Win32Assemblies
        hDefKey = HKCR
        sSubKeyName  = "Installer\Win32Assemblies\"
        If RegEnumKey(hDefKey,sSubKeyName,arrKeys) Then
            For Each Item in arrKeys
                If InStr(UCase(Item),"OFFICE12")>0 Then RegDeleteKey hDefKey,sSubKeyName & Item
            Next 'Item
        End If 'RegEnumKey
    End If 'bRemoveAll
    
    'Add/Remove Programs
    sSubKeyName = REG_ARP
    If RegEnumKey(HKLM,sSubKeyName,arrKeys) Then
        For Each Item in arrKeys
            '*0FF1CE*
            If Len(Item)>37 Then
                sGuid = UCase(Left(Item,38))
                If Right(sGuid,PRODLEN)=OFFICEID AND Mid(sGuid,4,2)="12" Then
                    If CheckDelete(sGuid) Then RegDeleteKey HKLM, sSubKeyName & Item
                End If 'Right(Item,PRODLEN)=OFFICEID
            End If 'Len(Item)>37
            
            'Config entries
            sCurKey = sSubKeyName & Item & "\"
            bSystemComponent0 = RegReadValue(HKLM,sCurKey,"SystemComponent",sValue,"REG_DWORD") AND sValue = "0"
            bPackageIDs = RegReadValue(HKLM,sCurKey,"PackageIds",sValue,"REG_MULTI_SZ")
            bDisplayVersion = RegReadValue(HKLM,sCurKey,"DisplayVersion",sValue,"REG_SZ") AND (Left(sValue,3) = "12.")
            If (bSystemComponent0 AND bPackageIDs AND bDisplayVersion) Then
                bKeep = False
                If Not bRemoveAll Then
                    For Each Sku in arrKeepSKUs
                        If UCase(Item) =  Sku Then bKeep = True
                    Next 'Sku
                End If
                If Not bKeep Then RegDeleteKey HKLM, sSubKeyName & Item
            End If
        Next 'Item
    End If 'RegEnumKey
    
    'UpgradeCodes, WI config, WI global config
    For iLoopCnt = 1 to 5
        Select Case iLoopCnt
        Case 1
            sSubKeyName = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UpgradeCodes\"
            hDefKey = HKLM
        Case 2 
            sSubKeyName = "Installer\UpgradeCodes\"
            hDefKey = HKCR
        Case 3
            sSubKeyName = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\"
            hDefKey = HKLM
        Case 4 
            sSubKeyName = "Installer\Features\"
            hDefKey = HKCR
        Case 5 
            sSubKeyName = "Installer\Products\"
            hDefKey = HKCR
        Case Else
            sSubKeyName = ""
            hDefKey = ""
        End Select
        If RegEnumKey(hDefKey,sSubKeyName,arrKeys) Then
            For Each Item in arrKeys
                'Ensure we have the expected length for a compressed GUID
                If Len(Item)=32 Then
                    'Expand the GUID
                    sGuid = GetExpandedGuid(Item) 
                    'Check if it's a Office 2007 key
                    If Right(sGuid,PRODLEN)=OFFICEID  AND Mid(sGuid,4,2)="12" Then
                        If bRemoveAll Then
                            RegDeleteKey hDefKey,sSubKeyName & Item
                        Else
                            If iLoopCnt < 3 Then
                                'Enum all entries
                                RegEnumValues hDefKey,sSubKeyName & Item,arrNames,arrTypes
                                If IsArray(arrNames) Then
                                    'Delete entries within removal scope
                                    For Each Name in arrNames
                                        If Len(Name)=32 Then
                                            sGuid = GetExpandedGuid(Name)
                                            If CheckDelete(sGuid) Then RegDeleteValue hDefKey, sSubKeyName & Item, Name
                                        Else
                                            'Invalid data -> delete the value
                                            RegDeleteValue hDefKey, sSubKeyName & Item, Name
                                        End If
                                    Next 'Name
                                End If 'IsArray(arrNames)
                                'If all entries were removed - delete the key
                                RegEnumValues hDefKey,sSubKeyName & Item,arrNames,arrTypes
                                If Not IsArray(arrNames) Then RegDeleteKey hDefKey, sSubKeyName & Item
                            Else 'iLoopCnt >= 3
                                If CheckDelete(sGuid) Then RegDeleteKey hDefKey, sSubKeyName & Item
                            End If 'iLoopCnt < 3
                        End If 'bRemoveAll
                    End If 'Right(Item,PRODLEN)=OFFICEID
                End If 'Len(Item)=32
            Next 'Item
        End If 'RegEnumKey
    Next 'iLoopCnt

    'Delivery
    hDefKey = HKLM
    sSubKeyName = "SOFTWARE\Microsoft\Office\Delivery\SourceEngine\Downloads\"
    If RegEnumKey(HKLM,sSubKeyName,arrKeys) Then
        For Each Item in arrKeys
            If bRemoveAll Then
                If (Mid(Item,26,PRODLEN)=OFFICEID AND Mid(Item,4,2)="12") OR _
                   LCase(Right(Item,7))="12.data" Then RegDeleteKey HKLM,sSubKeyName & Item
            Else
                If (Mid(Item,26,PRODLEN)=OFFICEID AND Mid(Item,4,2)="12") AND _
                   CheckDelete(UCase(Left(Item,38))) Then RegDeleteKey HKLM,sSubKeyName & Item
            End If
        Next 'Item
    End If 'RegEnumKey
    
    'Registration
    hDefKey = HKLM
    sSubKeyName = "SOFTWARE\Microsoft\Office\12.0\Registration\"
    If RegEnumKey(HKLM,sSubKeyName,arrKeys) Then
        For Each Item in arrKeys
            If CheckDelete(UCase(Left(Item,38))) Then RegDeleteKey HKLM,sSubKeyName & Item
        Next 'Item
    End If 'RegEnumKey
    
    'User Preconfigurations
    hDefKey = HKLM
    sSubKeyName = "SOFTWARE\Microsoft\Office\12.0\User Settings\"
    If RegEnumKey(HKLM,sSubKeyName,arrKeys) Then
        For Each Item in arrKeys
            If CheckDelete(UCase(Left(Item,38))) Then RegDeleteKey HKLM,sSubKeyName & Item
        Next 'Item
    End If 'RegEnumKey

    'Temporary entries in ARP
    TmpKeyCleanUp
End Sub 'RegWipeAll
'=======================================================================================================

'Clean up temporary registry keys
Sub TmpKeyCleanUp
    Dim TmpKey
    
    If bLogInitialized Then Log "Remove temporary registry entries:"
    If IsArray(arrTmpSKUs) Then
        For Each TmpKey in arrTmpSKUs
            'RegDeleteKey HKLM,REG_ARP & TmpKey
            oReg.DeleteKey HKLM, REG_ARP & TmpKey
        Next 'Item
    End If 'IsArray
End Sub 'TmpKeyCleanUp

'=======================================================================================================
' Helper Functions
'=======================================================================================================

'Kill all running instances of applications that will be removed
Sub KillApps
    Dim Processes, Process
    
    'On Error Resume Next
    Log "Doing Action: KillApps"
    Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process")
    For Each Process in Processes
        If InStr(LCase(sApps),"\" & LCase(Process.Name) & ";")>0 Then 
            Log "Killing process " & Process.Name
            Process.Terminate()
            CheckError "KillApps: " & "Process.Name"
        End If
    Next 'Process
    Log "End Action: KillApps"
End Sub 'KillApps
'=======================================================================================================

'Ensure Windows Explorer is restarted if needed
Sub RestoreExplorer
    Dim Processes
    
    On Error Resume Next
    wscript.sleep 1000
    Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process Where Name='explorer.exe'")
    If Processes.Count < 1 Then oWShell.Run "explorer.exe"
End Sub 'RestoreExploer
'=======================================================================================================

'Check registry access permissions. Failure will terminate the script
Function CheckRegPermissions
    Const KEY_QUERY_VALUE       = &H0001
    Const KEY_SET_VALUE         = &H0002
    Const KEY_CREATE_SUB_KEY    = &H0004
    Const DELETE                = &H00010000

    Dim sSubKeyName
    Dim bReturn

    CheckRegPermissions = True
    sSubKeyName = "Software\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\"
    oReg.CheckAccess HKLM, sSubKeyName, KEY_QUERY_VALUE, bReturn
    If Not bReturn Then CheckRegPermissions = False
    oReg.CheckAccess HKLM, sSubKeyName, KEY_SET_VALUE, bReturn
    If Not bReturn Then CheckRegPermissions = False
    oReg.CheckAccess HKLM, sSubKeyName, KEY_CREATE_SUB_KEY, bReturn
    If Not bReturn Then CheckRegPermissions = False
    oReg.CheckAccess HKLM, sSubKeyName, DELETE, bReturn
    If Not bReturn Then CheckRegPermissions = False

End Function 'CheckRegPermissions
'=======================================================================================================

'Check if an Office 12 product is still registered with a SKU that stays on the computer
Function CheckDelete(sProductCode)
    Dim Sku
    Dim RetVal
    Dim sProductCodeList
        
    'If it's a non Office 12 ProductCode exit with false right away
    CheckDelete = Right(sProductCode,PRODLEN) = OFFICEID
    If CheckDelete Then CheckDelete = Mid(sProductCode,4,2) = "12"
    If Not CheckDelete Then Exit Function
    If Not IsArray(arrKeepSKUs) Then Exit Function
    For Each Sku in arrKeepSKUs
        RetVal = RegReadValue(HKLM,REG_ARP & Sku,"ProductCodes",sProductCodeList,"REG_MULTI_SZ")
        If InStr(sProductCodeList,sProductCode) > 0 Then 
            CheckDelete = False
            Exit Function
        End If
    Next 'Sku
End Function 'CheckDelete
'=======================================================================================================

'Returns a string with a list of ProductCodes from the summary information stream
Function MapTargets (sMspFile)
    Const MSIOPENDATABASEMODE_PATCHFILE = 32
    Const PID_TEMPLATE                  =  7
    
    Dim Msp
    On Error Resume Next
    MapTargets = ""
    If oFso.FileExists(sMspFile) Then
        Set Msp = Msi.OpenDatabase(WScript.Arguments(0),MSIOPENDATABASEMODE_PATCHFILE)
        If Err = 0 Then MapTargets = Msp.SummaryInformation.Property(PID_TEMPLATE)
    End If 'oFso.FileExists(sMspFile)
End Function 'MspTargets
'=======================================================================================================

'Return the ProductCode {GUID} from a .MSI package
Function ProductCode(sMsi)
    Const MSIUILEVELNONE = 2 'No UI
    Dim MsiSession

    On Error Resume Next
    If oFso.FileExists(sMsi) Then
        oMsi.UILevel = MSIUILEVELNONE
        Set MsiSession = oMsi.OpenPackage(sMsi,1)
        ProductCode = MsiSession.ProductProperty("ProductCode")
        Set MsiSession = Nothing
    Else
        ProductCode = ""
    End If 'oFso.FileExists(sMsi)
End Function 'ProductCode
'=======================================================================================================

Function GetExpandedGuid (sGuid)
    Dim i

    GetExpandedGuid = "{" & StrReverse(Mid(sGuid,1,8)) & "-" & _
                       StrReverse(Mid(sGuid,9,4)) & "-" & _
                       StrReverse(Mid(sGuid,13,4))& "-"
    For i = 17 To 20
	    If i Mod 2 Then
		    GetExpandedGuid = GetExpandedGuid & mid(sGuid,(i + 1),1)
	    Else
		    GetExpandedGuid = GetExpandedGuid & mid(sGuid,(i - 1),1)
	    End If
    Next
    GetExpandedGuid = GetExpandedGuid & "-"
    For i = 21 To 32
	    If i Mod 2 Then
		    GetExpandedGuid = GetExpandedGuid & mid(sGuid,(i + 1),1)
	    Else
		    GetExpandedGuid = GetExpandedGuid & mid(sGuid,(i - 1),1)
	    End If
    Next
    GetExpandedGuid = GetExpandedGuid & "}"
End Function
'=======================================================================================================

'Converts a GUID into the compressed format
Function GetCompressedGuid (sGuid)
    Dim sCompGUID
    Dim i
    sCompGUID = StrReverse(Mid(sGuid,2,8))  & _
                StrReverse(Mid(sGuid,11,4)) & _
                StrReverse(Mid(sGuid,16,4)) 
    For i = 21 To 24
	    If i Mod 2 Then
		    sCompGUID = sCompGUID & Mid(sGuid, (i + 1), 1)
	    Else
		    sCompGUID = sCompGUID & Mid(sGuid, (i - 1), 1)
	    End If
    Next
    For i = 26 To 37
	    If i Mod 2 Then
		    sCompGUID = sCompGUID & Mid(sGuid, (i - 1), 1)
	    Else
		    sCompGUID = sCompGUID & Mid(sGuid, (i + 1), 1)
	    End If
    Next
    GetCompressedGuid = sCompGUID
End Function
'=======================================================================================================

'Create a backup copy of the file in the ScrubDir then delete the file
Sub CopyAndDeleteFile(sFile)
    Dim File
    On Error Resume Next
    If oFso.FileExists(sFile) Then
        Set File = oFso.GetFile(sFile)
        If Not oFso.FolderExists(sScrubDir & "\" & File.ParentFolder.Name) Then oFso.CreateFolder sScrubDir & "\" & File.ParentFolder.Name
        oFso.CopyFile sFile,sScrubDir & "\" & File.ParentFolder.Name & "\" & File.Name,True : CheckError "CopyAndDeleteFile"
        Set File = Nothing
        DeleteFile(sFile)
    End If 'oFso.FileExists
End Sub 'CopyAndDeleteFile
'=======================================================================================================

'Wrapper to delete a file
Sub DeleteFile(sFile)
    Dim File, Process, Processes
    Dim sFileName, sNewPath, sProcessList
    
    On Error Resume Next
    If oFso.FileExists(sFile) Then
        Log " - Delete file: " & sFile
        If Not bSimulate Then oFso.DeleteFile sFile,True ': CheckError "DeleteFile"
        If Err <> 0 Then
            CheckError "DeleteFile"
            'Try to move the file and delete from there
            Set File = oFso.GetFile(sFile)
            sFileName = File.Name
            sNewPath = File.Drive.Path & "\" & "ScrubTmp"
            Set File = Nothing
            'Ensure we stay within the same drive
            If Not oFso.FolderExists(sNewPath) Then oFso.CreateFolder(sNewPath)
            'Move the file
            Log " - Move file to: " & sNewPath & "\" & sFileName
            oFso.MoveFile sFile,sNewPath & "\" & sFileName
            If Err <> 0 Then
                CheckError "DeleteFile (move)"
            Else
                If Not InStr(sMoveMessage,sNewPath)>0 Then sMoveMessage = sMoveMessage & sNewPath & ";"
                oFso.DeleteFile sNewPath & "\" & sFileName,True 
                If Err <> 0 And bForce Then 
                    CheckError "DeleteFile (moved)"
                End If
            End If 'Err <> 0
        End If 'Err <> 0
    End If 'oFso.FileExists
End Sub 'DeleteFile
'=======================================================================================================

'64 bit aware wrapper to return the requested folder 
Function GetFolderPath(sPath)
    GetFolderPath = True
    If oFso.FolderExists(sPath) Then Exit Function
    If b64 AND oFso.FolderExists(Wow64Folder(sPath)) Then
        sPath = Wow64Folder(sPath)
        Exit Function
    End If
    GetFolderPath = False
End Function 'GetFolderPath
'=======================================================================================================

'Enumerates subfolder names of a folder and returns True if subfolders exist
Function EnumFolderNames (sFolder, arrSubFolders)
    Dim Folder, Subfolder
    Dim sSubFolders
    
    If oFso.FolderExists(sFolder) Then
        Set Folder = oFso.GetFolder(sFolder)
        For Each Subfolder in Folder.Subfolders
            sSubFolders = sSubFolders & Subfolder.Name & ","
        Next 'Subfolder
    End If
    If b64 AND oFso.FolderExists(Wow64Folder(sFolder)) Then
        Set Folder = oFso.GetFolder(Wow64Folder(sFolder))
        For Each Subfolder in Folder.Subfolders
            sSubFolders = sSubFolders & Subfolder.Name & ","
        Next 'Subfolder
    End If
    If Len(sSubFolders)>0 Then arrSubFolders = RemoveDuplicates(Split(Left(sSubFolders,Len(sSubFolders)-1),","))
    EnumFolderNames = Len(sSubFolders)>0
End Function 'EnumFolderNames
'=======================================================================================================

'Enumerates subfolders of a folder and returns True if subfolders exist
Function EnumFolders (sFolder, arrSubFolders)
    Dim Folder, Subfolder
    Dim sSubFolders
    
    If oFso.FolderExists(sFolder) Then
        Set Folder = oFso.GetFolder(sFolder)
        For Each Subfolder in Folder.Subfolders
            sSubFolders = sSubFolders & Subfolder.Path & ","
        Next 'Subfolder
    End If
    If b64 AND oFso.FolderExists(Wow64Folder(sFolder)) Then
        Set Folder = oFso.GetFolder(Wow64Folder(sFolder))
        For Each Subfolder in Folder.Subfolders
            sSubFolders = sSubFolders & Subfolder.Path & ","
        Next 'Subfolder
    End If
    If Len(sSubFolders)>0 Then arrSubFolders = RemoveDuplicates(Split(Left(sSubFolders,Len(sSubFolders)-1),","))
    EnumFolders = Len(sSubFolders)>0
End Function 'EnumFolders
'=======================================================================================================

Sub GetFolderStructure (Folder)
    Dim SubFolder
    
    For Each SubFolder in Folder.SubFolders
        ReDim Preserve arrMseFolders(UBound(arrMseFolders)+1)
        arrMseFolders(UBound(arrMseFolders)) = SubFolder.Path
        GetFolderStructure SubFolder
    Next 'SubFolder
End Sub 'GetFolderStructure
'=======================================================================================================

'Wrapper to delete a folder 
Sub DeleteFolder(sFolder)
    Dim Folder
    Dim sDelFolder, sFolderName, sNewPath
    On Error Resume Next
    If oFso.FolderExists(sFolder) Then 
        sDelFolder = sFolder
    ElseIf b64 AND oFso.FolderExists(Wow64Folder(sFolder)) Then 
        sDelFolder = Wow64Folder(sFolder)
    Else
        Exit Sub
    End If
    Log " - Delete folder: " & sDelFolder
    If Not bSimulate Then oFso.DeleteFolder sDelFolder,True
    If Err <> 0 Then
        CheckError "DeleteFolder"
        stop
        'Try to move the folder and delete from there
        Set Folder = oFso.GetFolder(sDelFolder)
        sFolderName = Folder.Name
        sNewPath = Folder.Drive.Path & "\" & "ScrubTmp"
        Set Folder = Nothing
        'Ensure we stay within the same drive
        If Not oFso.FolderExists(sNewPath) Then oFso.CreateFolder(sNewPath)
        'Move the folder
        Log " - Moving folder to: " & sNewPath & "\" & sFolderName
        oFso.MoveFolder sFolder,sNewPath & "\" & sFolderName
        If Err <> 0 Then
            CheckError "DeleteFolder (move)"
        Else
            oFso.DeleteFolder sNewPath & "\" & sFolderName,True 
            If Err <> 0 And bForce Then 
                CheckError "DeleteFolder (moved)"
            End If
        End If 'Err <> 0
    End If 'Err <> 0
End Sub 'DeleteFolder
'=======================================================================================================

'Delete empty folder structures
Sub DeleteEmptyFolders
    Dim Folder
    Dim sFolder
    
    If Not IsArray(arrDeleteFolders) Then Exit Sub
    Log "Empty Folder Cleanup"
    For Each sFolder in arrDeleteFolders
        If oFso.FolderExists(sFolder) Then
            Set Folder = oFso.GetFolder(sFolder)
            If (Folder.Subfolders.Count = 0) AND (Folder.Files.Count = 0) Then 
                Set Folder = Nothing
                SmartDeleteFolder sFolder
            End If
        End If
    Next 'sFolder
End Sub 'DeleteEmptyFolders
'=======================================================================================================

'Wrapper to delete a folder and remove the empty parent folder structure
Sub SmartDeleteFolder(sFolder)
    If oFso.FolderExists(sFolder) Then 
        Log "Request SmartDelete for folder: " & sFolder
        SmartDeleteFolderEx sFolder
    End If
    If b64 AND oFso.FolderExists(Wow64Folder(sFolder)) Then 
        Log "Request SmartDelete for folder: " & Wow64Folder(sFolder)
        SmartDeleteFolderEx Wow64Folder(sFolder)
    End If
End Sub 'SmartDeleteFolder
'=======================================================================================================

'Executes the folder delete operation
Sub SmartDeleteFolderEx(sFolder)
    Dim Folder
    
    On Error Resume Next
    DeleteFolder sFolder : CheckError "SmartDeleteFolderEx"
    On Error Goto 0
    Set Folder = oFso.GetFolder(oFso.GetParentFolderName(sFolder))
    If (Folder.Subfolders.Count = 0) AND (Folder.Files.Count = 0) Then SmartDeleteFolderEx(Folder.Path)
End Sub 'SmartDeleteFolderEx
'=======================================================================================================

'Handles additional folder-path operations on 64 bit environments
Function Wow64Folder(sFolder)
    If LCase(Left(sFolder,Len(sWinDir & "\System32"))) = LCase(sWinDir & "\System32") Then 
        Wow64Folder = sWinDir & "\syswow64" & Right(sFolder,Len(sFolder)-Len(sSys32Dir))
    ElseIf LCase(Left(sFolder,Len(sProgramFiles))) = LCase(sProgramFiles) Then 
        Wow64Folder = sProgramFilesX86 & Right(sFolder,Len(sFolder)-Len(sProgramFiles))
    Else
        Wow64Folder = "?" 'Return invalid string to ensure the folder cannot exist
    End If
End Function 'Wow64Folder
'=======================================================================================================

Function HiveString(hDefKey)
    On Error Resume Next
    Select Case hDefKey
        Case HKCR : HiveString = "HKEY_CLASSES_ROOT"
        Case HKCU : HiveString = "HKEY_CURRENT_USER"
        Case HKLM : HiveString = "HKEY_LOCAL_MACHINE"
        Case HKU  : HiveString = "HKEY_USERS"
        Case Else : HiveString = hDefKey
    End Select
End Function
'=======================================================================================================

Function RegKeyExists(hDefKey,sSubKeyName)
    Dim arrKeys
    RegKeyExists = False
    If oReg.EnumKey(hDefKey,sSubKeyName,arrKeys) = 0 Then RegKeyExists = True
End Function
'=======================================================================================================

Function RegValExists(hDefKey,sSubKeyName,sName)
    Dim arrValueTypes, arrValueNames
    Dim i

    RegValExists = False
    If Not RegKeyExists(hDefKey,sSubKeyName) Then Exit Function
    If oReg.EnumValues(hDefKey,sSubKeyName,arrValueNames,arrValueTypes) = 0 AND IsArray(arrValueNames) Then
        For i = 0 To UBound(arrValueNames) 
            If LCase(arrValueNames(i)) = Trim(LCase(sName)) Then RegValExists = True
        Next 
    End If 'oReg.EnumValues
End Function
'=======================================================================================================

'Read the value of a given registry entry
Function RegReadValue(hDefKey, sSubKeyName, sName, sValue, sType)
    Dim RetVal
    Dim Item
    Dim arrValues
    
    Select Case UCase(sType)
        Case "REG_SZ"
            RetVal = oReg.GetStringValue(hDefKey,sSubKeyName,sName,sValue)
            If Not RetVal = 0 AND b64 Then RetVal = oReg.GetStringValue(hDefKey,Wow64Key(hDefKey, sSubKeyName),sName,sValue)
        
        Case "REG_EXPAND_SZ"
            RetVal = oReg.GetExpandedStringValue(hDefKey,sSubKeyName,sName,sValue)
            If Not RetVal = 0 AND b64 Then RetVal = oReg.GetExpandedStringValue(hDefKey,Wow64Key(hDefKey, sSubKeyName),sName,sValue)
        
        Case "REG_MULTI_SZ"
            RetVal = oReg.GetMultiStringValue(hDefKey,sSubKeyName,sName,arrValues)
            If Not RetVal = 0 AND b64 Then RetVal = oReg.GetMultiStringValue(hDefKey,Wow64Key(hDefKey, sSubKeyName),sName,arrValues)
            If RetVal = 0 Then sValue = Join(arrValues,chr(34))
        
        Case "REG_DWORD"
            RetVal = oReg.GetDWORDValue(hDefKey,sSubKeyName,sName,sValue)
            If Not RetVal = 0 AND b64 Then 
                RetVal = oReg.GetDWORDValue(hDefKey,Wow64Key(hDefKey, sSubKeyName),sName,sValue)
            End If
        
        Case "REG_BINARY"
            RetVal = oReg.GetBinaryValue(hDefKey,sSubKeyName,sName,sValue)
            If Not RetVal = 0 AND b64 Then RetVal = oReg.GetBinaryValue(hDefKey,Wow64Key(hDefKey, sSubKeyName),sName,sValue)
        
        Case "REG_QWORD"
            RetVal = oReg.GetQWORDValue(hDefKey,sSubKeyName,sName,sValue)
            If Not RetVal = 0 AND b64 Then RetVal = oReg.GetQWORDValue(hDefKey,Wow64Key(hDefKey, sSubKeyName),sName,sValue)
        
        Case Else
            RetVal = -1
    End Select 'sValue
    
    RegReadValue = (RetVal = 0)
End Function 'RegReadValue
'=======================================================================================================

'Enumerate a registry key to return all values
Function RegEnumValues(hDefKey,sSubKeyName,arrNames, arrTypes)
    Dim RetVal, RetVal64
    Dim arrNames32, arrNames64, arrTypes32, arrTypes64
    
    If b64 Then
        RetVal = oReg.EnumValues(hDefKey,sSubKeyName,arrNames32,arrTypes32)
        RetVal64 = oReg.EnumValues(hDefKey,Wow64Key(hDefKey, sSubKeyName),arrNames64,arrTypes64)
        If (RetVal = 0) AND (Not RetVal64 = 0) AND IsArray(arrNames32) AND IsArray(arrTypes32) Then 
            arrNames = arrNames32
            arrTypes = arrTypes32
        End If
        If (Not RetVal = 0) AND (RetVal64 = 0) AND IsArray(arrNames64) AND IsArray(arrTypes64) Then 
            arrNames = arrNames64
            arrTypes = arrTypes64
        End If
        If (RetVal = 0) AND (RetVal64 = 0) AND IsArray(arrNames32) AND IsArray(arrNames64) AND IsArray(arrTypes32) AND IsArray(arrTypes64) Then 
            arrNames = RemoveDuplicates(Split((Join(arrNames32,"\") & "\" & Join(arrNames64,"\")),"\"))
            arrTypes = RemoveDuplicates(Split((Join(arrTypes32,"\") & "\" & Join(arrTypes64,"\")),"\"))
        End If
    Else
        RetVal = oReg.EnumValues(hDefKey,sSubKeyName,arrNames,arrTypes)
    End If 'b64
    RegEnumValues = ((RetVal = 0) OR (RetVal64 = 0)) AND IsArray(arrNames) AND IsArray(arrTypes)
End Function 'RegEnumValues
'=======================================================================================================

'Enumerate a registry key to return all subkeys
Function RegEnumKey(hDefKey,sSubKeyName,arrKeys)
    Dim RetVal, RetVal64
    Dim arrKeys32, arrKeys64
    
    If b64 Then
        RetVal = oReg.EnumKey(hDefKey,sSubKeyName,arrKeys32)
        RetVal64 = oReg.EnumKey(hDefKey,Wow64Key(hDefKey, sSubKeyName),arrKeys64)
        If (RetVal = 0) AND (Not RetVal64 = 0) AND IsArray(arrKeys32) Then arrKeys = arrKeys32
        If (Not RetVal = 0) AND (RetVal64 = 0) AND IsArray(arrKeys64) Then arrKeys = arrKeys64
        If (RetVal = 0) AND (RetVal64 = 0) Then 
            If IsArray(arrKeys32) AND IsArray (arrKeys64) Then 
                arrKeys = RemoveDuplicates(Split((Join(arrKeys32,"\") & "\" & Join(arrKeys64,"\")),"\"))
            ElseIf IsArray(arrKeys64) Then
                arrKeys = arrKeys64
            Else
                arrKeys = arrKeys32
            End If
        End If
    Else
        RetVal = oReg.EnumKey(hDefKey,sSubKeyName,arrKeys)
    End If 'b64
    RegEnumKey = ((RetVal = 0) OR (RetVal64 = 0)) AND IsArray(arrKeys)
End Function 'RegEnumKey
'=======================================================================================================

'Wrapper around oReg.DeleteValue to handle 64 bit
Sub RegDeleteValue(hDefKey, sSubKeyName, sName)
    Dim sWow64Key
    
    If RegValExists(hDefKey,sSubKeyName,sName) Then
        Log " - Delete registry value: " & HiveString(hDefKey) & "\" & sSubKeyName & " -> " & sName
        On Error Resume Next
        If Not bSimulate Then oReg.DeleteValue hDefKey, sSubKeyName, sName : CheckError "RegDeleteValue"
        On Error Goto 0
    End If 'RegValExists
    If b64 Then 
        sWow64Key = Wow64Key(hDefKey, sSubKeyName)
        If RegValExists(hDefKey,sWow64Key,sName) Then
            Log " - Delete registry value: " & HiveString(hDefKey) & "\" & sWow64Key & " -> " & sName
            On Error Resume Next
            If Not bSimulate Then oReg.DeleteValue hDefKey, sWow64Key, sName
            On Error Goto 0
        End If 'RegKeyExists
    End If
End Sub 'RegDeleteValue
'=======================================================================================================

'Wrappper around RegDeleteKeyEx to handle 64bit scenrios
Sub RegDeleteKey(hDefKey, sSubKeyName)
    Dim sWow64Key
    
    If RegKeyExists(hDefKey, sSubKeyName) Then
        Log " - Delete registry value: " & HiveString(hDefKey) & "\" & sSubKeyName
        On Error Resume Next
        RegDeleteKeyEx hDefKey, sSubKeyName
        On Error Goto 0
    End If 'RegKeyExists
    If b64 Then 
        sWow64Key = Wow64Key(hDefKey, sSubKeyName)
        If RegKeyExists(hDefKey,sWow64Key) Then
            Log " - Delete registry value: " & HiveString(hDefKey) & "\" & sWow64Key
            On Error Resume Next
            RegDeleteKeyEx hDefKey, sWow64Key
            On Error Goto 0
        End If 'RegKeyExists
    End If
End Sub 'RegDeleteKey
'=======================================================================================================

'Recursively delete a registry structure
Sub RegDeleteKeyEx(hDefKey, sSubKeyName) 
    Dim arrSubkeys
    Dim sSubkey

    On Error Resume Next
    oReg.EnumKey hDefKey, sSubKeyName, arrSubkeys 
    If IsArray(arrSubkeys) Then 
        For Each sSubkey In arrSubkeys 
            RegDeleteKeyEx hDefKey, sSubKeyName & "\" & sSubkey 
        Next 
    End If 
    If Not bSimulate Then oReg.DeleteKey hDefKey, sSubKeyName 
End Sub 'RegDeleteKeyEx
'=======================================================================================================

'Return the alternate regkey location on 64bit environment
Function Wow64Key(hDefKey, sSubKeyName)
    Dim iPos

    Select Case hDefKey
        Case HKCU
            If Left(sSubKeyName,17) = "Software\Classes\" Then
                Wow64Key = Left(sSubKeyName,17) & "Wow6432Node\" & Right(sSubKeyName,Len(sSubKeyName)-17)
            Else
                iPos = InStr(sSubKeyName,"\")
                Wow64Key = Left(sSubKeyName,iPos) & "Wow6432Node\" & Right(sSubKeyName,Len(sSubKeyName)-iPos)
            End If
        
        Case HKLM
            If Left(sSubKeyName,17) = "Software\Classes\" Then
                Wow64Key = Left(sSubKeyName,17) & "Wow6432Node\" & Right(sSubKeyName,Len(sSubKeyName)-17)
            Else
                iPos = InStr(sSubKeyName,"\")
                Wow64Key = Left(sSubKeyName,iPos) & "Wow6432Node\" & Right(sSubKeyName,Len(sSubKeyName)-iPos)
            End If
        
        Case Else
            Wow64Key = "Wow6432Node\" & sSubKeyName
        
    End Select 'hDefKey
End Function 'Wow64Key
'=======================================================================================================

'Remove duplicate entries from a one dimensional array
Function RemoveDuplicates(Array)
    Dim Item
    Dim oDic
    
    Set oDic = CreateObject("Scripting.Dictionary")
    For Each Item in Array
        If Not oDic.Exists(Item) Then oDic.Add Item,Item
    Next 'Item
    RemoveDuplicates = oDic.Keys
End Function 'RemoveDuplicates
'=======================================================================================================

'Delete a service
Function DeleteService(sService)
    Dim Services, Service, Processes, Process
    Dim sQuery, sStates
    
    On Error Resume Next
    
    sStates = "STARTED;RUNNING"    
    sQuery = "Select * From Win32_Service Where Name='" & sService & "'"
    Set Services = oWmiLocal.Execquery(sQuery)
    
    'Stop and delete the service
    For Each Service in Services
        If InStr(sStates,UCase(Service.State))>0 Then
            Log "Send stop command to service: " & sService
            Service.StopService
        End If
        'Ensure no more instances of the service are running
        Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process")
        For Each Process in Processes
            If LCase(Left(Process.Name,Len(sService)))=sService Then 
                Log "Terminating running process: " & Process.Name
                Process.Terminate
            End If
        Next 'Process
        Log " - Deleting Service -> " & sService
        If Not bSimulate Then Service.Delete
    Next 'Service
    Set Services = Nothing
    
End Function 'DeleteService
'=======================================================================================================

'Translation for setup.exe error codes
Function SetupExeRetVal(RetVal)
    Select Case RetVal
        Case 0 : SetupExeRetVal = "Success"
        Case 30001,1 : SetupExeRetVal = "AbstractMethod"
        Case 30002,2 : SetupExeRetVal = "ApiProhibited"
        Case 30003,3  : SetupExeRetVal = "AlreadyImpersonatingAUser"
        Case 30004,4 : SetupExeRetVal = "AlreadyInitialized"
        Case 30005,5 : SetupExeRetVal = "ArgumentNullException"
        Case 30006,6 : SetupExeRetVal = "AssertionFailed"
        Case 30007,7 : SetupExeRetVal = "CABFileAddFailed"
        Case 30008,8 : SetupExeRetVal = "CommandFailed"
        Case 30009,9 : SetupExeRetVal = "ConcatenationFailed"
        Case 30010,10 : SetupExeRetVal = "CopyFailed"
        Case 30011,11 : SetupExeRetVal = "CreateEventFailed"
        Case 30012,12 : SetupExeRetVal = "CustomizationPatchNotFound"
        Case 30013,13 : SetupExeRetVal = "CustomizationPatchNotApplicable"
        Case 30014,14 : SetupExeRetVal = "DuplicateDefinition"
        Case 30015,15 : SetupExeRetVal = "ErrorCodeOnly - Passthrough for Win32 error"
        Case 30016,16 : SetupExeRetVal = "ExceptionNotThrown"
        Case 30017,17 : SetupExeRetVal = "FailedToImpersonateUser"
        Case 30018,18 : SetupExeRetVal = "FailedToInitializeFlexDataSource"
        Case 30019,19 : SetupExeRetVal = "FailedToStartClassFactories"
        Case 30020,20 : SetupExeRetVal = "FileNotFound"
        Case 30021,21 : SetupExeRetVal = "FileNotOpen"
        Case 30022,22 : SetupExeRetVal = "FlexDialogAlreadyInitialized"
        Case 30023,23 : SetupExeRetVal = "HResultOnly - Passthrough for HRESULT errors"
        Case 30024,24 : SetupExeRetVal = "HWNDNotFound"
        Case 30025,25 : SetupExeRetVal = "IncompatibleCacheAction"
        Case 30026,26 : SetupExeRetVal = "IncompleteProductAddOns"
        Case 30027,27 : SetupExeRetVal = "InstalledProductStateCorrupt"
        Case 30028,28 : SetupExeRetVal = "InsufficientBuffer"
        Case 30029,29 : SetupExeRetVal = "InvalidArgument"
        Case 30030,30 : SetupExeRetVal = "InvalidCDKey"
        Case 30031,31 : SetupExeRetVal = "InvalidColumnType"
        Case 30032,31 : SetupExeRetVal = "InvalidConfigAddLanguage"
        Case 30033,33 : SetupExeRetVal = "InvalidData"
        Case 30034,34 : SetupExeRetVal = "InvalidDirectory"
        Case 30035,35 : SetupExeRetVal = "InvalidFormat"
        Case 30036,36 : SetupExeRetVal = "InvalidInitialization"
        Case 30037,37 : SetupExeRetVal = "InvalidMethod"
        Case 30038,38 : SetupExeRetVal = "InvalidOperation"
        Case 30039,39 : SetupExeRetVal = "InvalidParameter"
        Case 30040,40 : SetupExeRetVal = "InvalidProductFromARP"
        Case 30041,41 : SetupExeRetVal = "InvalidProductInConfigXml"
        Case 30042,42 : SetupExeRetVal = "InvalidReference"
        Case 30043,43 : SetupExeRetVal = "InvalidRegistryValueType"
        Case 30044,44 : SetupExeRetVal = "InvalidXMLProperty"
        Case 30045,45 : SetupExeRetVal = "InvalidMetadataFile"
        Case 30046,46 : SetupExeRetVal = "LogNotInitialized"
        Case 30047,47 : SetupExeRetVal = "LogAlreadyInitialized"
        Case 30048,48 : SetupExeRetVal = "MissingXMLNode"
        Case 30049,49 : SetupExeRetVal = "MsiTableNotFound"
        Case 30050,50 : SetupExeRetVal = "MsiAPICallFailure"
        Case 30051,51 : SetupExeRetVal = "NodeNotOfTypeElement"
        Case 30052,52 : SetupExeRetVal = "NoMoreGraceBoots"
        Case 30053,53 : SetupExeRetVal = "NoProductsFound"
        Case 30054,54 : SetupExeRetVal = "NoSupportedCulture"
        Case 30055,55 : SetupExeRetVal = "NotYetImplemented"
        Case 30056,56 : SetupExeRetVal = "NotAvailableCulture"
        Case 30057,57 : SetupExeRetVal = "NotCustomizationPatch"
        Case 30058,58 : SetupExeRetVal = "NullReference"
        Case 30059,59 : SetupExeRetVal = "OCTPatchForbidden"
        Case 30060,60 : SetupExeRetVal = "OCTWrongMSIDll"
        Case 30061,61 : SetupExeRetVal = "OutOfBoundsIndex"
        Case 30062,62 : SetupExeRetVal = "OutOfDiskSpace"
        Case 30063,63 : SetupExeRetVal = "OutOfMemory"
        Case 30064,64 : SetupExeRetVal = "OutOfRange"
        Case 30065,65 : SetupExeRetVal = "PatchApplicationFailure"
        Case 30066,66 : SetupExeRetVal = "PreReqCheckFailure"
        Case 30067,67 : SetupExeRetVal = "ProcessAlreadyStarted"
        Case 30068,68 : SetupExeRetVal = "ProcessNotStarted"
        Case 30069,69 : SetupExeRetVal = "ProcessNotFinished"
        Case 30070,70 : SetupExeRetVal = "ProductAlreadyDefined"
        Case 30071,71 : SetupExeRetVal = "ResourceAlreadyTracked"
        Case 30072,72 : SetupExeRetVal = "ResourceNotFound"
        Case 30073,73 : SetupExeRetVal = "ResourceNotTracked"
        Case 30074,74 : SetupExeRetVal = "SQLAlreadyConnected"
        Case 30075,75 : SetupExeRetVal = "SQLFailedToAllocateHandle"
        Case 30076,76 : SetupExeRetVal = "SQLFailedToConnect"
        Case 30077,77 : SetupExeRetVal = "SQLFailedToExecuteStatement"
        Case 30078,78 : SetupExeRetVal = "SQLFailedToRetrieveData"
        Case 30079,79 : SetupExeRetVal = "SQLFailedToSetAttribute"
        Case 30080,80 : SetupExeRetVal = "StorageNotCreated"
        Case 30081,81 : SetupExeRetVal = "StreamNameTooLong"
        Case 30082,82 : SetupExeRetVal = "SystemError"
        Case 30083,83 : SetupExeRetVal = "ThreadAlreadyStarted"
        Case 30084,84 : SetupExeRetVal = "ThreadNotStarted"
        Case 30085,85 : SetupExeRetVal = "ThreadNotFinished"
        Case 30086,86 : SetupExeRetVal = "TooManyProducts"
        Case 30087,87 : SetupExeRetVal = "UnexpectedXMLNodeType"
        Case 30088,88 : SetupExeRetVal = "UnexpectedError"
        Case 30089,89 : SetupExeRetVal = "Unitialized"
        Case 30090,90 : SetupExeRetVal = "UserCancel"
        Case 30091,91 : SetupExeRetVal = "ExternalCommandFailed"
        Case 30092,92 : SetupExeRetVal = "SPDatabaseOverSize"
        Case 30093,93 : SetupExeRetVal = "IntegerTruncation"
        'msiexec return values
        Case 1259 : SetupExeRetVal = "APPHELP_BLOCK"
        Case 1601 : SetupExeRetVal = "INSTALL_SERVICE_FAILURE"
        Case 1602 : SetupExeRetVal = "INSTALL_USEREXIT"
        Case 1603 : SetupExeRetVal = "INSTALL_FAILURE"
        Case 1604 : SetupExeRetVal = "INSTALL_SUSPEND"
        Case 1605 : SetupExeRetVal = "UNKNOWN_PRODUCT"
        Case 1606 : SetupExeRetVal = "UNKNOWN_FEATURE"
        Case 1607 : SetupExeRetVal = "UNKNOWN_COMPONENT"
        Case 1608 : SetupExeRetVal = "UNKNOWN_PROPERTY"
        Case 1609 : SetupExeRetVal = "INVALID_HANDLE_STATE"
        Case 1610 : SetupExeRetVal = "BAD_CONFIGURATION"
        Case 1611 : SetupExeRetVal = "INDEX_ABSENT"
        Case 1612 : SetupExeRetVal = "INSTALL_SOURCE_ABSENT"
        Case 1613 : SetupExeRetVal = "INSTALL_PACKAGE_VERSION"
        Case 1614 : SetupExeRetVal = "PRODUCT_UNINSTALLED"
        Case 1615 : SetupExeRetVal = "BAD_QUERY_SYNTAX"
        Case 1616 : SetupExeRetVal = "INVALID_FIELD"
        Case 1618 : SetupExeRetVal = "INSTALL_ALREADY_RUNNING"
        Case 1619 : SetupExeRetVal = "INSTALL_PACKAGE_OPEN_FAILED"
        Case 1620 : SetupExeRetVal = "INSTALL_PACKAGE_INVALID"
        Case 1621 : SetupExeRetVal = "INSTALL_UI_FAILURE"
        Case 1622 : SetupExeRetVal = "INSTALL_LOG_FAILURE"
        Case 1623 : SetupExeRetVal = "INSTALL_LANGUAGE_UNSUPPORTED"
        Case 1624 : SetupExeRetVal = "INSTALL_TRANSFORM_FAILURE"
        Case 1625 : SetupExeRetVal = "INSTALL_PACKAGE_REJECTED"
        Case 1626 : SetupExeRetVal = "FUNCTION_NOT_CALLED"
        Case 1627 : SetupExeRetVal = "FUNCTION_FAILED"
        Case 1628 : SetupExeRetVal = "INVALID_TABLE"
        Case 1629 : SetupExeRetVal = "DATATYPE_MISMATCH"
        Case 1630 : SetupExeRetVal = "UNSUPPORTED_TYPE"
        Case 1631 : SetupExeRetVal = "CREATE_FAILED"
        Case 1632 : SetupExeRetVal = "INSTALL_TEMP_UNWRITABLE"
        Case 1633 : SetupExeRetVal = "INSTALL_PLATFORM_UNSUPPORTED"
        Case 1634 : SetupExeRetVal = "INSTALL_NOTUSED"
        Case 1635 : SetupExeRetVal = "PATCH_PACKAGE_OPEN_FAILED"
        Case 1636 : SetupExeRetVal = "PATCH_PACKAGE_INVALID"
        Case 1637 : SetupExeRetVal = "PATCH_PACKAGE_UNSUPPORTED"
        Case 1638 : SetupExeRetVal = "PRODUCT_VERSION"
        Case 1639 : SetupExeRetVal = "INVALID_COMMAND_LINE"
        Case 1640 : SetupExeRetVal = "INSTALL_REMOTE_DISALLOWED"
        Case 1641 : SetupExeRetVal = "SUCCESS_REBOOT_INITIATED"
        Case 1642 : SetupExeRetVal = "PATCH_TARGET_NOT_FOUND"
        Case 1643 : SetupExeRetVal = "PATCH_PACKAGE_REJECTED"
        Case 1644 : SetupExeRetVal = "INSTALL_TRANSFORM_REJECTED"
        Case 1645 : SetupExeRetVal = "INSTALL_REMOTE_PROHIBITED"
        Case 1646 : SetupExeRetVal = "PATCH_REMOVAL_UNSUPPORTED"
        Case 1647 : SetupExeRetVal = "UNKNOWN_PATCH"
        Case 1648 : SetupExeRetVal = "PATCH_NO_SEQUENCE"
        Case 1649 : SetupExeRetVal = "PATCH_REMOVAL_DISALLOWED"
        Case 1650 : SetupExeRetVal = "INVALID_PATCH_XML"
        Case 3010 : SetupExeRetVal = "SUCCESS_REBOOT_REQUIRED"
        Case Else : SetupExeRetVal = "Unknown Return Value"
    End Select
End Function 'SetupExeRetVal
'=======================================================================================================

Function GetProductID(sProdID)
        Dim sReturn
        
        Select Case sProdId
        
        Case "0010" : sReturn = "WEBFLDRS"
        Case "0011" : sReturn = "PROPLUS"
        Case "0012" : sReturn = "STANDARD"
        Case "0013" : sReturn = "BASIC"
        Case "0014" : sReturn = "PRO"
        Case "0015" : sReturn = "ACCESS"
        Case "0016" : sReturn = "EXCEL"
        Case "0017" : sReturn = "SharePointDesigner"
        Case "0018" : sReturn = "PowerPoint"
        Case "0019" : sReturn = "Publisher"
        Case "001A" : sReturn = "Outlook"
        Case "001B" : sReturn = "Word"
        Case "001C" : sReturn = "AccessRuntime"
        Case "001F" : sReturn = "Proof"
        Case "0020" : sReturn = "O2007CNV"
        Case "0021" : sReturn = "VisualWebDeveloper"
        Case "0026" : sReturn = "ExpressionWeb"
        Case "0029" : sReturn = "Excel"
        Case "002A" : sReturn = "Office64"
        Case "002B" : sReturn = "Word"
        Case "002C" : sReturn = "Proofing"
        Case "002E" : sReturn = "Ultimate"
        Case "002F" : sReturn = "HomeAndStudent"
        Case "0028" : sReturn = "IME"
        Case "0030" : sReturn = "Enterprise"
        Case "0031" : sReturn = "ProfessionalHybrid"
        Case "0033" : sReturn = "Personal"
        Case "0035" : sReturn = "ProfessionalHybrid"
        Case "0037" : sReturn = "PowerPoint"
        Case "003A" : sReturn = "PrjStd"
        Case "003B" : sReturn = "PrjPro"
        Case "0044" : sReturn = "InfoPath"
        Case "0045" : sReturn = "XWEB"
        Case "004A" : sReturn = "OWC11"
        Case "0051" : sReturn = "VISPRO"
        Case "0052" : sReturn = "VisView"
        Case "0053" : sReturn = "VisStd"
        Case "0054" : sReturn = "VisMUI"
        Case "0055" : sReturn = "VisMUI"
        Case "006E" : sReturn = "Shared"
        Case "008A" : sReturn = "RecentDocs"
        Case "00A1" : sReturn = "ONENOTE"
        Case "00A3" : sReturn = "OneNoteHomeStudent"
        Case "00A7" : sReturn = "CPAO"
        Case "00A9" : sReturn = "InterConnect"
        Case "00AF" : sReturn = "PPtView"
        Case "00B0" : sReturn = "ExPdf"
        Case "00B1" : sReturn = "ExXps"
        Case "00B2" : sReturn = "ExPdfXps"
        Case "00B4" : sReturn = "PrjMUI"
        Case "00B5" : sReturn = "PrjtMUI"
        Case "00B9" : sReturn = "AER"
        Case "00BA" : sReturn = "Groove"
        Case "00CA" : sReturn = "SmallBusiness"
        Case "00E0" : sReturn = "Outlook"
        Case "00D1" : sReturn = "ACE"
        Case "0100" : sReturn = "OfficeMUI"
        Case "0101" : sReturn = "OfficeXMUI"
        Case "0103" : sReturn = "PTK"
        Case "0114" : sReturn = "GrooveSetupMetadata"
        Case "0115" : sReturn = "SharedSetupMetadata"
        Case "0116" : sReturn = "SharedSetupMetadata"
        Case "0117" : sReturn = "AccessSetupMetadata"
        Case "011A" : sReturn = "LWConnect"
        Case "011F" : sReturn = "OLConnect"
        Case "1014" : sReturn = "STS"
        Case "1015" : sReturn = "WSSMUI"
        Case "1032" : sReturn = "PJSVRAPP"
        Case "104B" : sReturn = "SPS"
        Case "104E" : sReturn = "SPSMUI"
        Case "107F" : sReturn = "OSrv"
        Case "1080" : sReturn = "OSrv"
        Case "1088" : sReturn = "lpsrvwfe"
        Case "10D7" : sReturn = "IFS"
        Case "10D8" : sReturn = "IFSMUI"
        Case "10EB" : sReturn = "DLCAPP"
        Case "10F5" : sReturn = "XLSRVAPP"
        Case "10F6" : sReturn = "XlSrvWFE"
        Case "10F7" : sReturn = "DLC"
        Case "10F8" : sReturn = "SlSrvMui"
        Case "10FB" : sReturn = "OSrchWFE"
        Case "10FC" : sReturn = "OSRCHAPP"
        Case "10FD" : sReturn = "OSrchMUI"
        Case "1103" : sReturn = "DLC"
        Case "1104" : sReturn = "LHPSRV"
        Case "1105" : sReturn = "PIA"
        Case "110D" : sReturn = "OSERVER"
        Case "110F" : sReturn = "PSERVER"
        Case "1110" : sReturn = "WSS"
        Case "1121" : sReturn = "SPSSDK"
        Case "1122" : sReturn = "SPSDev"
        Case Else : sReturn = ""
        
        End Select 'sProdId
    GetProductID = sReturn
End Function 'GetProductID
'=======================================================================================================

Sub Log (sLog)
    wscript.echo sLog
    LogStream.WriteLine sLog
End Sub 'Log
'=======================================================================================================

Sub CheckError(sModule)
    If Err <> 0 Then 
        Log Now & " - " & sModule & " - Source: " & Err.Source & "; Err# (Hex): " & Hex( Err ) & _
               "; Err# (Dec): " & Err & "; Description : " & Err.Description
    End If 'Err = 0
    Err.Clear
End Sub
'=======================================================================================================

'Command line parser
Sub ParseCmdLine

    Dim iCnt, iArgCnt
    Dim arrArguments
    Dim sArg0
    
    iArgCnt = Wscript.Arguments.Count
    If iArgCnt = 0 Then
        'Create the log
        CreateLog
        Log "No argument specified. Preparing user prompt"
        FindInstalledO12Products
        If IsArray(arrInstalledSKUs) Then sDefault = Join(arrInstalledSKUs,",") Else sDefault = "ALL"
        sDefault = InputBox("Enter a list of Office 2007 products to remove" & vbCrLf & vbCrLf & _
                "Examples:" & vbCrLf & _
                "ALL" & vbTab & vbTab & "-> remove all of Office 2007" & vbCrLf & _
                "ProPlus,Project" & vbTab & "-> remove ProPlus and Project" & vbCrLf &_
                "?" & vbTab & vbTab & "-> display Help", _
                "Products to remove", _
                sDefault)
        If IsEmpty(sDefault) Then 'User cancelled
            Log "User cancelled. CleanUp & Exit."
            'Undo temporary entries created in ARP
            TmpKeyCleanUp
            wscript.quit 
        End If 'IsEmpty(sDefault)
        Log "Answer from prompt: " & sDefault
        sDefault = Trim(UCase(Trim(Replace(sDefault,Chr(34),""))))
        arrArguments = Split(Trim(sDefault)," ")
        If UBound(arrArguments) = -1 Then ReDim arrArguments(0)
    Else
        ReDim arrArguments(iArgCnt-1)
        For iCnt = 0 To (iArgCnt-1)
            arrArguments(iCnt) = UCase(Wscript.Arguments(iCnt))
        Next 'iCnt
    End If

    'Handle the SKU list
    sArg0 = Replace(arrArguments(0),"/","")
    sArg0 = Replace(sArg0,"-","")
    Select Case sArg0
    
    Case "?"
        ShowSyntax
    
    Case "ALL"
        bRemoveAll = True
        bRemoveOSE = False
    
    Case "ALL,OSE"
        bRemoveAll = True
        bRemoveOSE = True
    
    Case Else
        bRemoveAll = False
        bRemoveOSE = False
        sSkuRemoveList = arrArguments(0)
    
    End Select
    
    For iCnt = 0 To UBound(arrArguments)

        Select Case arrArguments(iCnt)
        
        Case "?","/?","-?"
            ShowSyntax
        
        Case "/B","/BYPASS"
            If UBound(arrArguments)>iCnt Then
                If InStr(arrArguments(iCnt+1),"1")>0 Then bBypass_Stage1 = True
                If InStr(arrArguments(iCnt+1),"2")>0 Then bBypass_Stage2 = True
                If InStr(arrArguments(iCnt+1),"3")>0 Then bBypass_Stage3 = True
                If InStr(arrArguments(iCnt+1),"4")>0 Then bBypass_Stage4 = True
            End If
        
        Case "/F","/FORCE"
            bForce = True
        
        Case "/L","/LOG"
            bLogInitialized = False
            If UBound(arrArguments)>iCnt Then
                If oFso.FolderExists(arrArguments(iCnt+1)) Then 
                    sLogDir = arrArguments(iCnt+1)
                Else
                    On Error Resume Next
                    oFso.CreateFolder(arrArguments(iCnt+1))
                    If Err <> 0 Then sLogDir = sScrubDir Else sLogDir = arrArguments(iCnt+1)
                End If
            End If
        
        Case "/N","/NOCANCEL"
            bNoCancel = True
        
        Case "/O","/OSE"
            bRemoveOSE = True
        
        Case "/Q","/QUIET"
            bQuiet = True
        
        Case "/P","/PREVIEW"
            bSimulate = True
        
        Case Else
        
        End Select
    Next 'iCnt
    If Not bLogInitialized Then CreateLog

End Sub 'ParseCmdLine
'=======================================================================================================

Sub CreateLog
    Dim DateTime
    
    'Create the log file
    Set DateTime = CreateObject("WbemScripting.SWbemDateTime")
    DateTime.SetVarDate Now,True
    On Error Resume Next
    Set LogStream = oFso.CreateTextFile(sLogDir & "\" & oWShell.ExpandEnvironmentStrings("%COMPUTERNAME%") & _
        "_" & Left(DateTime.Value,14) & "_ScrubLog.txt",True,True)
    If Err <> 0 Then 
        Err.Clear
        sLogDir = sScrubDir
        Set LogStream = oFso.CreateTextFile(sLogDir & "\" & oWShell.ExpandEnvironmentStrings("%COMPUTERNAME%") & _
            "_" & Left(DateTime.Value,14) & "_ScrubLog.txt",True,True)
    End If

    Log "Microsoft Customer Support Services - Office 2007 Removal Utility" & vbCrLf & vbCrLf & _
                "Version: " & VERSION & vbCrLf & _
                "64 bit OS: " & b64 & vbCrLf & _
                "Start removal: " & Now & vbCrLf
    bLogInitialized = True
End Sub 'CreateLog
'=======================================================================================================

Sub RelaunchAsCScript
    Dim Argument
    Dim sCmdLine
    
    sCmdLine = "cmd.exe /k " & WScript.Path & "\cscript.exe //NOLOGO " & Chr(34) & WScript.scriptFullName & Chr(34)
    If Wscript.Arguments.Count > 0 Then
        For Each Argument in Wscript.Arguments
            sCmdLine = sCmdLine  &  " " & chr(34) & Argument & chr(34)
        Next 'Argument
    End If
    oWShell.Run sCmdLine,1,False
    Wscript.Quit
End Sub 'RelaunchAsCScript
'=======================================================================================================

'Show the expected syntax for the script usage
Sub ShowSyntax
    TmpKeyCleanUp
    Wscript.Echo sErr & vbCrLf & _
             "OffScrub07 V " & VERSION & vbCrLf & _
             "Copyright (c) Microsoft Corporation. All Rights Reserved" & vbCrLf & vbCrLf & _
             "OffScrub07 helps to remove Office 2007 when a regular uninstall is no longer possible" & vbCrLf & vbCrLf & _
             "Usage:" & vbTab & "OffScrub07.vbs [List of config ProductIDs] [Options]" & vbCrLf & vbCrLf & _
             vbTab & "OffScrub07.vbs ALL               ' Remove all Office 2007 products" & vbCrLf &_
             vbTab & "OffScrub07.vbs ProPlus,Project   ' Remove ProPlus and Project" & vbCrLf &_
             vbTab & "OffScrub07.vbs ALL,OSE           ' Remove all products & OSE Service" & vbCrLf &_
             vbTab & "/Bypass [List of stage#]         ' List of stages that should not run" & vbCrLf & vbCrLf &_
             vbTab & vbTab & "1 = Component Detection" & vbCrLf &_
             vbTab & vbTab & "2 = Setup.exe" & vbCrLf &_
             vbTab & vbTab & "3 = Msiexec.exe" & vbCrLf &_
             vbTab & vbTab & "4 = CleanUp of additonal files and registry settings" & vbCrLf & vbCrLf &_
             vbTab & "/?                               ' Displays this help"& vbCrLf &_
             vbTab & "/Force                           ' Forces termination of running processes. May cause data loss!" & vbCrLf &_
             vbTab & "/Log [LogfolderPath]             ' Custom folder for log files" & vbCrLf & _
             vbTab & "/NoCancel                        ' Setup.exe and Msiexec.exe have no Cancel button" & vbCrLf &_
             vbTab & "/OSE                             ' Forces removal of the Office Source Engine service" & vbCrLf &_
             vbTab & "/Quiet                           ' Setup.exe and Msiexec.exe run quiet with no UI" & vbCrLf &_
             vbTab & "/Preview                         ' Run this script to preview what would get removed"
    Wscript.Quit
End Sub 'ShowSyntax
'=======================================================================================================