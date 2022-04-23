'=======================================================================================================
' Name: OffScrub03.vbs
' Author: Microsoft Customer Support Services
' Copyright (c) 2010, Microsoft Corporation
' Script to remove (scrub) Office 2003 products
'=======================================================================================================
Option Explicit
On Error Resume Next

Const VERSION       = "1.07"
Const HKCR          = &H80000000
Const HKCU          = &H80000001
Const HKLM          = &H80000002
Const HKU           = &H80000003
Const FOR_WRITING   = 2
Const PRODLEN       = 28
Const OFFICE_2003   = "6000-11D3-8CFE-0150048383C9}"
Const COMPPERMANENT = "00000000000000000000000000000000"
Const UNCOMPRESSED  = 38
Const COMPRESSED    = 32
Const REG_ARP       = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
Const DELIVERY      = "SOFTWARE\Microsoft\Office\11.0\Delivery\"

'=======================================================================================================
Dim oFso, oMsi, oReg, oWShell, oWmiLocal
Dim ComputerItem, Item, LogStream, TmpKey
Dim arrInstalledSKUs, arrRemoveSKUs, arrKeepSKUs, arrTmpSKUs
Dim arrDeleteFiles, arrDeleteFolders, arrMseFolders
Dim dicKeepProd, dicRemoveProd, dicInstalledProd, dicKeepLis, dicApps, dicKeepFolder
Dim f64
Dim sErr, sTmp, sSkuInstalledList, sSkuRemoveList, sDefault, sWinDir, sMode
Dim sAppData, sTemp, sScrubDir, sProgramFiles, sProgramFilesX86, sCommonProgramFiles, sAllusersProfile
Dim sOInstallRoot

'=======================================================================================================
'Main
'=======================================================================================================
'Configure defaults
Dim sLogDir : sLogDir = ""
Dim sMoveMessage: sMoveMessage = ""
Dim fRemoveOSE      : fRemoveOSE = False
Dim fRemoveAll      : fRemoveAll = False
Dim fKeepUser       : fKeepUser = True  'Default to keep per user settings
Dim fForceKeepUser  : fForceKeepUser = False 'Extended user settings behavior
Dim fSkipSD         : fSkipSD = False 'Default to not Skip the Shortcut Detection
Dim fDetectOnly     : fDetectOnly = False
Dim bQuiet          : bQuiet = False
'CAUTION! -> "fForce" will kill running applications which can result in data loss! <- CAUTION
Dim fForce          : fForce = False
'CAUTION! -> "fForce" will kill running applications which can result in data loss! <- CAUTION
Dim fLogInitialized : fLogInitialized = False
Dim fBypass_Stage1  : fBypass_Stage1 = False 'Component Detection
Dim fBypass_Stage2  : fBypass_Stage2 = False 'Msiexec
Dim fBypass_Stage3  : fBypass_Stage3 = False 'CleanUp
Dim fRebootRequired : fRebootRequired = False

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
sScrubDir           = sTemp & "\OffScrub03"

'Create Dictionaries
Set dicKeepProd = CreateObject("Scripting.Dictionary")
Set dicRemoveProd = CreateObject("Scripting.Dictionary")
Set dicInstalledProd = CreateObject("Scripting.Dictionary")
Set dicKeepLis = CreateObject("Scripting.Dictionary")
Set dicKeepFolder = CreateObject("Scripting.Dictionary")
Set dicApps = CreateObject("Scripting.Dictionary")
dicApps.Add "communicator.exe","communicator.exe"

'Create the temp folder
If Not oFso.FolderExists(sScrubDir) Then oFso.CreateFolder sScrubDir

'Set the default logging directory
sLogDir = sScrubDir

'Detect if we're running on a 64 bit OS
Set ComputerItem = oWmiLocal.ExecQuery("Select * from Win32_ComputerSystem")
For Each Item In ComputerItem
    f64 = Instr(Left(Item.SystemType,3),"64") > 0
Next
If f64 Then sProgramFilesX86 = oWShell.ExpandEnvironmentStrings("%programfiles(x86)%")

'Call the command line parser
ParseCmdLine

If Not CheckRegPermissions Then
    Log vbCrLf & "Insufficient registry access permissions - exiting"
    wscript.quit 
End If


'-------------------
'Stage # 0 - Basics |
'-------------------
'Get Office Install Folder
If NOT RegReadValue(HKLM,"SOFTWARE\Microsoft\Office\11.0\Common\InstallRoot","Path",sOInstallRoot,"REG_SZ") Then 
    sOInstallRoot = sProgramFiles & "\Microsoft Office\Office11"
End If

'Build a list with installed/registered Office 2003 products
sTmp = "Stage # 0 " & chr(34) & "Basics" & chr(34) & " (" & Time & ")"
Log vbCrLf & sTmp & vbCrLf & String(Len(sTmp),"=") & vbCrLf

FindInstalledO11Products
If dicInstalledProd.Count > 0 Then Log "Found registered product(s): " & Join(RemoveDuplicates(dicInstalledProd.Items),",") &vbCrLf

'Validate the list of products we got from the command line if applicable
ValidateRemoveSkuList

'Check which parts of the local installation source has to remain
CheckLIS

'Log detection results
If dicRemoveProd.Count > 0 Then Log "Product(s) to be removed: " & Join(RemoveDuplicates(dicRemoveProd.Items),",")
sMode = "Selected Office 2003 products"
If NOT dicRemoveProd.Count > 0 Then sMode = "Orphaned Office 2003 products"
If fRemoveAll Then sMode = "All Office 2003 products"
Log "Final removal mode: " & sMode & vbCrLf
Log "Remove OSE service: " & fRemoveOSE &vbCrLf

'Log preview mode if applicable
If fDetectOnly Then Log "*************************************************************************"
If fDetectOnly Then Log "*                          PREVIEW MODE                                 *"
If fDetectOnly Then Log "* All uninstall and delete operations will only be logged not executed! *"
If fDetectOnly Then Log "*************************************************************************"

'Cache .msi files
If dicRemoveProd.Count > 0 Then CacheMsiFiles

'--------------------------------
'Stage # 1 - Component Detection |
'--------------------------------
sTmp = "Stage # 1 " & chr(34) & "Component Detection" & chr(34) & " (" & Time & ")"
Log vbCrLf & sTmp & vbCrLf & String(Len(sTmp),"=") & vbCrLf
If Not fBypass_Stage1 Then
    'Build a list with files which are installed/registered to a product that's going to be removed
    Log "Prepare for CleanUp stages."
    Log "Identifying removable files. This can take several minutes."
    BuildFileList 
Else
    Log "Skipping Component Detection because bypass was requested."
End If

'Kill all running Office applications
If fForce OR bQuiet Then CloseOfficeApps

'------------------------
'Stage # 2 - Msiexec.exe |
'------------------------
sTmp = "Stage # 2 " & chr(34) & "Msiexec.exe" & chr(34) & " (" & Time & ")"
Log vbCrLf & sTmp & vbCrLf & String(Len(sTmp),"=") & vbCrLf
If Not fBypass_Stage2 Then
    MsiexecRemoval
Else
    Log "Skipping Msiexec.exe because bypass was requested."
End If

'--------------------
'Stage # 3 - CleanUp |
'--------------------
'Removal of files and registry settings
sTmp = "Stage # 3 " & chr(34) & "CleanUp" & chr(34) & " (" & Time & ")"
Log vbCrLf & sTmp & vbCrLf & String(Len(sTmp),"=") & vbCrLf
If Not fBypass_Stage3 Then

    'Office Source Engine
    If fRemoveOSE Then RemoveOSE

    'Local Installation Source (MSOCache)
    WipeLIS
    
    'Obsolete files
    If fRemoveAll Then 
        FileWipeAll 
    Else 
        FileWipeIndividual
    End If
    
    'Empty Folders
    DeleteEmptyFolders
    
    'Restore Explorer if needed
    If fForce Then RestoreExplorer
    
    'Registry data
    RegWipe
    
    'Wipe orphaned files from Windows Installer cache
    MsiClearOrphanedFiles
    
    'Temporary .msi files in scrubcache
    DeleteMsiScrubCache
    
    'Temporary files from file move operations
    DelScrubTmp
    
Else
    Log "Skipping CleanUp because bypass was requested."
End If

If Not sMoveMessage = "" Then Log vbCrLf & "Please remove this folder after next reboot: " & sMoveMessage
If fRebootRequired Then Log vbCrLf & "A restart is required to complete the operation."

'THE END
Log vbCrLf & "End removal: " & Now & vbCrLf
Log vbCrLf & "For detailed logging please refer to the log in folder " &chr(34)&sScrubDir&chr(34)&vbCrLf
'=======================================================================================================
'=======================================================================================================

'Stage 0 - 4 Subroutines
'=======================================================================================================

'Office 2003 products are listed with their ProductCode in the "Uninstall" key
Sub FindInstalledO11Products
    Dim ArpItem, prod, Item
    Dim sConfigName
    Dim arrKeys
    
    If dicInstalledProd.Count > 0 Then Exit Sub 'Already done from InputBox prompt
    
    'Query msi to get a list of Office 2003 products
    EnsureValidWIMetadata HKCU,"Software\Classes\Installer\Products",COMPRESSED
    EnsureValidWIMetadata HKCR,"Installer\Products",COMPRESSED
    EnsureValidWIMetadata HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products",COMPRESSED
    For Each prod in oMsi.Products
        If UCase(Right(prod,PRODLEN)) = OFFICE_2003 Then
            sConfigName = ""
            sConfigName = UCase(GetProductID(Mid(prod,4,2)))
            If Not dicInstalledProd.Exists(prod) Then dicInstalledProd.Add UCase(prod),sConfigName
        End If 'OFFICE_2003
    Next 'prod
    
    'Locate Office 2003 products from ARP
    If RegEnumKey(HKLM,REG_ARP,arrKeys) Then
        For Each ArpItem in arrKeys
            If UCase(Right(ArpItem,PRODLEN)) = OFFICE_2003 Then
                sConfigName = ""
                sConfigName = UCase(GetProductID(Mid(ArpItem,4,2)))
                If Not dicInstalledProd.Exists(ArpItem) Then dicInstalledProd.Add UCase(ArpItem),sConfigName
            End If
        Next 'ArpItem
    End If 'RegEnumKey
End Sub 'FindInstalledO11Products
'=======================================================================================================

'Create clean list of Products to remove.
'Strip off bad & empty contents
Sub ValidateRemoveSkuList
    Dim Sku, InstalledSku, sSkuKeepList, Keys, Key
    Dim iPos
    
    If fRemoveAll Then
        'Remove all mode
        For Each Key in dicInstalledProd.Keys
            Select Case dicInstalledProd.Item(Key)
            'Override default to remove Project server
            Case  "PRJSRV"
                dicKeepProd.Add Key,dicInstalledProd.Item(Key)
                fRemoveAll = False
            Case Else
                dicRemoveProd.Add Key,dicInstalledProd.Item(Key)
            End Select
        Next 'Key
    Else
        'Remove individual products mode
        
        'Ensure to have a string with no unexpected contents
        sSkuRemoveList = Replace(sSkuRemoveList,";",",")
        sSkuRemoveList = Replace(sSkuRemoveList," ","")
        sSkuRemoveList = Replace(sSkuRemoveList,Chr(34),"")
        While InStr(sSkuRemoveList,",,")>0
            sSkuRemoveList = Replace(sSkuRemoveList,",,",",")
        Wend
        
        'Prepare 'remove' and 'keep' dictionaries to determine what has to be removed
        
        'Initial pre-fill of 'keep' dic
        For Each Key in dicInstalledProd.Keys
            dicKeepProd.Add Key,dicInstalledProd.Item(Key)
        Next 'Key
        
        'Determine contents of keep and remove dic
        arrRemoveSKUs = Split(UCase(sSkuRemoveList),",")
        For Each Sku in arrRemoveSKUs
            If Sku = "OSE" Then fRemoveOse = True
            If dicKeepProd.Exists(Sku) Then
                'A productcode has been passed in
                'remove the item from the keep dic
                dicKeepProd.Remove(Sku)
                
                'Now add it to the remove dic
                dicRemoveProd.Add Sku,Sku
            End If
            'Check the Sku based entries
            For Each Key in dicKeepProd.Keys
                If dicKeepProd.Item(Key) = Sku Then
                    dicRemoveProd.Add Sku,Sku
                    dicKeepProd.Remove(Key)
                End If
            Next 'Key
        Next 'Sku
        
        If NOT dicKeepProd.Count > 0 Then fRemoveAll = True
        
    End If 'fRemoveAll
        
    If fRemoveAll OR fRemoveOSE Then CheckRemoveOSE
End Sub 'ValidateRemoveSkuList
'=======================================================================================================

'Check if OSE service can be scrubbed
Sub CheckRemoveOSE
    Dim Product,Service,Services
    Dim sOsePath
    
    'Keep OSE if a later version than 11.x is found
    Set Services = oWmiLocal.Execquery("Select * From Win32_Service Where Name='OSE'")
    For Each Service in Services
        sOsePath = Replace(Service.PathName,chr(34),"")
        If oFso.FileExists(sOsePath) Then
            If CInt(Left(oFso.GetFileVersion(sOsePath),2)) > 11 Then
                LogOnly "Disallowing OSE removal. OSE version found: " & oFso.GetFileVersion(sOsePath)
                fRemoveOse = False
                Exit Sub
            End If
        End If
    Next 'Service
    fRemoveOSE = True
End Sub 'CheckRemoveOSE
'=======================================================================================================

'Cache .msi files for products that will be removed in case they are needed for later file detection
Sub CacheMsiFiles
    Dim Product
    Dim sMsiFile
    
    'Non critical routine for failures.
    'Errors will be logged but must not fail the execution
    On Error Resume Next
    Log " Cache .msi files to temporary Scrub folder"
    'Cache the files
    EnsureValidWIMetadata HKCU,"Software\Classes\Installer\Products",COMPRESSED
    EnsureValidWIMetadata HKCR,"Installer\Products",COMPRESSED
    EnsureValidWIMetadata HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products",COMPRESSED
    For Each Product in oMsi.Products
        'Ensure valid GUID length
        If Len(Product) = 38 Then
            If (Right(Product,PRODLEN) = OFFICE_2003) AND (fRemoveAll OR CheckDelete(Product))Then
                CheckError "CacheMsiFiles"
                sMsiFile = oMsi.ProductInfo(Product,"LocalPackage") : CheckError "CacheMsiFiles"
                LogOnly " - " & Product & ".msi"
                If oFso.FileExists(sMsiFile) Then oFso.CopyFile sMsiFile,sScrubDir & "\" & Product & ".msi",True
                CheckError "CacheMsiFiles"
            End If  'OFFICE_2003
        End If '38
    Next 'Product
    Err.Clear
End Sub 'CacheMsiFiles
'=======================================================================================================

'Build a list of all files that will be deleted
Sub BuildFileList
    Const MSIOPENDATABASEREADONLY   = 0

    Dim FileList, ComponentID, CompClient, Record, qView, MsiDb
    Dim sQuery, sSubKeyName, sPath, sFile, sMsiFile, sCompClient, sComponent
    Dim fRemoveComponent
    Dim i, iProgress, iCompCnt
    Dim dicFLError, oDic, oFolderDic
    
    'Logfile
    Set FileList = oFso.OpenTextFile(sScrubDir & "\FileList.txt",FOR_WRITING,True,True)
    
    'FileListError dic
    Set dicFLError = CreateObject("Scripting.Dictionary")
    
    Set oDic = CreateObject("Scripting.Dictionary")
    Set oFolderDic = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    EnsureValidWIMetadata HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components",COMPRESSED
    EnsureValidWIMetadata HKCR,"Installer\Components",COMPRESSED
    iCompCnt = oMsi.Components.Count
    If NOT Err = 0 Then
        'API failure
        Err.Clear
        Exit Sub
    End If
    'Ensure to not divide by zero
    If iCompCnt = 0 Then iCompCnt = 1
    'Enum all Components
    For Each ComponentID In oMsi.Components
        'Progress bar
        i = i + 1
        If iProgress < (i / iCompCnt) * 100 Then 
            wscript.stdout.write "." : LogStream.Write "."
            iProgress = iProgress + 1
            If iProgress = 35 OR iProgress = 70 Then Log ""
        End If
        fRemoveComponent = False
        'Check if all ComponentClients will be removed
        For Each CompClient In oMsi.ComponentClients(ComponentID)
            If Err = 0 Then
                'Ensure valid guid length
                If Len(CompClient) = 38 Then
                    fRemoveComponent = Right(CompClient,PRODLEN)=OFFICE_2003 AND CheckDelete(CompClient)
                    If Not fRemoveComponent Then Exit For
                    'In "force" mode all components will be removed regardless of msidbComponentAttributesPermanent flag.
                    'Default is to honour the msidbComponentAttributesPermanent attribute and keep the files
                    If Not fForce Then
                        sSubKeyName = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\"
                        If RegValExists(HKLM,sSubKeyName & GetCompressedGuid(CompClient),COMPPERMANENT) Then
                            fRemoveComponent = False
                            Exit For
                        End If
                    End If 'fForce
                    sCompClient = CompClient
                Else
                    If NOT dicFLError.Exists("Error: Invalid metadata found. ComponentID: "&ComponentID &", ComponentClient: "&CompClient) Then _
                        dicFLError.Add "Error: Invalid metadata found. ComponentID: "&ComponentID &", ComponentClient: "&CompClient, ComponentID
                End If '38
            Else
                Err.Clear
            End If 'Err = 0
        Next 'CompClient

        If fRemoveComponent Then
            Err.Clear
            'Get the component path
            sPath = oMsi.ComponentPath(sCompClient,ComponentID)
            If oFso.FileExists(sPath) Then
                sPath = oFso.GetFile(sPath).ParentFolder
                If Not oFolderDic.Exists(sPath) Then oFolderDic.Add sPath,sPath
                'Get the .msi file
                If oFso.FileExists(sScrubDir & "\" & sCompClient & ".msi") Then
                    sMsiFile = sScrubDir & "\" & sCompClient & ".msi"
                Else
                    sMsiFile = oMsi.ProductInfo(sCompClient,"LocalPackage")
                End If
                If Not Err = 0 Then
                    If NOT dicFLError.Exists("Failed to obtain .msi file for product "&sCompClient) Then _
                        dicFLError.Add "Failed to obtain .msi file for product "&sCompClient, ComponentID
                    Err.Clear
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
                        If Len(sFile)>4 Then
                            If LCase(Right(sFile,4))=".exe" Then 
                                If NOT dicApps.Exists(LCase(sFile)) Then dicApps.Add LCase(sFile),LCase(sPath & "\" & sFile)
                            End If
                        End If
                        sFile = sPath & "\" & sFile
                        If Not oDic.Exists(sFile) Then 
                            oDic.Add sFile,sFile
                            FileList.WriteLine sFile
                        End If
                        Set Record = qView.Fetch()
                    Loop
                    Set Record = Nothing
                    qView.Close
                    Set qView = Nothing
                Else
                    If NOT dicFLError.Exists("Error: Could not read from .msi file: "&sMsiFile) Then _
                        dicFLError.Add "Error: Could not read from .msi file: "&sMsiFile, ComponentID
                    Err.Clear
                End If 'Err = 0
            End If 'FileExists(sPath)
        Else
            'Add the path to the 'KeepFolder' dictionary
            Err.Clear
            For Each CompClient In oMsi.ComponentClients(ComponentID)
                'Get the component path
                sPath = "" : sPath = LCase(oMsi.ComponentPath(CompClient,ComponentID))
                If oFso.FileExists(sPath) Then
                    sPath = LCase(oFso.GetFile(sPath).ParentFolder)
                    If Not dicKeepFolder.Exists(sPath) Then AddKeepFolder sPath
                End If
            Next 'CompClient
        End If 'fRemoveComponent
    Next 'ComponentID
    Err.Clear
    On Error Goto 0
    Log " Done"
    If dicFLError.Count > 0 Then LogOnly Join(dicFLError.Keys,vbCrLf)
    If Not oFolderDic.Count = 0 Then arrDeleteFolders = oFolderDic.Keys Else Set arrDeleteFolders = Nothing
    If Not oDic.Count = 0 Then arrDeleteFiles = oDic.Keys Else Set arrDeleteFiles = Nothing
End Sub 'BuildFileList
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
    Dim sCmd, sReturn
    
    'Check registered products
    'Removal can only happen for per machine and current user context -> use Installer.Products object
    i = 0
    EnsureValidWIMetadata HKCU,"Software\Classes\Installer\Products",COMPRESSED
    EnsureValidWIMetadata HKCR,"Installer\Products",COMPRESSED
    EnsureValidWIMetadata HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products",COMPRESSED
    For Each Product in oMsi.Products
        If (Right(UCase(Product),PRODLEN) = OFFICE_2003) Then
            If fRemoveAll OR CheckDelete(Product)Then
                i = i + 1 
                Log " Removing product " & Product
                sCmd = "msiexec.exe /x" & Product & " REBOOT=ReallySuppress"
                If bQuiet Then 
                    sCmd = sCmd & " /q"
                Else
                    sCmd = sCmd & " /qb-"
                End If
                sCmd = sCmd & " /l*v+ "&chr(34)&sLogDir&"\Uninstall_"&Product&".log"&chr(34)
                If NOT fDetectOnly Then 
                    Log " - Calling msiexec with '"&sCmd&"'"
                    'Execute the patch uninstall
                    sReturn = oWShell.Run(sCmd, 0, True)
                    Log " - msiexec returned: " & SetupRetVal(sReturn) & " (" & sReturn & ")" & vbCrLf
                    fRebootRequired = fRebootRequired OR (sReturn = "3010")
                Else
                    Log " -> Removal suppressed in preview mode. Command: "&sCmd
                End If
            End If
        End If 'OFFICE_2003
    Next 'Product
    If i = 0 Then Log "Nothing to remove for msiexec"
End Sub 'MsiexecRemoval
'=======================================================================================================

'Remove the OSE (Office Source Engine) service
Sub RemoveOSE
    On Error Resume Next
    Log " OSE CleanUp"
    DeleteService "ose"
    'Delete the folder
    DeleteFolder sCommonProgramFiles & "\Microsoft Shared\Source Engine"
    'Delete the registration
    RegDeleteKey HKLM,"SYSTEM\CurrentControlSet\Services\ose"
End Sub 'RemoveOSE
'=======================================================================================================

'Identify which parts of the LIS (MSOCache) will not be removed
Sub CheckLIS

    Dim Prod
    Dim sDownloadCode
    
    If NOT dicKeepProd.Count > 0 Then Exit Sub

    'Loop all products that remain installed
    EnsureValidWIMetadata HKCU,"Software\Classes\Installer\Products",COMPRESSED
    EnsureValidWIMetadata HKCR,"Installer\Products",COMPRESSED
    EnsureValidWIMetadata HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products",COMPRESSED
    For Each Prod in dicKeepProd.Keys
        If RegReadValue(HKLM,DELIVERY&Prod,"DownloadCode",sDownloadCode,"REG_SZ") Then
            If dicKeepLis.Exists(UCase(sDownloadCode)) Then 
                dicKeepLis.Item(sDownloadCode) = dicKeepLis.Item(sDownloadCode)&","&UCase(Prod)
            Else
                dicKeepLis.Add UCase(sDownloadCode),UCase(Prod)
            End If
        End If
    Next 'Prod

End Sub 'CheckLIS
'=======================================================================================================

'File cleanup operations for the Local Installation Source (MSOCache)
Sub WipeLIS
    Const LISROOT = "MSOCache\All Users\"
    Dim LogicalDisks, Disk, Folder, SubFolder, MseFolder, File, Files
    Dim arrSubFolders
    Dim sFolder
    Dim fRemoveFolder
    
    Log " LIS CleanUp"
    'Search all hard disks
    Set LogicalDisks = oWmiLocal.ExecQuery("Select * from Win32_LogicalDisk")
    For Each Disk in LogicalDisks
        If Disk.DriveType = 3 Then
            If oFso.FolderExists(Disk.DeviceID & "\" & LISROOT)Then
                Set Folder = oFso.GetFolder(Disk.DeviceID & "\" & LISROOT)
                For Each Subfolder in Folder.Subfolders
                    If InStr(UCase(Subfolder.Name)&"}",OFFICE_2003)>0 Then
                        If NOT dicKeepLis.Exists(UCase(Subfolder.Name)) Then DeleteFolder Subfolder.Path
                    ElseIf LCase(Subfolder.Name) = "microsoft.watson.alrtintl.data" OR _
                       LCase(Subfolder.Name) = "microsoft.watson.watsonrc.data" Then
                        If NOT dicKeepLis.Count > 0 Then DeleteFolder Subfolder.Path
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
                GetMseFolderStructure Folder
                For Each MseFolder in arrMseFolders
                    If oFso.FolderExists(MseFolder) Then
                        fRemoveFolder = False
                        Set Folder = oFso.GetFolder(MseFolder)
                        Set Files = Folder.Files
                        For Each File in Files
                            If (LCase(Right(File.Name,4))=".msi") Then
                                If CheckDelete(ProductCode(File.Path)) Then 
                                    fRemoveFolder = True
                                    Exit For
                                End If 'CheckDelete
                            End If
                        Next 'File
                        Set Files = Nothing
                        Set Folder = Nothing
                        If fRemoveFolder Then SmartDeleteFolder MseFolder
                    End If 'oFso.FolderExists(MseFolder)
                Next 'MseFolder
            End If
        Next 'SubFolder
    End If 'oFso.FolderExists
End Sub 'WipeLis
'=======================================================================================================

'Wipe files and folders
Sub FileWipeAll
    
    'User specific files
    If NOT fKeepUser Then
        'Delete files that should be backed up before deleting them
        CopyAndDeleteFile sAppdata & "\Microsoft\Templates\Normal.dot"
    End If
    
    'Run the individual filewipe first
    FileWipeIndividual
    
    'Take care of the rest
    DeleteFolder sOInstallRoot
    DeleteFolder sCommonProgramFiles & "\Microsoft Shared\Office11"
    DeleteFile sAllUsersProfile & "\Application Data\Microsoft\Office\Data\opa11.dat"

End Sub 'FileWipeAll
'=======================================================================================================

'Wipe individual files & folders related to SKU's that are no longer installed
Sub FileWipeIndividual
    Dim LogicalDisks, Disk
    Dim File, Files, XmlFile, scFiles, oFile, Folder, SubFolder, Processes, Process, item
    Dim sFile, sFolder, sPath, sConfigName, sContents, sProductCode, sLocalDrives,sScQuery
    Dim arrSubfolders
    Dim fDeleteSC
    
    Log " File CleanUp"
    If IsArray(arrDeleteFiles) Then
        If fForce Then
            Log " Doing Action: EndOSE"
            Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process")
            For Each Process in Processes
                LogOnly " - Running process : " & Process.Name
                If Len(Process.Name)>2 Then
                    If LCase(Left(Process.Name,3))="ose" Then 
                        Log " -> Ending process: " & Process.Name
                        Process.Terminate
                    End If
                End If 'Len>2
            Next 'Process
            LogOnly " End Action: EndOSE"
            CloseOfficeApps
        End If
        'Wipe individual files detected earlier
        LogOnly " Removing left behind files"
        For Each sFile in arrDeleteFiles
            If oFso.FileExists(sFile) Then DeleteFile sFile
        Next 'File
    End If 'IsArray
    
    'Wipe Shortcuts from local hard disks
    If NOT fSkipSD Then
        On Error Resume Next
        Log " Searching for shortcuts. This can take some time ..."
        Set LogicalDisks = oWmiLocal.ExecQuery("Select * From Win32_LogicalDisk WHERE DriveType=3")
        For Each Disk in LogicalDisks
            sLocalDrives = sLocalDrives & UCase(Disk.DeviceID) & "\;"
            sScQuery = "Select * From Win32_ShortcutFile WHERE Drive='"&Disk.DeviceID&"'"
            Set scFiles = oWmiLocal.ExecQuery(sScQuery)
            For Each File in scFiles
                fDeleteSC = False
                'Compare if the shortcut target is in the list of executables that will be removed
                If Len(File.Target)>0 Then
                    For Each item in dicApps.Items
                        If LCase(File.Target) = item Then
                            fDeleteSC = True
                            Exit For
                        End If
                    Next 'item
                End If
                'Handle Windows Installer shortcuts
                If InStr(File.Target,"{")>0 Then
                    If Len(File.Target)>=InStr(File.Target,"{")+37 Then
                        If CheckDelete(Mid(File.Target,InStr(File.Target,"{"),38)) Then fDeleteSC = True
                    End If
                End If
                If fDeleteSC Then 
                    If Not IsArray(arrDeleteFolders) Then ReDim arrDeleteFolders(0)
                    sFolder = Left(File.Description,InStrRev(File.Description,"\")-1)
                    If Not arrDeleteFolders(UBound(arrDeleteFolders)) = sFolder Then
                        ReDim Preserve arrDeleteFolders(UBound(arrDeleteFolders)+1)
                        arrDeleteFolders(UBound(arrDeleteFolders)) = sFolder
                    End If
                    DeleteFile File.Description
                End If 'fDeleteSC
            Next 'scFile
        Next
        On Error Goto 0
    End If 'NOT SkipSD
    Err.Clear

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
    
    Log " ScrubCache CleanUp"
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

    'Error handling inlined
    On Error Resume Next

    Dim Patch, AllPatches, Product, AllProducts
    Dim File, Files, Folder
    Dim sFName, sLocalMsp, sLocalMsi, sPatchList, sMsiList

    Set Folder = oFso.GetFolder(sWinDir & "\Installer")
    Set Files = Folder.Files

    Log " Windows Installer cache CleanUp"
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
                    If InStr(UCase(MspTargets(File.Path)),OFFICE_2003)>0 Then DeleteFile File.Path
                End If
            End If 'LCase(Right(sFName,4))
        Next 'File
    End If 'Err=0

    'Get a complete list products
    Err.Clear
    EnsureValidWIMetadata HKCU,"Software\Classes\Installer\Products",COMPRESSED
    EnsureValidWIMetadata HKCR,"Installer\Products",COMPRESSED
    EnsureValidWIMetadata HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products",COMPRESSED
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
                    If UCase(Right(ProductCode(File.Path),PRODLEN))=OFFICE_2003 Then DeleteFile File.Path
                End If
            End If 'LCase(Right(sFName,4)) = ".msi"
        Next 'File
    End If 'Err=0

End Sub 'MsiClearOrphanedFiles
'=======================================================================================================

Sub RegWipe
    Dim Item, Name, Sku
    Dim hDefKey, sSubKeyName, sCurKey, sValue, sGuid
    Dim arrKeys, arrNames, arrTypes
    Dim i, iLoopCnt
    
    Log " Registry CleanUp"
    'Wipe registry data
    
    'User Profile settings
    RegDeleteKey HKCU,"Software\Policies\Microsoft\Office\11.0"
    If fRemoveAll Then
        RegDeleteKey HKCU,"SOFTWARE\Microsoft\OfficeCustomizeWizard\11.0"
    End If
    If NOT fKeepUser Then
        RegDeleteKey HKCU,"Software\Microsoft\Office\11.0"
    End If 'fKeepUser
    
    'Computer specific settings
    If fRemoveAll Then
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Office\11.0"
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Terminal Server\Install\Software\Microsoft\Office\11.0"
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\OfficeCustomizeWizard\11.0"
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Terminal Server\Install\SOFTWARE\Microsoft\OfficeCustomizeWizard\11.0"
        
        'Jet_Replication
        sValue = ""
        If RegReadValue(HKCR,"CLSID\{CC2C83A6-9BE4-11D0-98E7-00C04FC2CAF5}\InprocServer32","SystemDB",sValue,"REG_SZ") Then
            If Len(sValue) > Len(sOInstallRoot) Then
                If LCase(Left(sValue,Len(sOInstallRoot))) = LCase(sOInstallRoot) Then RegDeleteKey HKCR,"CLSID\{CC2C83A6-9BE4-11D0-98E7-00C04FC2CAF5}\InprocServer32"
            End If
        End If
        
        'Win32Assemblies
        hDefKey = HKCR
        sSubKeyName  = "Installer\Win32Assemblies\"
        If RegEnumKey(hDefKey,sSubKeyName,arrKeys) Then
            For Each Item in arrKeys
                If InStr(UCase(Item),"OFFICE11")>0 Then RegDeleteKey hDefKey,sSubKeyName & Item
            Next 'Item
        End If 'RegEnumKey
    End If 'fRemoveAll
    
    For iLoopCnt = 1 to 3
        Select Case iLoopCnt
        Case 1
            'CIW - HKCU
            sSubKeyName = "Software\Microsoft\OfficeCustomizeWizard\11.0\RegKeyPaths\"
            hDefKey = HKCU
        Case 2 
            'CIW - HKLM
            sSubKeyName = "SOFTWARE\Microsoft\OfficeCustomizeWizard\11.0\RegKeyPaths\"
            hDefKey = HKLM
        Case 3
            'Add/Remove Programs
            sSubKeyName = REG_ARP
            hDefKey = HKLM
        End Select
        
        If RegEnumKey(hDefKey,sSubKeyName,arrKeys) Then
            For Each Item in arrKeys
                'OFFICE_2003 id
                If Len(Item)>37 Then
                    sGuid = UCase(Left(Item,38))
                    If Right(sGuid,PRODLEN)=OFFICE_2003 Then
                        If CheckDelete(sGuid) Then 
                            RegDeleteKey hDefKey, sSubKeyName & Item
                        End If
                    End If 'Right(Item,PRODLEN)=OFFICE_2003
                End If 'Len(Item)>37
            Next 'Item
            If iLoopCnt < 3 Then
                If RegEnumValues(hDefKey,sSubKeyName,arrNames,arrTypes) Then
                    i = 0
                    For Each Name in arrNames
                        If RegReadValue(hDefKey,sSubKeyName,Name,sValue,arrTypes(i)) Then
                            If sValue = sGuid Then RegDeleteValue hDefKey,sSubKeyName,Name
                        End If
                        i = i + 1
                    Next
                End If
            End If
        End If
        If NOT RegEnumKey(hDefKey,sSubKeyName,arrKeys) Then RegDeleteKey hDefKey,"Software\Microsoft\OfficeCustomizeWizard\11.0\"
        If NOT RegEnumKey(hDefKey,"Software\Microsoft\OfficeCustomizeWizard\11.0\",arrKeys) Then RegDeleteKey hDefKey,"Software\Microsoft\OfficeCustomizeWizard\"
    Next 'iLoopCnt
    
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
                    'Check if it's a Office 2003 key
                    If Right(sGuid,PRODLEN)=OFFICE_2003 Then
                        If fRemoveAll Then
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
                        End If 'fRemoveAll
                    End If 'Right(Item,PRODLEN)=OFFICE_2003
                End If 'Len(Item)=32
            Next 'Item
        End If 'RegEnumKey
    Next 'iLoopCnt

    'Delivery
    hDefKey = HKLM
    sSubKeyName = "SOFTWARE\Microsoft\Office\Delivery\SourceEngine\Downloads\"
    If RegEnumKey(HKLM,sSubKeyName,arrKeys) Then
        For Each Item in arrKeys
            If InStr(UCase(Item)&"}",OFFICE_2003)>0 Then
                If NOT dicKeepLis.Exists(UCase(Item)) Then RegDeleteKey HKLM,sSubKeyName & Item
            ElseIf LCase(Item) = "microsoft.watson.alrtintl.data" OR _
                   LCase(Item) = "microsoft.watson.watsonrc.data" Then
                    If NOT dicKeepLis.Count > 0 Then RegDeleteKey HKLM,sSubKeyName & Item
            End If
        Next 'Item
    End If 'RegEnumKey
    
    'Registration
    hDefKey = HKLM
    sSubKeyName = "SOFTWARE\Microsoft\Office\11.0\Registration\"
    If RegEnumKey(HKLM,sSubKeyName,arrKeys) Then
        For Each Item in arrKeys
            If Len(Item)>37 Then
                If CheckDelete(UCase(Left(Item,38))) Then RegDeleteKey HKLM,sSubKeyName & Item
            End If
        Next 'Item
    End If 'RegEnumKey
    
End Sub 'RegWipeAll
'=======================================================================================================

'=======================================================================================================
' Helper Functions
'=======================================================================================================

'Kill all running instances of applications that will be removed
Sub CloseOfficeApps
    Dim Processes, Process
    
    Log " Doing Action: CloseOfficeApps"
    Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process")
    For Each Process in Processes
        If dicApps.Exists(LCase(Process.Name)) Then
            Log " - End process " & Process.Name
            Process.Terminate()
            CheckError "CloseOfficeApps: " & "Process.Name"
        End If
    Next 'Process
    LogOnly " End Action: CloseOfficeApps"
End Sub 'CloseOfficeApps
'=======================================================================================================

'Ensure Windows Explorer is restarted if needed
Sub RestoreExplorer
    Dim Processes
    
    'Non critical routine. Don't fail on error
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
    Dim fReturn

    CheckRegPermissions = True
    sSubKeyName = "Software\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\"
    oReg.CheckAccess HKLM, sSubKeyName, KEY_QUERY_VALUE, fReturn
    If Not fReturn Then CheckRegPermissions = False
    oReg.CheckAccess HKLM, sSubKeyName, KEY_SET_VALUE, fReturn
    If Not fReturn Then CheckRegPermissions = False
    oReg.CheckAccess HKLM, sSubKeyName, KEY_CREATE_SUB_KEY, fReturn
    If Not fReturn Then CheckRegPermissions = False
    oReg.CheckAccess HKLM, sSubKeyName, DELETE, fReturn
    If Not fReturn Then CheckRegPermissions = False

End Function 'CheckRegPermissions
'=======================================================================================================

'Check if an Office 11 product is still registered with a SKU that stays on the computer
Function CheckDelete(sProductCode)
    Dim Sku
    Dim RetVal
    Dim sProductCodeList
        
    'If it's a non Office 11 ProductCode exit with false right away
    CheckDelete = Right(sProductCode,PRODLEN) = OFFICE_2003
    If Not CheckDelete OR NOT (dicKeepProd.Count > 0) Then Exit Function
    If dicKeepProd.Exists(UCase(sProductCode)) Then CheckDelete = False
    
End Function 'CheckDelete
'=======================================================================================================

'Returns a string with a list of ProductCodes from the summary information stream
Function MspTargets (sMspFile)
    Const MSIOPENDATABASEMODE_PATCHFILE = 32
    Const PID_TEMPLATE                  =  7
    
    Dim Msp
    'Non critical routine. Don't fail on error
    On Error Resume Next
    MspTargets = ""
    If oFso.FileExists(sMspFile) Then
        Set Msp = Msi.OpenDatabase(WScript.Arguments(0),MSIOPENDATABASEMODE_PATCHFILE)
        If Err = 0 Then MspTargets = Msp.SummaryInformation.Property(PID_TEMPLATE)
    End If 'oFso.FileExists(sMspFile)
End Function 'MspTargets
'=======================================================================================================

'Return the ProductCode {GUID} from a .MSI package
Function ProductCode(sMsi)
    Const MSIUILEVELNONE = 2 'No UI
    Dim MsiSession

    On Error Resume Next
    'Non critical routine. Don't fail on error
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

'Ensures that only valid metadata entries exist to avoid API failures
Sub EnsureValidWIMetadata (hDefKey,sKey,iValidLength)

Dim arrKeys
Dim SubKey

If Len(sKey) > 1 Then
    If Right(sKey,1) = "\" Then sKey = Left(sKey,Len(sKey)-1)
End If

If RegEnumKey(hDefKey,sKey,arrKeys) Then
    For Each SubKey in arrKeys
        If NOT Len(SubKey) = iValidLength Then
            RegDeleteKey hDefKey,sKey & "\" & SubKey
        End If
    Next 'SubKey
End If

End Sub 'EnsureValidWIMetadata
'=======================================================================================================

'Create a backup copy of the file in the ScrubDir then delete the file
Sub CopyAndDeleteFile(sFile)
    Dim File
    
    'Error handling inlined
    On Error Resume Next
    If oFso.FileExists(sFile) Then
        Set File = oFso.GetFile(sFile)
        If Not oFso.FolderExists(sScrubDir & "\" & File.ParentFolder.Name) Then oFso.CreateFolder sScrubDir & "\" & File.ParentFolder.Name
        If Not fDetectOnly Then
            LogOnly " - Backing up file: " & sFile
            oFso.CopyFile sFile,sScrubDir & "\" & File.ParentFolder.Name & "\" & File.Name,True : CheckError "CopyAndDeleteFile"
            Set File = Nothing
            DeleteFile(sFile)
        Else
            LogOnly " - Simulate CopyAndDelete file: " & sFile
        End If
    End If 'oFso.FileExists
End Sub 'CopyAndDeleteFile
'=======================================================================================================

'Wrapper to delete a file
Sub DeleteFile(sFile)
    Dim File
    Dim sFileName, sNewPath
    
    'Error handling inlined
    On Error Resume Next
    If oFso.FileExists(sFile) Then
        If Not fDetectOnly Then
            LogOnly " - Delete file: " & sFile
            oFso.DeleteFile sFile,True 
        Else
            LogOnly " - Simulate delete file: " & sFile
        End If
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
            LogOnly " - Move file to: " & sNewPath & "\" & sFileName
            oFso.MoveFile sFile,sNewPath & "\" & sFileName
            If Err <> 0 Then 
                CheckError "DeleteFile (move)"
            Else
                If Not InStr(sMoveMessage,sNewPath)>0 Then sMoveMessage = sMoveMessage & sNewPath & ";"
                oFso.DeleteFile sNewPath & "\" & sFileName,True 
                If Err <> 0 Then CheckError "DeleteFile (moved)"
            End If 'Err <> 0
        End If 'Err <> 0
    End If 'oFso.FileExists
End Sub 'DeleteFile
'=======================================================================================================

'64 bit aware wrapper to return the requested folder 
Function GetFolderPath(sPath)
    GetFolderPath = True
    If oFso.FolderExists(sPath) Then Exit Function
    If f64 AND oFso.FolderExists(Wow64Folder(sPath)) Then
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
    If f64 AND oFso.FolderExists(Wow64Folder(sFolder)) Then
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
    If f64 AND oFso.FolderExists(Wow64Folder(sFolder)) Then
        Set Folder = oFso.GetFolder(Wow64Folder(sFolder))
        For Each Subfolder in Folder.Subfolders
            sSubFolders = sSubFolders & Subfolder.Path & ","
        Next 'Subfolder
    End If
    If Len(sSubFolders)>0 Then arrSubFolders = RemoveDuplicates(Split(Left(sSubFolders,Len(sSubFolders)-1),","))
    EnumFolders = Len(sSubFolders)>0
End Function 'EnumFolders
'=======================================================================================================

Sub GetMseFolderStructure (Folder)
    Dim SubFolder
    
    For Each SubFolder in Folder.SubFolders
        ReDim Preserve arrMseFolders(UBound(arrMseFolders)+1)
        arrMseFolders(UBound(arrMseFolders)) = SubFolder.Path
        GetMseFolderStructure SubFolder
    Next 'SubFolder
End Sub 'GetMseFolderStructure
'=======================================================================================================

'Wrapper to delete a folder 
Sub DeleteFolder(sFolder)
    Dim Folder
    Dim sDelFolder, sFolderName, sNewPath
    
    If dicKeepFolder.Exists(LCase(sFolder)) Then Exit Sub
    If f64 Then
        If dicKeepFolder.Exists(LCase(Wow64Folder(sFolder))) Then Exit Sub
    End If
    If Len(sFolder) > 1 Then
        If Right(sFolder,1) = "\" Then sFolder = Left(sFolder,Len(sFolder)-1)
    End If
    
    'Error handling inlined
    On Error Resume Next
    If oFso.FolderExists(sFolder) Then 
        sDelFolder = sFolder
    ElseIf f64 AND oFso.FolderExists(Wow64Folder(sFolder)) Then 
        sDelFolder = Wow64Folder(sFolder)
    Else
        Exit Sub
    End If
    If Not fDetectOnly Then 
        LogOnly " - Delete folder: " & sDelFolder
        oFso.DeleteFolder sDelFolder,True
    Else
        LogOnly " - Simulate delete folder: " & sDelFolder
    End If
    If Err <> 0 Then
        CheckError "DeleteFolder"
        'Try to move the folder and delete from there
        Set Folder = oFso.GetFolder(sDelFolder)
        sFolderName = Folder.Name
        sNewPath = Folder.Drive.Path & "\" & "ScrubTmp"
        Set Folder = Nothing
        'Ensure we stay within the same drive
        If Not oFso.FolderExists(sNewPath) Then oFso.CreateFolder(sNewPath)
        'Move the folder
        LogOnly " - Moving folder to: " & sNewPath & "\" & sFolderName
        oFso.MoveFolder sFolder,sNewPath & "\" & sFolderName
        If Err <> 0 Then
            CheckError "DeleteFolder (move)"
        Else
            oFso.DeleteFolder sNewPath & "\" & sFolderName,True 
            If Err <> 0 And fForce Then 
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
    Log " Empty Folder Cleanup"
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
        If Not fDetectOnly Then
            LogOnly "Request SmartDelete for folder: " & sFolder
            SmartDeleteFolderEx sFolder
        Else
            LogOnly "Simulate request SmartDelete for folder: " & sFolder
        End If
    End If
    If f64 AND oFso.FolderExists(Wow64Folder(sFolder)) Then 
        If Not fDetectOnly Then 
            LogOnly "Request SmartDelete for folder: " & Wow64Folder(sFolder)
            SmartDeleteFolderEx Wow64Folder(sFolder)
        Else
            LogOnly "Simulate request SmartDelete for folder: " & Wow64Folder(sFolder)
        End If
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

'Adds the folder structure to the 'KeepFolder' dictionary
Sub AddKeepFolder(sPath)

Dim Folder

If NOT dicKeepFolder.Exists (sPath) Then
    dicKeepFolder.Add sPath,sPath
End If
sPath = LCase(oFso.GetParentFolderName(sPath))
If oFso.FolderExists(sPath) Then AddKeepFolder(sPath)

End Sub
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
        Case "1","REG_SZ"
            RetVal = oReg.GetStringValue(hDefKey,sSubKeyName,sName,sValue)
            If Not RetVal = 0 AND f64 Then RetVal = oReg.GetStringValue(hDefKey,Wow64Key(hDefKey, sSubKeyName),sName,sValue)
        
        Case "2","REG_EXPAND_SZ"
            RetVal = oReg.GetExpandedStringValue(hDefKey,sSubKeyName,sName,sValue)
            If Not RetVal = 0 AND f64 Then RetVal = oReg.GetExpandedStringValue(hDefKey,Wow64Key(hDefKey, sSubKeyName),sName,sValue)
        
        Case "7","REG_MULTI_SZ"
            RetVal = oReg.GetMultiStringValue(hDefKey,sSubKeyName,sName,arrValues)
            If Not RetVal = 0 AND f64 Then RetVal = oReg.GetMultiStringValue(hDefKey,Wow64Key(hDefKey, sSubKeyName),sName,arrValues)
            If RetVal = 0 Then sValue = Join(arrValues,chr(34))
        
        Case "4","REG_DWORD"
            RetVal = oReg.GetDWORDValue(hDefKey,sSubKeyName,sName,sValue)
            If Not RetVal = 0 AND f64 Then 
                RetVal = oReg.GetDWORDValue(hDefKey,Wow64Key(hDefKey, sSubKeyName),sName,sValue)
            End If
        
        Case "3","REG_BINARY"
            RetVal = oReg.GetBinaryValue(hDefKey,sSubKeyName,sName,sValue)
            If Not RetVal = 0 AND f64 Then RetVal = oReg.GetBinaryValue(hDefKey,Wow64Key(hDefKey, sSubKeyName),sName,sValue)
        
        Case "11","REG_QWORD"
            RetVal = oReg.GetQWORDValue(hDefKey,sSubKeyName,sName,sValue)
            If Not RetVal = 0 AND f64 Then RetVal = oReg.GetQWORDValue(hDefKey,Wow64Key(hDefKey, sSubKeyName),sName,sValue)
        
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
    
    If f64 Then
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
    End If 'f64
    RegEnumValues = ((RetVal = 0) OR (RetVal64 = 0)) AND IsArray(arrNames) AND IsArray(arrTypes)
End Function 'RegEnumValues
'=======================================================================================================

'Enumerate a registry key to return all subkeys
Function RegEnumKey(hDefKey,sSubKeyName,arrKeys)
    Dim RetVal, RetVal64
    Dim arrKeys32, arrKeys64
    
    If f64 Then
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
    End If 'f64
    RegEnumKey = ((RetVal = 0) OR (RetVal64 = 0)) AND IsArray(arrKeys)
End Function 'RegEnumKey
'=======================================================================================================

'Wrapper around oReg.DeleteValue to handle 64 bit
Sub RegDeleteValue(hDefKey, sSubKeyName, sName)
    Dim sWow64Key
    Dim iRetVal
    
    If RegValExists(hDefKey,sSubKeyName,sName) Then
        On Error Resume Next
        If Not fDetectOnly Then 
            LogOnly " - Delete registry value: " & HiveString(hDefKey) & "\" & sSubKeyName & " -> " & sName
            iRetVal = 0
            iRetVal = oReg.DeleteValue(hDefKey, sSubKeyName, sName)
            CheckError "RegDeleteValue"
            If NOT (iRetVal=0) Then LogOnly "     Delete failed. Return value: "&iRetVal
        Else
            LogOnly " - Simulate delete registry value: " & HiveString(hDefKey) & "\" & sSubKeyName & " -> " & sName
        End If
        On Error Goto 0
    End If 'RegValExists
    If f64 Then 
        sWow64Key = Wow64Key(hDefKey, sSubKeyName)
        If RegValExists(hDefKey,sWow64Key,sName) Then
            On Error Resume Next
            If Not fDetectOnly Then 
            LogOnly " - Delete registry value: " & HiveString(hDefKey) & "\" & sWow64Key & " -> " & sName
                iRetVal = 0
                iRetVal = oReg.DeleteValue(hDefKey, sWow64Key, sName)
                CheckError "RegDeleteValue"
                If NOT (iRetVal=0) Then LogOnly "     Delete failed. Return value: "&iRetVal
            Else
                LogOnly " - Simulate delete registry value: " & HiveString(hDefKey) & "\" & sWow64Key & " -> " & sName
            End If
            On Error Goto 0
        End If 'RegKeyExists
    End If
End Sub 'RegDeleteValue
'=======================================================================================================

'Wrappper around RegDeleteKeyEx to handle 64bit scenrios
Sub RegDeleteKey(hDefKey, sSubKeyName)
    Dim sWow64Key
    
    If RegKeyExists(hDefKey, sSubKeyName) Then
        If Not fDetectOnly Then
            LogOnly " - Delete registry key: " & HiveString(hDefKey) & "\" & sSubKeyName
            On Error Resume Next
            RegDeleteKeyEx hDefKey, sSubKeyName
            On Error Goto 0
        Else
            LogOnly " - Simulate delete registry key: " & HiveString(hDefKey) & "\" & sSubKeyName
        End If
    End If 'RegKeyExists
    If f64 Then 
        sWow64Key = Wow64Key(hDefKey, sSubKeyName)
        If RegKeyExists(hDefKey,sWow64Key) Then
            If Not fDetectOnly Then
                LogOnly " - Delete registry key: " & HiveString(hDefKey) & "\" & sWow64Key
                On Error Resume Next
                RegDeleteKeyEx hDefKey, sWow64Key
                On Error Goto 0
            Else
                LogOnly " - Simulate delete registry key: " & HiveString(hDefKey) & "\" & sWow64Key
            End If
        End If 'RegKeyExists
    End If
End Sub 'RegDeleteKey
'=======================================================================================================

'Recursively delete a registry structure
Sub RegDeleteKeyEx(hDefKey, sSubKeyName) 
    Dim arrSubkeys
    Dim sSubkey
    Dim iRetVal

    On Error Resume Next
    oReg.EnumKey hDefKey, sSubKeyName, arrSubkeys
    If IsArray(arrSubkeys) Then 
        For Each sSubkey In arrSubkeys 
            RegDeleteKeyEx hDefKey, sSubKeyName & "\" & sSubkey 
        Next 
    End If 
    If Not fDetectOnly Then 
        iRetVal = 0
        iRetVal = oReg.DeleteKey(hDefKey,sSubKeyName)
        If NOT (iRetVal=0) Then LogOnly "     Delete failed. Return value: "&iRetVal
    End If
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

'Uses WMI to stop a service
Function StopService(sService)
    Dim Services, Service
    Dim sQuery
    
    On Error Resume Next
    
    sQuery = "Select * From Win32_Service Where Name='" & sService & "'"
    Set Services = oWmiLocal.Execquery(sQuery)
    'Stop the service
    For Each Service in Services
        If Service.State = "Started" Then Service.StopService
    Next 'Service
    Err.Clear
End Function 'StopService
'=======================================================================================================

'Delete a service
Sub DeleteService(sService)
    Dim Services, Service, Processes, Process
    Dim sQuery, sStates
    
    On Error Resume Next
    
    sStates = "STARTED;RUNNING"
    sQuery = "Select * From Win32_Service Where Name='" & sService & "'"
    Set Services = oWmiLocal.Execquery(sQuery)
    
    'Stop and delete the service
    For Each Service in Services
        Log " Found service " & sService & " in state " & Service.State
        If InStr(sStates,UCase(Service.State))>0 Then Service.StopService
        'Ensure no more instances of the service are running
        Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process Where Name='" & sService & ".exe'")
        For Each Process in Processes
            Process.Terminate
        Next 'Process
        If Not fDetectOnly Then 
            Log " - Deleting Service -> " & sService
            Service.Delete
        Else
            Log " - Simulate deleting Service -> " & sService
        End If
    Next 'Service
    Set Services = Nothing
    Err.Clear

End Sub 'DeleteService
'=======================================================================================================

'Translation for setup.exe error codes
Function SetupRetVal(RetVal)
    Select Case RetVal
        Case 0 : SetupRetVal = "Success"
        Case 30001,1 : SetupRetVal = "AbstractMethod"
        Case 30002,2 : SetupRetVal = "ApiProhibited"
        Case 30003,3  : SetupRetVal = "AlreadyImpersonatingAUser"
        Case 30004,4 : SetupRetVal = "AlreadyInitialized"
        Case 30005,5 : SetupRetVal = "ArgumentNullException"
        Case 30006,6 : SetupRetVal = "AssertionFailed"
        Case 30007,7 : SetupRetVal = "CABFileAddFailed"
        Case 30008,8 : SetupRetVal = "CommandFailed"
        Case 30009,9 : SetupRetVal = "ConcatenationFailed"
        Case 30010,10 : SetupRetVal = "CopyFailed"
        Case 30011,11 : SetupRetVal = "CreateEventFailed"
        Case 30012,12 : SetupRetVal = "CustomizationPatchNotFound"
        Case 30013,13 : SetupRetVal = "CustomizationPatchNotApplicable"
        Case 30014,14 : SetupRetVal = "DuplicateDefinition"
        Case 30015,15 : SetupRetVal = "ErrorCodeOnly - Passthrough for Win32 error"
        Case 30016,16 : SetupRetVal = "ExceptionNotThrown"
        Case 30017,17 : SetupRetVal = "FailedToImpersonateUser"
        Case 30018,18 : SetupRetVal = "FailedToInitializeFlexDataSource"
        Case 30019,19 : SetupRetVal = "FailedToStartClassFactories"
        Case 30020,20 : SetupRetVal = "FileNotFound"
        Case 30021,21 : SetupRetVal = "FileNotOpen"
        Case 30022,22 : SetupRetVal = "FlexDialogAlreadyInitialized"
        Case 30023,23 : SetupRetVal = "HResultOnly - Passthrough for HRESULT errors"
        Case 30024,24 : SetupRetVal = "HWNDNotFound"
        Case 30025,25 : SetupRetVal = "IncompatibleCacheAction"
        Case 30026,26 : SetupRetVal = "IncompleteProductAddOns"
        Case 30027,27 : SetupRetVal = "InstalledProductStateCorrupt"
        Case 30028,28 : SetupRetVal = "InsufficientBuffer"
        Case 30029,29 : SetupRetVal = "InvalidArgument"
        Case 30030,30 : SetupRetVal = "InvalidCDKey"
        Case 30031,31 : SetupRetVal = "InvalidColumnType"
        Case 30032,31 : SetupRetVal = "InvalidConfigAddLanguage"
        Case 30033,33 : SetupRetVal = "InvalidData"
        Case 30034,34 : SetupRetVal = "InvalidDirectory"
        Case 30035,35 : SetupRetVal = "InvalidFormat"
        Case 30036,36 : SetupRetVal = "InvalidInitialization"
        Case 30037,37 : SetupRetVal = "InvalidMethod"
        Case 30038,38 : SetupRetVal = "InvalidOperation"
        Case 30039,39 : SetupRetVal = "InvalidParameter"
        Case 30040,40 : SetupRetVal = "InvalidProductFromARP"
        Case 30041,41 : SetupRetVal = "InvalidProductInConfigXml"
        Case 30042,42 : SetupRetVal = "InvalidReference"
        Case 30043,43 : SetupRetVal = "InvalidRegistryValueType"
        Case 30044,44 : SetupRetVal = "InvalidXMLProperty"
        Case 30045,45 : SetupRetVal = "InvalidMetadataFile"
        Case 30046,46 : SetupRetVal = "LogNotInitialized"
        Case 30047,47 : SetupRetVal = "LogAlreadyInitialized"
        Case 30048,48 : SetupRetVal = "MissingXMLNode"
        Case 30049,49 : SetupRetVal = "MsiTableNotFound"
        Case 30050,50 : SetupRetVal = "MsiAPICallFailure"
        Case 30051,51 : SetupRetVal = "NodeNotOfTypeElement"
        Case 30052,52 : SetupRetVal = "NoMoreGraceBoots"
        Case 30053,53 : SetupRetVal = "NoProductsFound"
        Case 30054,54 : SetupRetVal = "NoSupportedCulture"
        Case 30055,55 : SetupRetVal = "NotYetImplemented"
        Case 30056,56 : SetupRetVal = "NotAvailableCulture"
        Case 30057,57 : SetupRetVal = "NotCustomizationPatch"
        Case 30058,58 : SetupRetVal = "NullReference"
        Case 30059,59 : SetupRetVal = "OCTPatchForbidden"
        Case 30060,60 : SetupRetVal = "OCTWrongMSIDll"
        Case 30061,61 : SetupRetVal = "OutOfBoundsIndex"
        Case 30062,62 : SetupRetVal = "OutOfDiskSpace"
        Case 30063,63 : SetupRetVal = "OutOfMemory"
        Case 30064,64 : SetupRetVal = "OutOfRange"
        Case 30065,65 : SetupRetVal = "PatchApplicationFailure"
        Case 30066,66 : SetupRetVal = "PreReqCheckFailure"
        Case 30067,67 : SetupRetVal = "ProcessAlreadyStarted"
        Case 30068,68 : SetupRetVal = "ProcessNotStarted"
        Case 30069,69 : SetupRetVal = "ProcessNotFinished"
        Case 30070,70 : SetupRetVal = "ProductAlreadyDefined"
        Case 30071,71 : SetupRetVal = "ResourceAlreadyTracked"
        Case 30072,72 : SetupRetVal = "ResourceNotFound"
        Case 30073,73 : SetupRetVal = "ResourceNotTracked"
        Case 30074,74 : SetupRetVal = "SQLAlreadyConnected"
        Case 30075,75 : SetupRetVal = "SQLFailedToAllocateHandle"
        Case 30076,76 : SetupRetVal = "SQLFailedToConnect"
        Case 30077,77 : SetupRetVal = "SQLFailedToExecuteStatement"
        Case 30078,78 : SetupRetVal = "SQLFailedToRetrieveData"
        Case 30079,79 : SetupRetVal = "SQLFailedToSetAttribute"
        Case 30080,80 : SetupRetVal = "StorageNotCreated"
        Case 30081,81 : SetupRetVal = "StreamNameTooLong"
        Case 30082,82 : SetupRetVal = "SystemError"
        Case 30083,83 : SetupRetVal = "ThreadAlreadyStarted"
        Case 30084,84 : SetupRetVal = "ThreadNotStarted"
        Case 30085,85 : SetupRetVal = "ThreadNotFinished"
        Case 30086,86 : SetupRetVal = "TooManyProducts"
        Case 30087,87 : SetupRetVal = "UnexpectedXMLNodeType"
        Case 30088,88 : SetupRetVal = "UnexpectedError"
        Case 30089,89 : SetupRetVal = "Unitialized"
        Case 30090,90 : SetupRetVal = "UserCancel"
        Case 30091,91 : SetupRetVal = "ExternalCommandFailed"
        Case 30092,92 : SetupRetVal = "SPDatabaseOverSize"
        Case 30093,93 : SetupRetVal = "IntegerTruncation"
        'msiexec return values
        Case 1259 : SetupRetVal = "APPHELP_BLOCK"
        Case 1601 : SetupRetVal = "INSTALL_SERVICE_FAILURE"
        Case 1602 : SetupRetVal = "INSTALL_USEREXIT"
        Case 1603 : SetupRetVal = "INSTALL_FAILURE"
        Case 1604 : SetupRetVal = "INSTALL_SUSPEND"
        Case 1605 : SetupRetVal = "UNKNOWN_PRODUCT"
        Case 1606 : SetupRetVal = "UNKNOWN_FEATURE"
        Case 1607 : SetupRetVal = "UNKNOWN_COMPONENT"
        Case 1608 : SetupRetVal = "UNKNOWN_PROPERTY"
        Case 1609 : SetupRetVal = "INVALID_HANDLE_STATE"
        Case 1610 : SetupRetVal = "BAD_CONFIGURATION"
        Case 1611 : SetupRetVal = "INDEX_ABSENT"
        Case 1612 : SetupRetVal = "INSTALL_SOURCE_ABSENT"
        Case 1613 : SetupRetVal = "INSTALL_PACKAGE_VERSION"
        Case 1614 : SetupRetVal = "PRODUCT_UNINSTALLED"
        Case 1615 : SetupRetVal = "BAD_QUERY_SYNTAX"
        Case 1616 : SetupRetVal = "INVALID_FIELD"
        Case 1618 : SetupRetVal = "INSTALL_ALREADY_RUNNING"
        Case 1619 : SetupRetVal = "INSTALL_PACKAGE_OPEN_FAILED"
        Case 1620 : SetupRetVal = "INSTALL_PACKAGE_INVALID"
        Case 1621 : SetupRetVal = "INSTALL_UI_FAILURE"
        Case 1622 : SetupRetVal = "INSTALL_LOG_FAILURE"
        Case 1623 : SetupRetVal = "INSTALL_LANGUAGE_UNSUPPORTED"
        Case 1624 : SetupRetVal = "INSTALL_TRANSFORM_FAILURE"
        Case 1625 : SetupRetVal = "INSTALL_PACKAGE_REJECTED"
        Case 1626 : SetupRetVal = "FUNCTION_NOT_CALLED"
        Case 1627 : SetupRetVal = "FUNCTION_FAILED"
        Case 1628 : SetupRetVal = "INVALID_TABLE"
        Case 1629 : SetupRetVal = "DATATYPE_MISMATCH"
        Case 1630 : SetupRetVal = "UNSUPPORTED_TYPE"
        Case 1631 : SetupRetVal = "CREATE_FAILED"
        Case 1632 : SetupRetVal = "INSTALL_TEMP_UNWRITABLE"
        Case 1633 : SetupRetVal = "INSTALL_PLATFORM_UNSUPPORTED"
        Case 1634 : SetupRetVal = "INSTALL_NOTUSED"
        Case 1635 : SetupRetVal = "PATCH_PACKAGE_OPEN_FAILED"
        Case 1636 : SetupRetVal = "PATCH_PACKAGE_INVALID"
        Case 1637 : SetupRetVal = "PATCH_PACKAGE_UNSUPPORTED"
        Case 1638 : SetupRetVal = "PRODUCT_VERSION"
        Case 1639 : SetupRetVal = "INVALID_COMMAND_LINE"
        Case 1640 : SetupRetVal = "INSTALL_REMOTE_DISALLOWED"
        Case 1641 : SetupRetVal = "SUCCESS_REBOOT_INITIATED"
        Case 1642 : SetupRetVal = "PATCH_TARGET_NOT_FOUND"
        Case 1643 : SetupRetVal = "PATCH_PACKAGE_REJECTED"
        Case 1644 : SetupRetVal = "INSTALL_TRANSFORM_REJECTED"
        Case 1645 : SetupRetVal = "INSTALL_REMOTE_PROHIBITED"
        Case 1646 : SetupRetVal = "PATCH_REMOVAL_UNSUPPORTED"
        Case 1647 : SetupRetVal = "UNKNOWN_PATCH"
        Case 1648 : SetupRetVal = "PATCH_NO_SEQUENCE"
        Case 1649 : SetupRetVal = "PATCH_REMOVAL_DISALLOWED"
        Case 1650 : SetupRetVal = "INVALID_PATCH_XML"
        Case 3010 : SetupRetVal = "SUCCESS_REBOOT_REQUIRED"
        Case Else : SetupRetVal = "Unknown Return Value"
    End Select
End Function 'SetupRetVal
'=======================================================================================================

Function GetProductID(sProdID)
        Dim sReturn
        
        Select Case sProdId

        Case "11" : sReturn = "PRO"
        Case "12" : sReturn = "STANDARD"
        Case "13" : sReturn = "BASIC"
        Case "14" : sReturn = "WSS2"
        Case "15" : sReturn = "Access"
        Case "16" : sReturn = "Excel"
        Case "17" : sReturn = "FrontPage"
        Case "18" : sReturn = "PowerPoint"
        Case "19" : sReturn = "Publisher"
        Case "1A" : sReturn = "Outlook"
        Case "1B" : sReturn = "Word"
        Case "1C" : sReturn = "AccessRuntime"
        Case "1E" : sReturn = "OfficeMUI"
        Case "1F" : sReturn = "PTK"
        Case "23" : sReturn = "OfficeMUI"
        Case "24" : sReturn = "ORK"
        Case "26" : sReturn = "XPWebComponents"
        Case "2E" : sReturn = "OSSSDK"
        Case "32" : sReturn = "PrjSrv"
        Case "33" : sReturn = "PERSONAL"
        Case "3A" : sReturn = "PrjStd" 
        Case "3B" : sReturn = "PrjPro"
        Case "3C" : sReturn = "PrjMUI"
        Case "44" : sReturn = "InfoPath"
        Case "48" : sReturn = "InfoPathVSToolkit"
        Case "49" : sReturn = "PIA"
        Case "51" : sReturn = "VisPro"
        Case "52" : sReturn = "VisView"
        Case "53" : sReturn = "VisStd"
        Case "55" : sReturn = "VisEA"
        Case "5E" : sReturn = "VisMUI"
        Case "83" : sReturn = "HtmlView"
        Case "84" : sReturn = "XLView"
        Case "85" : sReturn = "WDView"
        Case "92" : sReturn = "WSS2Pack"
        Case "93" : sReturn = "OWP&C"
        Case "A1" : sReturn = "OneNote"
        Case "A4" : sReturn = "OWC"
        Case "A5" : sReturn = "WSSMig"
        Case "A9" : sReturn = "InterConnect"
        Case "AA" : sReturn = "PPTCast"
        Case "AB" : sReturn = "PPTPack1"
        Case "AC" : sReturn = "PPTPack2"
        Case "AD" : sReturn = "PPTPack3"
        Case "AE" : sReturn = "OrgChart"
        Case "CA" : sReturn = "SmallBusiness"
        Case "D0" : sReturn = "AccessDE"
        Case "DC" : sReturn = "SmartDocSDK"
        Case "E0" : sReturn = "Outlook"
        Case "E3" : sReturn = "PROPLUS"
        Case "F7" : sReturn = "InfoPathVST"
        Case "F8" : sReturn = "RHDTool"
        Case "FD" : sReturn = "Outlook"
        Case "FF" : sReturn = "LIP"
        Case Else : sReturn = ProdId
        
        End Select 'ProdId

    GetProductID = sReturn
End Function 'GetProductID
'=======================================================================================================

Sub Log (sLog)
    wscript.echo sLog
    LogStream.WriteLine sLog
End Sub 'Log
'=======================================================================================================

Sub LogOnly (sLog)
    LogStream.WriteLine sLog
End Sub 'Log
'=======================================================================================================

Sub CheckError(sModule)
    If Err <> 0 Then 
        LogOnly Now & " - " & sModule & " - Source: " & Err.Source & "; Err# (Hex): " & Hex( Err ) & _
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
        Log "No argument specified. Preparing user prompt" & vbCrLf
        FindInstalledO11Products
        If dicInstalledProd.Count > 0 Then sDefault = Join(RemoveDuplicates(dicInstalledProd.Items),",") Else sDefault = "ALL"
        sDefault = InputBox("Enter a list of Office 2003 products to remove" & vbCrLf & vbCrLf & _
                "Examples:" & vbCrLf & _
                "ALL" & vbTab & vbTab & "-> remove all of Office 2003" & vbCrLf & _
                "ProPlus,PrjPro" & vbTab & "-> remove ProPlus and Project" & vbCrLf &_
                "?" & vbTab & vbTab & "-> display Help", _
                "OffScrub03 - Office 2003 remover", _
                sDefault)
        If IsEmpty(sDefault) Then 'User cancelled
            Log "User cancelled. CleanUp & Exit."
            wscript.quit 
        End If 'IsEmpty(sDefault)
        Log "Answer from prompt: " & sDefault & vbCrLf
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
        fRemoveAll = True
        fRemoveOSE = False
    
    Case "ALL,OSE","OSE,ALL"
        fRemoveAll = True
        fRemoveOSE = True
    
    Case Else
        fRemoveAll = False
        fRemoveOSE = False
        sSkuRemoveList = arrArguments(0)
    
    End Select
    
    For iCnt = 0 To UBound(arrArguments)

        Select Case arrArguments(iCnt)
        
        Case "?","/?","-?"
            ShowSyntax
        
        Case "/B","/BYPASS"
            If UBound(arrArguments)>iCnt Then
                If InStr(arrArguments(iCnt+1),"1")>0 Then fBypass_Stage1 = True
                If InStr(arrArguments(iCnt+1),"2")>0 Then fBypass_Stage2 = True
                If InStr(arrArguments(iCnt+1),"3")>0 Then fBypass_Stage3 = True
                If InStr(arrArguments(iCnt+1),"4")>0 Then fBypass_Stage4 = True
            End If
        
        Case "/D","/DELETEUSERSETTINGS"
            fKeepUser = False
        
        Case "/F","/FORCE"
            fForce = True
        
        Case "/L","/LOG"
            fLogInitialized = False
            If UBound(arrArguments)>iCnt Then
                If oFso.FolderExists(arrArguments(iCnt+1)) Then 
                    sLogDir = arrArguments(iCnt+1)
                Else
                    On Error Resume Next
                    oFso.CreateFolder(arrArguments(iCnt+1))
                    If Err <> 0 Then sLogDir = sScrubDir Else sLogDir = arrArguments(iCnt+1)
                End If
            End If
        
        Case "/O","/OSE"
            fRemoveOSE = True
        
        Case "/P","/PREVIEW"
            fDetectOnly = True
        
        Case "/Q","/QUIET"
            bQuiet = True
        
        Case "/S","/SKIPSD","/SKIPSHORTCUSTDETECTION"
            fSkipSD = True
        
        Case Else
        
        End Select
    Next 'iCnt
    If Not fLogInitialized Then CreateLog

End Sub 'ParseCmdLine
'=======================================================================================================

Sub CreateLog
    Dim DateTime
    Dim sLogName
    
    On Error Resume Next
    'Create the log file
    Set DateTime = CreateObject("WbemScripting.SWbemDateTime")
    DateTime.SetVarDate Now,True
    sLogName = sLogDir & "\" & oWShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
    sLogName = sLogName &  "_" & Left(DateTime.Value,14)
    sLogName = sLogName & "_ScrubLog.txt"
    Err.Clear
    Set LogStream = oFso.CreateTextFile(sLogName,True,True)
    If Err <> 0 Then 
        Err.Clear
        sLogDir = sScrubDir
        sLogName = sLogDir & "\" & oWShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
        sLogName = sLogName &  "_" & Left(DateTime.Value,14)
        sLogName = sLogName & "_ScrubLog.txt"
        Set LogStream = oFso.CreateTextFile(sLogName,True,True)
    End If

    Log "Microsoft Customer Support Services - Office 2003 Removal Utility" & vbCrLf & vbCrLf & _
                "Version: " & VERSION & vbCrLf & _
                "64 bit OS: " & f64 & vbCrLf & _
                "Start removal: " & Now & vbCrLf
    fLogInitialized = True
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
    Wscript.Echo sErr & vbCrLf & _
             "OffScrub03 V " & VERSION & vbCrLf & _
             "Copyright (c) Microsoft Corporation. All Rights Reserved" & vbCrLf & vbCrLf & _
             "OffScrub03 helps to remove Office 2003 when a regular uninstall is no longer possible" & vbCrLf & vbCrLf & _
             "Usage:" & vbTab & "OffScrub03.vbs [List of config ProductIDs] [Options]" & vbCrLf & vbCrLf & _
             vbTab & "/?                               ' Displays this help"& vbCrLf &_
             vbTab & "/DeleteUserSettings              ' Deletes some user profile contents & data"& vbCrLf &_
             vbTab & "/Force                           ' Enforces file removal. May cause data loss!" & vbCrLf &_
             vbTab & "/SkipShortcutDetection           ' Does not search the local hard drives for shortcuts" & vbCrLf & _
             vbTab & "/Log [LogfolderPath]             ' Custom folder for log files" & vbCrLf & _
             vbTab & "/OSE                             ' Forces removal of the Office Source Engine service" & vbCrLf &_
             vbTab & "/Quiet                           ' Setup.exe and Msiexec.exe run quiet with no UI" & vbCrLf &_
             vbTab & "/Preview                         ' Run this script to preview what would get removed"& vbCrLf & vbCrLf & _
             "Examples:"& vbCrLf & _
             vbTab & "OffScrub03.vbs ALL               ' Remove all Office 2003 products" & vbCrLf &_
             vbTab & "OffScrub03.vbs ProPlus,PrjPro    ' Remove ProPlus and Project" & vbCrLf
    Wscript.Quit
End Sub 'ShowSyntax
'=======================================================================================================