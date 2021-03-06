'=======================================================================================================
' Name: OffScrub10.vbs
' Author: Microsoft Customer Support Services
' Copyright (c) 2009,2010 Microsoft Corporation
' Script to remove (scrub) Office 2010 products
'=======================================================================================================
Option Explicit

Const SCRIPTVERSION = "1.26"
Const SCRIPTFILE    = "OffScrub10.vbs"
Const SCRIPTNAME    = "OffScrub10"
Const OVERSION      = "14.0"
Const OVERSIONMAJOR = "14"
Const OREF          = "Office14"
Const OREGREF       = "OFFICE14."
Const ONAME         = "Office 2010"
Const OPACKAGE      = "PackageRefs"
Const OFFICEID      = "0000000FF1CE}"
Const HKCR          = &H80000000
Const HKCU          = &H80000001
Const HKLM          = &H80000002
Const HKU           = &H80000003
Const FOR_WRITING   = 2
Const PRODLEN       = 13
Const COMPPERMANENT = "00000000000000000000000000000000"
Const UNCOMPRESSED  = 38
Const COMPRESSED    = 32
Const REG_ARP       = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"

'=======================================================================================================
Dim oFso, oMsi, oReg, oWShell, oWmiLocal
Dim ComputerItem, Item, LogStream, TmpKey
Dim arrTmpSKUs, arrDeleteFiles, arrDeleteFolders, arrMseFolders
Dim dicKeepProd, dicKeepLis, dicApps, dicKeepFolder, dicDelRegKey, dicKeepReg
Dim dicInstalledSku, dicRemoveSku, dicKeepSku, dicSrv, dicCSuite, dicCSingle
Dim f64, fLegacyProductFound
Dim sErr, sTmp, sSkuRemoveList, sDefault, sWinDir, sMode
Dim sAppData, sTemp, sScrubDir, sProgramFiles, sProgramFilesX86, sCommonProgramFiles, sCommonProgramFilesX86
Dim sAllusersProfile
Dim sProgramData, sLocalAppData, sOInstallRoot

'=======================================================================================================
'Main
'=======================================================================================================
'Configure defaults
Dim sLogDir : sLogDir = ""
Dim sMoveMessage: sMoveMessage = ""
Dim fRemoveOSE      : fRemoveOSE = False
Dim fRemoveAll      : fRemoveAll = False
Dim fRemoveC2R      : fRemoveC2R = False
Dim fRemoveAppV     : fRemoveAppV = False
Dim fRemoveCSuites  : fRemoveCSuites = False
Dim fRemoveCSingle  : fRemoveCSingle = False
Dim fRemoveSrv      : fRemoveSrv = False
Dim fKeepUser       : fKeepUser = True  'Default to keep per user settings
Dim fSkipSD         : fSkipSD = False 'Default to not Skip the Shortcut Detection
Dim fDetectOnly     : fDetectOnly = False
Dim fQuiet          : fQuiet = False
Dim fNoCancel       : fNoCancel = False
'CAUTION! -> "fForce" will kill running applications which can result in data loss! <- CAUTION
Dim fForce          : fForce = False
'CAUTION! -> "fForce" will kill running applications which can result in data loss! <- CAUTION
Dim fLogInitialized : fLogInitialized = False
Dim fBypass_Stage1  : fBypass_Stage1 = False 'Component Detection
Dim fBypass_Stage2  : fBypass_Stage2 = False 'Setup
Dim fBypass_Stage3  : fBypass_Stage3 = False 'Msiexec
Dim fBypass_Stage4  : fBypass_Stage4 = False 'CleanUp
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
sLocalAppData       = oWShell.ExpandEnvironmentStrings("%localappdata%")
sTemp               = oWShell.ExpandEnvironmentStrings("%temp%")
sAllUsersProfile    = oWShell.ExpandEnvironmentStrings("%allusersprofile%")
sProgramFiles       = oWShell.ExpandEnvironmentStrings("%programfiles%")
sCommonProgramFiles = oWShell.ExpandEnvironmentStrings("%commonprogramfiles%")
sProgramData        = oWSHell.ExpandEnvironmentStrings("%programdata%")
sWinDir             = oWShell.ExpandEnvironmentStrings("%windir%")
sScrubDir           = sTemp & "\" & SCRIPTNAME

'Create Dictionaries
Set dicKeepProd = CreateObject("Scripting.Dictionary")
Set dicInstalledSku = CreateObject("Scripting.Dictionary")
Set dicRemoveSku = CreateObject("Scripting.Dictionary")
Set dicKeepSku = CreateObject("Scripting.Dictionary")
Set dicKeepLis = CreateObject("Scripting.Dictionary")
Set dicKeepFolder = CreateObject("Scripting.Dictionary")
Set dicApps = CreateObject("Scripting.Dictionary")
Set dicDelRegKey = CreateObject("Scripting.Dictionary")
Set dicKeepReg = CreateObject("Scripting.Dictionary")
Set dicSrv = CreateObject("Scripting.Dictionary")
Set dicCSuite = CreateObject("Scripting.Dictionary")
Set dicCSingle = CreateObject("Scripting.Dictionary")
dicApps.Add "communicator.exe","communicator.exe"
dicApps.Add "bcssync.exe","bcssync.exe"
dicApps.Add "officesas.exe","officesas.exe"
dicApps.Add "officesasscheduler.exe","officesasscheduler.exe"
dicApps.Add "msosync.exe","msosync.exe"
dicApps.Add "onenotem.exe","onenotem.exe"

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
If f64 Then sCommonProgramFilesX86 = oWShell.ExpandEnvironmentStrings("%CommonProgramFiles(x86)%")

'Call the command line parser
ParseCmdLine

If Not CheckRegPermissions Then
    Log "Insufficient registry access permissions - exiting"
    'Undo temporary entries created in ARP
    TmpKeyCleanUp
    wscript.quit 
End If

'Get Office Install Folder
If NOT RegReadValue(HKLM,"SOFTWARE\Microsoft\Office\"&OVERSION&"\Common\InstallRoot","Path",sOInstallRoot,"REG_SZ") Then 
    sOInstallRoot = sProgramFiles & "\Microsoft Office\"&OREF
End If

'Ensure integrity of WI metadata which could fail used APIs otherwise
EnsureValidWIMetadata HKCU,"Software\Classes\Installer\Products",COMPRESSED
EnsureValidWIMetadata HKCR,"Installer\Products",COMPRESSED
EnsureValidWIMetadata HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products",COMPRESSED
EnsureValidWIMetadata HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components",COMPRESSED
EnsureValidWIMetadata HKCR,"Installer\Components",COMPRESSED

'-------------------
'Stage # 0 - Basics |
'-------------------
'Build a list with installed/registered Office products
sTmp = "Stage # 0 " & chr(34) & "Basics" & chr(34) & " (" & Time & ")"
Log vbCrLf & sTmp & vbCrLf & String(Len(sTmp),"=") & vbCrLf

FindInstalledOProducts
If dicInstalledSku.Count > 0 Then Log "Found registered product(s): " & Join(RemoveDuplicates(dicInstalledSku.Items),",") &vbCrLf

'Validate the list of products we got from the command line if applicable
ValidateRemoveSkuList

'Log detection results
If dicRemoveSku.Count > 0 Then Log "Product(s) to be removed: " & Join(RemoveDuplicates(dicRemoveSku.Items),",")
sMode = "Selected " & ONAME & " products"
If Not dicRemoveSku.Count > 0 Then sMode = "Orphaned " & ONAME & " products"
If fRemoveAll Then sMode = "All " & ONAME & " products"
Log "Final removal mode: " & sMode
Log "Remove OSE service: " & fRemoveOSE &vbCrLf

'Log preview mode if applicable
If fDetectOnly Then Log "*************************************************************************"
If fDetectOnly Then Log "*                          PREVIEW MODE                                 *"
If fDetectOnly Then Log "* All uninstall and delete operations will only be logged not executed! *"
If fDetectOnly Then Log "*************************************************************************" & vbCrLf

'Check if there are legacy products installed
CheckForLegacyProducts
If fLegacyProductFound Then Log "Found legacy Office products that will not be removed." Else Log "No legacy Office products found."

'Cache .msi files
If dicRemoveSku.Count > 0 Then CacheMsiFiles

'--------------------------------
'Stage # 1 - Component Detection |
'--------------------------------
sTmp = "Stage # 1 " & chr(34) & "Component Detection" & chr(34) & " (" & Time & ")"
Log vbCrLf & sTmp & vbCrLf & String(Len(sTmp),"=") & vbCrLf
If Not fBypass_Stage1 Then
    'Build a list with files which are installed/registered to a product that's going to be removed
    Log "Prepare for CleanUp stages."
    Log "Identifying removable elements. This can take several minutes."
    ScanComponents 
Else
    Log "Skipping Component Detection because bypass was requested."
End If

'End all running Office applications
If fForce OR fQuiet Then CloseOfficeApps

'----------------------
'Stage # 2 - Setup.exe |
'----------------------
sTmp = "Stage # 2 " & chr(34) & "Setup.exe" & chr(34) & " (" & Time & ")"
Log vbCrLf & sTmp & vbCrLf & String(Len(sTmp),"=") & vbCrLf
If Not fBypass_Stage2 Then
    SetupExeRemoval
Else
    Log "Skipping Setup.exe because bypass was requested."
End If

'------------------------
'Stage # 3 - Msiexec.exe |
'------------------------
sTmp = "Stage # 3 " & chr(34) & "Msiexec.exe" & chr(34) & " (" & Time & ")"
Log vbCrLf & sTmp & vbCrLf & String(Len(sTmp),"=") & vbCrLf
If Not fBypass_Stage3 Then
    MsiexecRemoval
Else
    Log "Skipping Msiexec.exe because bypass was requested."
End If

'--------------------
'Stage # 4 - CleanUp |
'--------------------
'Removal of files and registry settings
sTmp = "Stage # 4 " & chr(34) & "CleanUp" & chr(34) & " (" & Time & ")"
Log vbCrLf & sTmp & vbCrLf & String(Len(sTmp),"=") & vbCrLf
If Not fBypass_Stage4 Then
    
    'Office Source Engine
    If fRemoveOSE Then RemoveOSE

    'Softgrid Service
    If fRemoveAppV Then RemoveSG

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
    
    'Temporary files
    DelScrubTmp
    
Else
    Log "Skipping CleanUp because bypass was requested."
End If

If Not sMoveMessage = "" Then Log vbCrLf & "Please remove this folder after next reboot: " & sMoveMessage

'THE END
Log vbCrLf & "End removal: " & Now & vbCrLf
Log vbCrLf & "For detailed logging please refer to the log in folder " &chr(34)&sScrubDir&chr(34)&vbCrLf

If fRebootRequired Then Log vbCrLf & "A restart is required to complete the operation!"
'=======================================================================================================
'=======================================================================================================

'Stage 0 - 4 Subroutines
'=======================================================================================================

'Office 2010 configuration products are listed with their configuration product name in the "Uninstall" key
'To identify an Office 2010 configuration product all of these condiditions have to be met:
' - "SystemComponent" does not have a value of "1" (DWORD) 
' - "OPACKAGE" entry exists and is not empty
' - "DisplayVersion" exists and the 2 leftmost digits are "OVERSIONMAJOR"
Sub FindInstalledOProducts
    Dim ArpItem
    Dim hDefKey, sCurKey, sName, sValue, sConfigName, sLcid, sProdC, sCVHValue
    Dim sProductCodeList, sProductCode 
    Dim arrKeys, arrValues, arrMultiSzValues, arrProdC
    Dim fSystemComponent0, fPackages, fDisplayVersion, fReturn, fCategorized

    If dicInstalledSku.Count > 0 Then Exit Sub 'Already done from InputBox prompt
    
    'Locate standalone Office 2010 products that have no configuration product entry and create a
    'temporary configuration entry
    ReDim arrTmpSKUs(-1)
    If RegEnumKey(HKLM,REG_ARP,arrKeys) Then
        For Each ArpItem in arrKeys
            If Len(ArpItem) = 38 Then
                If UCase(Right(ArpItem,PRODLEN))=OFFICEID AND Mid(ArpItem,4,2)=OVERSIONMAJOR Then
                    sCurKey = REG_ARP & ArpItem & "\"
                    fSystemComponent0 = Not (RegReadValue(HKLM,sCurKey,"SystemComponent",sValue,"REG_DWORD") AND (sValue = "1"))
                    If (fSystemComponent0 AND (NOT RegReadValue(HKLM,sCurKey,"CVH",sCVHValue,"REG_DWORD"))) Then
                        RegReadValue HKLM,sCurKey,"DisplayVersion",sValue,"REG_SZ"
                        Redim arrMultiSzValues(0)
                        'Logic changed to drop the LCID identifier
                        'sConfigName = GetProductID(Mid(ArpItem,11,4)) & "_" & CInt("&h" & Mid(ArpItem,16,4))
                        sConfigName = OREGREF & GetProductID(Mid(ArpItem,11,4))
                        If NOT RegKeyExists(HKLM,REG_ARP&sConfigName) Then
                            'Create a new ARP item
                            ReDim Preserve arrTmpSKUs(UBound(arrTmpSKUs)+1)
                            arrTmpSKUs(UBound(arrTmpSKUs)) = sConfigName
                            oReg.CreateKey HKLM,REG_ARP & sConfigName
                            arrMultiSzValues(0) = sConfigName
                            oReg.SetMultiStringValue HKLM,REG_ARP & sConfigName,OPACKAGE,arrMultiSzValues
                            arrMultiSzValues(0) = ArpItem
                            oReg.SetMultiStringValue HKLM,REG_ARP & sConfigName,"ProductCodes",arrMultiSzValues
                            oReg.SetStringValue HKLM,REG_ARP & sConfigName,"DisplayVersion",sValue
                            oReg.SetDWordValue HKLM,REG_ARP & sConfigName,"SystemComponent",0
                        Else
                            'Update the existing temporary ARP item
                            fReturn = RegReadValue(HKLM,REG_ARP&sConfigName,"ProductCodes",sProdC,"REG_MULTI_SZ")
                            If NOT InStr(sProdC,ArpItem)>0 Then sProdC = sProdC & chr(34) & ArpItem
                            oReg.SetMultiStringValue HKLM,REG_ARP & sConfigName,"ProductCodes",Split(sProdC,chr(34))
                        End If 'RegKeyExists
                    End If 'fSystemComponent0
                End If 'OFFICEID
            End If 'Len 38
        Next 'ArpItem
    End If 'RegEnumKey
    
    'Find the configuration products
    If RegEnumKey(HKLM,REG_ARP,arrKeys) Then
        For Each ArpItem in arrKeys
            sCurKey = REG_ARP & ArpItem & "\"
            sValue = ""
            fSystemComponent0 = NOT (RegReadValue(HKLM,sCurKey,"SystemComponent",sValue,"REG_DWORD") AND (sValue = "1"))
            fPackages = RegReadValue(HKLM,sCurKey,OPACKAGE,sValue,"REG_MULTI_SZ")
            fDisplayVersion = RegReadValue(HKLM,sCurKey,"DisplayVersion",sValue,"REG_SZ")
            If fDisplayVersion Then
                If Len(sValue) > 1 Then
                    fDisplayVersion = (Left(sValue,2) = OVERSIONMAJOR)
                Else
                    fDisplayVersion = False
                End If
            End If
            If (fSystemComponent0 AND fPackages AND fDisplayVersion) OR (fSystemComponent0 AND fDisplayVersion AND InStr(UCase(ArpItem),"CLICK2RUN")>0) Then
                If InStr(ArpItem,".")>0 Then sConfigName = UCase(Mid(ArpItem,InStr(ArpItem,".")+1)) Else sConfigName = UCase(ArpItem)
                If NOT dicInstalledSku.Exists(sConfigName) Then dicInstalledSku.Add sConfigName,sConfigName

                'Categorize the SKU
                'Three categories are available: ClientSuite, ClientSingleProduct, Server
                If RegReadValue(HKLM,REG_ARP&OREGREF&sConfigName,"ProductCodes",sProductCodeList,"REG_MULTI_SZ") OR (sConfigName = "CLICK2RUN") Then
                    fCategorized = False
                    If sConfigName = "CLICK2RUN" Then sProductCodeList = "{90140011-0062-0000-0000-0000000FF1CE}"
                    For Each sProductCode in Split(sProductCodeList,chr(34))
                        If Len(sProductCode) = 38 Then
                            If NOT Mid(sProductCode,11,1) = "0" Then
                                'Server product
                                If NOT dicSrv.Exists(UCase(sConfigName)) Then dicSrv.Add UCase(sConfigName),sConfigName
                                fCategorized = True
                                Exit For
                            Else
                                Select Case Mid(sProductCode,11,4)
                                'Client Suites
                                Case "000F","0011","0012","0013","0014","0015","0016","0017","0018","0019","001A","001B","0029","002B","002E","002F","0030","0031","0033","0035","0037","003D","0044","0049","0061","0062","0066","006C","006D","006F","0074","00A1","00A3","00A9","00BA","00CA","00E0","0100","0103","011A"
                                    If NOT dicCSuite.Exists(UCase(sConfigName)) Then dicCSuite.Add UCase(sConfigName),sConfigName
                                    fCategorized = True
                                    Exit For

                                Case Else
                                End Select
                            End If

                        End If 'Len 38
                    Next 'sProductCode
                    If NOT fCategorized Then
                        If NOT dicCSingle.Exists(UCase(sConfigName)) Then dicCSingle.Add UCase(sConfigName),sConfigName
                    End If 'fCategorized
                End If 'RegReadValue "ProductCodes"

            End If
        Next 'ArpItem
    End If 'RegEnumKey
End Sub 'FindInstalledOProducts
'=======================================================================================================

'Check if there are Office products from previous versions on the computer
Sub CheckForLegacyProducts
    Const OLEGACY = "78E1-11D2-B60F-006097C998E7}.6000-11D3-8CFE-0050048383C9}.6000-11D3-8CFE-0150048383C9}.BDCA-11D1-B7AE-00C04FB92F3D}.6D54-11D4-BEE3-00C04F990354}"
    Dim Product
    
    'Set safe default
    fLegacyProductFound = True
    
    For Each Product in oMsi.Products
        'Handle O09 - O11 Products
        If InStr(OLEGACY, UCase(Right(Product, 28)))>0 Then
            'Found legacy Office product. Keep flag in default and exit
            Exit Sub
        End If
        If UCase(Right(Product,PRODLEN))=OFFICEID AND Mid(Product,4,2)="12" Then
            'Found Office 2007 Product. Keep flag in default and exit
            Exit Sub
        End If
    Next 'Product
    fLegacyProductFound = False
    
End Sub 'CheckForLegacyProducts
'=======================================================================================================

'Create clean list of Products to remove.
'Strip off bad & empty contents
Sub ValidateRemoveSkuList
    Dim Sku, Key, sProductCode, sProductCodeList
    Dim arrRemoveSKUs
    
    If fRemoveAll Then
        'Remove all mode
        For Each Key in dicInstalledSku.Keys
            dicRemoveSku.Add Key,dicInstalledSku.Item(Key)
        Next 'Key
    Else
        'Remove individual products or preconfigured configurations mode
        
        'Ensure to have a string with no unexpected contents
        sSkuRemoveList = Replace(sSkuRemoveList,";",",")
        sSkuRemoveList = Replace(sSkuRemoveList," ","")
        sSkuRemoveList = Replace(sSkuRemoveList,Chr(34),"")
        While InStr(sSkuRemoveList,",,")>0
            sSkuRemoveList = Replace(sSkuRemoveList,",,",",")
        Wend
        
        'Prepare 'remove' and 'keep' dictionaries to determine what has to be removed
        
        'Initial pre-fill of 'keep' dic
        For Each Key in dicInstalledSku.Keys
            dicKeepSku.Add Key,dicInstalledSku.Item(Key)
        Next 'Key
        
        'Determine contents of keep and remove dic
        'Individual products
        arrRemoveSKUs = Split(UCase(sSkuRemoveList),",")
        For Each Sku in arrRemoveSKUs
            If Sku = "OSE" Then fRemoveOSE = True
            If Sku = "CLICK2RUN" Then fRemoveC2R = True
            If dicKeepSku.Exists(Sku) Then
                'A Sku to remove has been passed in
                'remove the item from the keep dic
                dicKeepSku.Remove(Sku)
                'Now add it to the remove dic
                If NOT dicRemoveSku.Exists(Sku) Then dicRemoveSku.Add Sku,Sku
            End If
        Next 'Sku

        'Client Suite Category
        If fRemoveCSuites Then
            fRemoveC2R = True
            For Each Key in dicInstalledSku.Keys
                If dicCSuite.Exists(Key) Then
                    If dicKeepSku.Exists(Key) Then dicKeepSku.Remove(Key)
                    If NOT dicRemoveSku.Exists(Key) Then dicRemoveSku.Add Key,Key
                End If
            Next 'Key
        End If 'fRemoveCSuites
        
        'Client Single/Standalone Category
        If fRemoveCSingle Then
            For Each Key in dicInstalledSku.Keys
                If dicCSingle.Exists(Key) Then
                    If dicKeepSku.Exists(Key) Then dicKeepSku.Remove(Key)
                    If NOT dicRemoveSku.Exists(Key) Then dicRemoveSku.Add Key,Key
                End If
            Next 'Key
        End If 'fRemoveCSingle
        
        'Server Category
        If fRemoveSrv Then
            For Each Key in dicInstalledSku.Keys
                If dicSrv.Exists(Key) Then
                    If dicKeepSku.Exists(Key) Then dicKeepSku.Remove(Key)
                    If NOT dicRemoveSku.Exists(Key) Then dicRemoveSku.Add Key,Key
                End If
            Next 'Key
        End If 'fRemoveSrv
        
        If NOT dicKeepSku.Count > 0 Then fRemoveAll = True

    End If 'fRemoveAll

    'Fill the KeepProd dic
    For Each Sku in dicKeepSku.Keys
        If RegReadValue(HKLM,REG_ARP & "Office14."&Sku,"ProductCodes",sProductCodeList,"REG_MULTI_SZ") Then
            For Each sProductCode in Split(sProductCodeList,chr(34))
                If NOT dicKeepProd.Exists(sProductCode) Then dicKeepProd.Add sProductCode,Sku
            Next 'sProductCod 
        End If
    Next 'Sku
        
    If fRemoveAll OR fRemoveOSE Then CheckRemoveOSE
    If fRemoveAll OR fRemoveC2R Then CheckRemoveSG

End Sub 'ValidateRemoveSkuList
'=======================================================================================================

'Check if SoftGrid Client can be scrubbed
Sub CheckRemoveSG

    Dim Key
    Dim sPKey
    Dim arrKeys
    'Wbem is not available on all OS
    'On Error Resume Next
    If fForce Then
        fRemoveAppV = True
        Exit Sub
    End If
    
    fRemoveAppV = False
    If RegEnumKey (HKLM,"SOFTWARE\Microsoft\SoftGrid\4.5\Client\Applications",arrKeys) Then
        For Each Key in arrKeys
            If Len(Key)>15 Then
                'Get Partial product Key
                sPKey = Right(Key,16)
                If Left(sPKey,4) = "9014" Then
                    If NOT GetProductID(Mid(sPKey,5,4)) = "CLICK2RUN" Then Exit Sub
                Else
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
        Next 'Key
    End If
    'If we got here it's only Click2Run apps
    fRemoveAppV = True

End Sub 'CheckRemoveSG
'=======================================================================================================

'Check if OSE service can be scrubbed
Sub CheckRemoveOSE
    Const O11 = "6000-11D3-8CFE-0150048383C9}"
    Dim Product
    
    fRemoveOSE = False
    For Each Product in oMsi.Products
        If UCase(Right(Product,28)) = O11 Then 
            'Found Office 2003 Product. Set flag to not remove the OSE service
            Exit Sub
        End If
        If UCase(Right(Product,PRODLEN))=OFFICEID AND Mid(Product,4,2)="12" Then
            'Found Office 2007 Product. Set flag to not remove the OSE service
            fRemoveOSE = False
            Exit Sub
        End If
    Next 'Product
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
    For Each Product in oMsi.Products
        'Ensure valid GUID length
        If Len(Product) = 38 Then
            If (Right(Product,PRODLEN) = OFFICEID AND Mid(Product,4,2)="14") AND (fRemoveAll OR CheckDelete(Product))Then
                CheckError "CacheMsiFiles"
                sMsiFile = oMsi.ProductInfo(Product,"LocalPackage") : CheckError "CacheMsiFiles"
                LogOnly " - " & Product & ".msi"
                If oFso.FileExists(sMsiFile) Then oFso.CopyFile sMsiFile,sScrubDir & "\" & Product & ".msi",True
                CheckError "CacheMsiFiles"
            End If 'Right(Product,PRODLEN) = OFFICEID ...
        End If '38
    Next 'Product
    Err.Clear
End Sub 'CacheMsiFiles
'=======================================================================================================

'Build a list of all files that will be deleted
Sub ScanComponents
    Const MSIOPENDATABASEREADONLY = 0
    Const MSIINSTALLSTATE_LOCAL = 3

    Dim FileList, RegList, ComponentID, CompClient, Record, qView, MsiDb
    Dim Processes, Process, Prop, prod
    Dim sQuery, sSubKeyName, sPath, sFile, sMsiFile, sCompClient, sComponent
    Dim fRemoveComponent, fAffectedComponent, fIsPermanent
    Dim i, iProgress, iCompCnt, iRemCnt
    Dim dicFLError, oDic, oFolderDic, dicCompPath
    Dim hDefKey

    'Logfile
    Set FileList = oFso.OpenTextFile(sScrubDir & "\FileList.txt",FOR_WRITING,True,True)
    Set RegList = oFso.OpenTextFile(sScrubDir & "\RegList.txt",FOR_WRITING,True,True)
    
    'FileListError dic
    Set dicFLError = CreateObject("Scripting.Dictionary")
    
    Set oDic = CreateObject("Scripting.Dictionary")
    Set oFolderDic = CreateObject("Scripting.Dictionary")
    Set dicCompPath = CreateObject("Scripting.Dictionary")

    'Prevent that API errors fail script execution
    On Error Resume Next

    iCompCnt = oMsi.Components.Count
    If NOT Err = 0 Then
        'API failure
        Log "Error during components detection. Cannot complete this task."
        Err.Clear
        Exit Sub
    End If

    'Ensure to not divide by zero
    If iCompCnt = 0 Then iCompCnt = 1
    LogOnly " Scanning " & iCompCnt & " components"
    'Enum all Components
    For Each ComponentID In oMsi.Components
        'Progress bar
        i = i + 1
        If iProgress < (i / iCompCnt) * 100 Then 
            wscript.stdout.write "." : LogStream.Write "."
            iProgress = iProgress + 1
            If iProgress = 35 OR iProgress = 70 Then Log ""
        End If

        'Check if all ComponentClients will be removed
        sCompClient = ""
        iRemCnt = 0
        fIsPermanent = False
        fRemoveComponent = False 'Flag to track if the component will be completely removed
        fAffectedComponent = False 'Flag to track if some clients remain installed who have a none shared location
        dicCompPath.RemoveAll
        For Each CompClient In oMsi.ComponentClients(ComponentID)
            If Err = 0 Then
                'Ensure valid guid length
                If Len(CompClient) = 38 Then
                    sPath = ""
                    sPath = LCase(oMsi.ComponentPath(CompClient,ComponentID))
                    sPath = Replace(sPath,"?",":")
                    'Scan for msidbComponentAttributesPermanent flag
                    If CompClient = "{00000000-0000-0000-0000-000000000000}" Then
                        fIsPermanent = True
                        iRemCnt = iRemCnt + 1
                    End If
                    fRemoveComponent = Right(CompClient,PRODLEN)=OFFICEID AND Mid(CompClient,4,2)="14"
                    If fRemoveComponent Then fRemoveComponent = CheckDelete(CompClient)
                    If fRemoveComponent Then
                        iRemCnt = iRemCnt + 1
                        fAffectedComponent = True
                        'Since the scope remains within one Office family the keypath for the component
                        'is assumed to be identical
                        If sCompClient = "" Then sCompClient = CompClient
                    Else
                        If NOT dicCompPath.Exists(sPath) Then dicCompPath.Add sPath,CompClient
                    End If
                Else
                    If NOT dicFLError.Exists("Error: Invalid metadata found. ComponentID: "&ComponentID &", ComponentClient: "&CompClient) Then _
                        dicFLError.Add "Error: Invalid metadata found. ComponentID: "&ComponentID &", ComponentClient: "&CompClient, ComponentID
                End If '38
            Else
                Err.Clear
            End If 'Err = 0
        Next 'CompClient
        
        'Determine if the component resources go away
        sPath = ""
        fRemoveComponent = fAffectedComponent AND (iRemCnt = oMsi.ComponentClients(ComponentID).Count)
        If NOT fRemoveComponent AND fAffectedComponent Then
            'Flag as removable if component has a unique keypath
            sPath = LCase(oMsi.ComponentPath(sCompClient,ComponentID))
            sPath = Replace(sPath,"?",":")
            fRemoveComponent = NOT dicCompPath.Exists(sPath)
        End If
        If fRemoveComponent Then
            'Check msidbComponentAttributesPermanent flag
            If fIsPermanent AND NOT fForce Then fRemoveComponent = False
        End If

        If fRemoveComponent Then
            'Component resources go away for this product
            Err.Clear
            'Get the component path
            If sPath = "" Then
                sPath = LCase(oMsi.ComponentPath(sCompClient,ComponentID))
                sPath = Replace(sPath,"?",":")
            End If
            If Len(sPath) > 4 Then
                If Left(sPath,1) = "0" Then
                    'Registry keypath

                    Select Case Left(sPath,2)
                    Case "00"
                        sPath = Mid(sPath,5)
                        hDefKey = HKCR
                    Case "01"
                        sPath = Mid(sPath,5)
                        hDefKey = HKCU
                    Case "02"
                        sPath = Mid(sPath,5)
                        hDefKey = HKLM
                    Case Else
                        '
                    End Select
                    If NOT dicDelRegKey.Exists(sPath) Then
                        dicDelRegKey.Add sPath,hDefKey
                        RegList.WriteLine HiveString(hDefKey)&"\"&sPath
                    End If
                Else
                
                    'File
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
                                'sFile = sPath & "\" & sFile
                                If Not oDic.Exists(sPath & "\" & sFile) Then 
                                    'Exception handler
                                    fAdd = True
                                    Select Case UCase(sFile)
                                    Case "FPERSON.DLL"
                                        For Each prod in oMsi.Products
                                            If NOT Checkdelete(prod) Then
                                                If oMsi.FeatureState(prod, "MSTagPluginNamesFiles") = MSIINSTALLSTATE_LOCAL Then
                                                    fAdd = False
                                                    Exit For
                                                End If
                                            End If
                                        Next 'prod
                                    Case Else
                                    End Select
                                    If fAdd Then
                                        oDic.Add sPath & "\" & sFile,sFile
                                        FileList.WriteLine sFile
                                        If Len(sFile)>4 Then
                                            sFile = LCase(sFile)
                                            If Right(sFile,4) = ".exe" Then
                                                If NOT dicApps.Exists(sFile) Then
                                                    Select Case sFile
                                                    Case "setup.exe","ose.exe","osppsvc.exe","explorer.exe","cvhsvc.exe","sftvsa.exe","sftlist.exe","sftplay.exe","sftvol.exe","sftfs.exe"
                                                    Case Else
                                                        dicApps.Add sFile,LCase(sPath) & "\" & sFile
                                                    End Select
                                                End If 'dicApps.Exists
                                            End If '.exe
                                        End If 'Len > 4
                                    End If 'fAdd
                                End If 'oDic.Exists
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
                End If
            End If 'Len(sPath) > 4
        Else
            'Add the path to the 'Keep' dictionary
            Err.Clear
            For Each CompClient In oMsi.ComponentClients(ComponentID)
                'Get the component path
                sPath = "" : sPath = LCase(oMsi.ComponentPath(CompClient,ComponentID))
                sPath = Replace(sPath,"?",":")
                
                If Len(sPath) > 4 Then
                    If Left(sPath,1) = "0" Then
                        'Registry keypath

                        Select Case Left(sPath,2)
                        Case "00"
                            sPath = Mid(sPath,5)
                            hDefKey = HKCR
                        Case "01"
                            sPath = Mid(sPath,5)
                            hDefKey = HKCU
                        Case "02"
                            sPath = Mid(sPath,5)
                            hDefKey = HKLM
                        Case Else
                            '
                        End Select
                        If NOT dicKeepReg.Exists(sPath) Then
                            dicKeepReg.Add sPath,hDefKey
                        End If
                    Else
                        'File keypath
                        If oFso.FileExists(sPath) Then
                            sPath = LCase(oFso.GetFile(sPath).ParentFolder)
                            If Not dicKeepFolder.Exists(sPath) Then AddKeepFolder sPath
                        End If
                    End If 'Is Registry
                End If 'sPath > 4
            Next 'CompClient
        End If 'fRemoveComponent
    Next 'ComponentID
    Err.Clear
    On Error Goto 0
    
    'Click2Run detection
    If C2RInstalled Then
        'Add executables that might need to be closed
        If NOT dicApps.Exists("cvh.exe") Then dicApps.Add "cvh.exe","cvh.exe"
        If NOT dicApps.Exists("officevirt.exe") Then dicApps.Add "officevirt.exe","officevirt.exe"

        Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process")
        For Each Process in Processes
            For Each Prop in Process.Properties_
                If Prop.Name = "ExecutablePath" Then 
                    If Len(Prop.Value) > 2 Then
                        If UCase(Left(Prop.Value,2)) = "Q:" Then
                            If NOT dicApps.Exists(LCase(Process.Name)) Then dicApps.Add LCase(Process.Name),Process.Name
                        End If 'Q:
                    End If '>2
                End If 'ExcecutablePath
            Next 'Prop
        Next 'Process

    End If 'C2RInstalled

    Log " Done"
    If dicFLError.Count > 0 Then LogOnly Join(dicFLError.Keys,vbCrLf)
    If Not oFolderDic.Count = 0 Then arrDeleteFolders = oFolderDic.Keys Else Set arrDeleteFolders = Nothing
    If Not oDic.Count = 0 Then arrDeleteFiles = oDic.Keys Else Set arrDeleteFiles = Nothing
End Sub 'ScanComponents
'=======================================================================================================


'Detect if Click2Run products are installed on the client
Function C2RInstalled

Dim Key, sPKey, sValue, VProd
Dim arrKeys

    If RegEnumKey (HKLM,REG_ARP,arrKeys) Then
        For Each Key in arrKeys
            If Len(Key)=38 Then
                If UCase(Right(Key,PRODLEN))=OFFICEID AND Mid(Key,4,2)="14" Then
                    If RegReadValue(HKLM,REG_ARP&"\"&Key,"CVH",sValue,"REG_DWORD") Then
                        If sValue = "1" Then
                            C2RInstalled = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next 'Key
    End If

    If RegEnumKey (HKLM,"SOFTWARE\Microsoft\SoftGrid\4.5\Client\Applications",arrKeys) Then
        For Each Key in arrKeys
            If Len(Key)>15 Then
                'Get Partial product Key
                sPKey = Right(Key,16)
                If Left(sPKey,4) = "9014" Then
                    If GetProductID(Mid(sPKey,5,4)) = "CLICK2RUN" Then
                        C2RInstalled = True
                        Exit Function
                    End If
                End If
            End If
        Next 'Key
    End If

End Function 'C2RInstalled
'=======================================================================================================

'Try to remove the products by calling setup.exe
Sub SetupExeRemoval
    Dim OseService, Service, TextStream
    Dim iSetupCnt, RetVal
    Dim Sku, sConfigFile, sUninstallCmd, sCatalyst, sCVHBS, sDll, sDisplayLevel, sNoCancel

    iSetupCnt = 0
    If Not dicRemoveSku.Count > 0 Then
        Log " Nothing to remove for Setup.exe"
        Exit Sub
    End If
    
    For Each Sku in dicRemoveSku.Keys
        If Sku="CLICK2RUN" Then
            
            'Reset Softgrid
            ResetSG 
            
            If f64 Then 
                sCVHBS = sCommonProgramFilesX86 & "\Microsoft Shared\Virtualization Handler\CVHBS.exe"
            Else
                sCVHBS = sCommonProgramFiles & "\Microsoft Shared\Virtualization Handler\CVHBS.exe"
            End If
            If oFso.FileExists(sCVHBS) Then
                If fQuiet Then CvhbsDialogHandler
                sUninstallCmd = Chr(34) & sCVHBS & Chr(34) & " /removesilent"
                iSetupCnt = iSetupCnt + 1
                Log " - Calling CVHBS.exe to remove " & Sku  
                If Not fDetectOnly Then
                    On Error Resume Next
                    RetVal = oWShell.Run(sUninstallCmd,0,True) : CheckError "CVHBSRemoval"
                    fRebootRequired = True
                    Log " - CVHBS.exe returned: " & SetupExeRetVal(Retval) & " (" & RetVal & ")" & vbCrLf
                    On Error Goto 0
                Else
                    Log " -> Removal suppressed in preview mode."
                End If
            Else
                Log "Error: Office Click-to-Run 2010 CVHBS.exe appears to be missing"
            End If 'oFso.FileExists

            'Make sure that C2R keys are gone to unblock the msiexec task

        End If 'Sku = Click2run
    Next 'Sku

    'Ensure that the OSE service is *installed, *not disabled, *running under System context.
    'If validation fails exit out of this sub.
    Set OseService = oWmiLocal.Execquery("Select * From Win32_Service Where Name like 'ose%'")
    If OseService.Count = 0 Then Exit Sub
    For Each Service in OseService
        If (Service.StartMode = "Disabled") AND (Not Service.ChangeStartMode("Manual")=0) Then Exit Sub
        If (Not Service.StartName = "LocalSystem") AND (Service.Change( , , , , , , "LocalSystem", "")) Then Exit Sub
    Next 'Service
    
    For Each Sku in dicRemoveSku.Keys
        If Sku="CLICK2RUN" Then
            'Already done
        Else
            'Create an "unattended" config.xml file for uninstall
            If fQuiet Then sDisplayLevel = "None" Else sDisplayLevel="Basic"
            If fNoCancel Then sNoCancel="Yes" Else sNoCancel="No"
            Set TextStream = oFso.OpenTextFile(sScrubDir & "\config.xml",FOR_WRITING,True,True)
            TextStream.Writeline "<Configuration Product=""" & Sku & """>"
            TextStream.Writeline "<Display Level=""" & sDisplayLevel & """ CompletionNotice=""No"" SuppressModal=""Yes"" NoCancel=""" & sNoCancel & """ AcceptEula=""Yes"" />"
            TextStream.Writeline "<Logging Type=""Verbose"" Path=""" & sLogDir & """ Template=""Microsoft Office " & Sku & " Setup(*).txt"" />"
            TextStream.Writeline "<Setting Id=""SETUP_REBOOT"" Value=""Never"" />"
            TextStream.Writeline "</Configuration>"
            TextStream.Close
            Set TextStream = Nothing
        
            'Ensure path to setup.exe is valid to prevent errors
            sDll = ""
            If RegReadValue(HKLM,REG_ARP & "Office14."&Sku,"UninstallString",sCatalyst,"REG_SZ") Then
                If InStr(LCase(sCatalyst),"/dll")>0 Then sDll = Right(sCatalyst,Len(sCatalyst)-InStr(LCase(sCatalyst),"/dll")+2)
                If InStr(sCatalyst,"/")>0 Then sCatalyst = Left(sCatalyst,InStr(sCatalyst,"/")-1)
                sCatalyst = Trim(Replace(sCatalyst,Chr(34),""))
                If NOT oFso.FileExists(sCatalyst) Then
                    sCatalyst = sCommonProgramFiles & "\OFFICE14\Office Setup Controller\setup.exe"
                    If NOT oFso.FileExists(sCatalyst) AND f64 Then
                        sCatalyst = sCommonProgramFilesX86 & "\OFFICE14\Office Setup Controller\setup.exe"
                    End If
                End If
                If oFso.FileExists(sCatalyst) Then
                    sUninstallCmd = Chr(34) & sCatalyst & Chr(34) & " /uninstall " & Sku & " /config " & Chr(34) & sScrubDir & "\config.xml" & Chr(34) & sDll 
                    iSetupCnt = iSetupCnt + 1
                    Log " - Calling Setup.exe to remove " & Sku '& vbCrLf & sUninstallCmd 
                    If Not fDetectOnly Then 
                        On Error Resume Next
                        RetVal = oWShell.Run(sUninstallCmd,0,True) : CheckError "SetupExeRemoval"
                        Log " - Setup.exe returned: " & SetupExeRetVal(Retval) & " (" & RetVal & ")" & vbCrLf
                        fRebootRequired = fRebootRequired OR (RetVal = "3010")
                        On Error Goto 0
                    Else
                        Log " -> Removal suppressed in preview mode."
                    End If
                Else
                    Log " Error: Office 2010 setup.exe appears to be missing"
                End If 'RetVal = 0) AND oFso.FileExists
            End If 'RegReadValue
        End If 'C2R
    Next 'Sku
    If iSetupCnt = 0 Then Log " Nothing to remove for setup."
End Sub 'SetupExeRemoval
'=======================================================================================================

'Invoke msiexec to remove individual .MSI packages
Sub MsiexecRemoval

    Dim Product
    Dim i
    Dim sCmd, sReturn

    'Clear up ARP first to avoid possible custom action dependencies
    RegWipeARP
    
    'Check MSI registered products
    'Office System does only support per machine installation so it's sufficient to use Installer.Products
    i = 0
    For Each Product in oMsi.Products
        If (Right(Product,PRODLEN) = OFFICEID AND Mid(Product,4,2)="14") Then
            If fRemoveAll OR CheckDelete(Product) Then
                i = i + 1 
                Log " Calling msiexec.exe to remove " & Product
                sCmd = "msiexec.exe /x" & Product & " REBOOT=ReallySuppress NOREMOVESPAWN=True"
                If Mid(Product,11,3)="006" Then 
                    'Need to clear out C2R registration first
                    RegWipeC2R
                End If
                If fQuiet Then 
                    sCmd = sCmd & " /q"
                Else
                    sCmd = sCmd & " /qb-"
                End If
                sCmd = sCmd & " /l*v+ "&chr(34)&sLogDir&"\Uninstall_"&Product&".log"&chr(34)
                If NOT fDetectOnly Then 
                    Log " - Calling msiexec with '"&sCmd&"'"
                    'Execute the patch uninstall
                    sReturn = oWShell.Run(sCmd, 0, True)
                    Log " - msiexec returned: " & SetupExeRetVal(sReturn) & " (" & sReturn & ")" & vbCrLf
                    fRebootRequired = fRebootRequired OR (sReturn = "3010")
                Else
                    Log " -> Removal suppressed in preview mode. Command: "&sCmd
                End If
            End If 'CheckDelete
        End If 'OFFICEID
    Next 'Product
    If i = 0 Then Log " Nothing to remove for msiexec"
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

'Remove the Softgrid services (App-V and Click2Run)
Sub RemoveSG
    On Error Resume Next
    Log " Softgrid CleanUp"
    DeleteService("cvhsvc")
    DeleteService("SftList")
    DeleteService("SftPlay")
    DeleteService("SftVol")
    DeleteService("SftFs")
    DeleteService("SftVsa")

    'Delete the folder
    DeleteFolder sAppdata & "\SoftGrid Client"
    DeleteFolder sLocalAppData & "\SoftGrid Client"
    DeleteFolder sProgramData & "\Microsoft\Application Virtualization Client\SoftGrid Client"
    DeleteFolder sProgramData & "\Microsoft\Application Virtualization Client"
    DeleteFolder sProgramfiles & "\Microsoft\Microsoft Application Virtualization Client"
    DeleteFolder sProgramfiles & "\Microsoft Application Virtualization Client"

    'Delete the registration
    RegDeleteKey HKLM,"SYSTEM\CurrentControlSet\Services\cvhsvc"
    RegDeleteKey HKLM,"SYSTEM\CurrentControlSet\Services\sftfs"
    RegDeleteKey HKLM,"SYSTEM\CurrentControlSet\Services\sftlist"
    RegDeleteKey HKLM,"SYSTEM\CurrentControlSet\Services\sftplay"
    RegDeleteKey HKLM,"SYSTEM\CurrentControlSet\Services\sftredir"
    RegDeleteKey HKLM,"SYSTEM\CurrentControlSet\Services\sftvol"
    RegDeleteKey HKLM,"SYSTEM\CurrentControlSet\Services\sftvsa"
    RegDeleteKey HKLM,"SYSTEM\CurrentControlSet\Services\sftfs"
    RegDeleteKey HKLM,"SOFTWARE\Microsoft\SoftGrid\4.5"
    RegDeleteKey HKCU,"Software\Microsoft\SoftGrid\4.5\Client\AppFS"
    RegDeleteKey HKCU,"Software\Microsoft\SoftGrid\4.5\Client\Applications"
    RegDeleteKey HKCU,"Software\Microsoft\SoftGrid\4.5\Client\FileExtensions"
    RegDeleteKey HKCU,"Software\Microsoft\SoftGrid\4.5\Client\FileTypes"
    RegDeleteKey HKCU,"Software\Microsoft\SoftGrid\4.5\Client\UserInfo"
    'C2R places custom permissions on these regkeys which prevent them from getting deleted
    'RegDeleteKey HKCU,"Software\Microsoft\SoftGrid\4.5\Client\Network"
    'RegDeleteKey HKCU,"Software\Microsoft\SoftGrid\4.5\Client\Packages"
    'RegDeleteKey HKCU,"Software\Microsoft\SoftGrid\4.5\Client"
    'RegDeleteKey HKCU,"Software\Microsoft\SoftGrid\4.5"

End Sub 'RemoveSG
'=======================================================================================================

'Stops all Softgrid services and virtual applications
Sub ResetSG

    Dim Processes, Process
    Dim fWait
    Dim iRet
    
    On Error Resume Next
    
    fWait = False
    Log " Doing Action: ResetSG"

    'Close all running (virtualized) Office applications
    'OfficeVirt.exe needs to be shut down first
    Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process Where Name like 'officevirt%.exe'")
    For Each Process in Processes
        Log " - End process " & Process.Name
        iRet = Process.Terminate()
        CheckError "ResetSG: " & "Process.Name"
        fWait = True
    Next 'Process
    'Shut down CVH.exe 
    Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process Where Name='cvh.exe'")
    For Each Process in Processes
        Log " - End process " & Process.Name
        iRet = Process.Terminate()
        CheckError "ResetSG: " & "Process.Name"
    Next 'Process
    'Close running instances
    Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process")
    For Each Process in Processes
        If dicApps.Exists(LCase(Process.Name)) Then
            Log " - End process " & Process.Name
            iRet = Process.Terminate()
            CheckError "CloseOfficeApps: " & "Process.Name"
        End If
    Next 'Process
    
    If fWait Then wscript.sleep 10000

    'Stop all SoftGrid services
    iRet = StopService("cvhsvc")
    iRet = StopService("SftList")
    iRet = StopService("SftPlay")
    iRet = StopService("SftVol")
    iRet = StopService("SftFs")
    iRet = StopService("SftVsa")
End Sub 'ResetSG
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
                    If fRemoveAll Then 
                        If  (Mid(Subfolder.Name,26,PRODLEN) = OFFICEID AND Mid(SubFolder.Name,4,2)="14") OR _
                            LCase(Right(Subfolder.Name,7)) = "14.data" Then DeleteFolder Subfolder.Path
                    Else
                        If  (Mid(Subfolder.Name,26,PRODLEN) = OFFICEID AND Mid(SubFolder.Name,4,2)="14") AND _
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

'Wipe files and folders as documented in KB 928218
Sub FileWipeAll
    Dim sFolder
    Dim Folder, Subfolder
    
    If fForce OR fQuiet Then CloseOfficeApps
    
    'Handle other services.
    If fRemoveOSE Then DeleteService "osppsvc"
    DeleteService "odserv"
    DeleteService "Microsoft Office Groove Audit Service"
    DeleteService "Microsoft SharePoint Workspace Audit Service"
    
    'User specific files
    If NOT fKeepUser Then
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
        End If 'oFso.FolderExists(sFolder)
    End If  
    
    'Run the individual filewipe from component detection first
    FileWipeIndividual
    
    'Take care of the rest
    DeleteFolder sOInstallRoot
    DeleteFolder sCommonProgramFiles & "\Microsoft Shared\Office14"
    DeleteFolder sCommonProgramFiles & "\Microsoft Shared\OfficeSoftwareProtectionPlatform"
    DeleteFile sAllUsersProfile & "\Application Data\Microsoft\Office\Data\opa14.dat"
    DeleteFile oWShell.SpecialFolders("AllUsersStartup")&"\OfficeSAS.lnk"
    DeleteFile oWShell.SpecialFolders("Startup")&"\OneNote 2010 Screen Clipper and Launcher.lnk"
End Sub 'FileWipeAll
'=======================================================================================================

'Wipe individual files & folders related to SKU's that are no longer installed
Sub FileWipeIndividual
    Dim LogicalDisks, Disk
    Dim File, Files, XmlFile, scFiles, oFile, Folder, SubFolder, Processes, Process
    Dim sFile, sFolder, sPath, sConfigName, sContents, sProductCode, sLocalDrives,sScQuery
    Dim arrSubfolders
    Dim fKeepFolder, fDeleteSC
    Dim iRet
    
    Log " File CleanUp"
    If IsArray(arrDeleteFiles) Then
        If fForce OR fQuiet Then
            Log " Doing Action: StopOSE"
            iRet = StopService("ose")
            Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Service Where Name like 'ose%.exe'")
            For Each Process in Processes
                LogOnly " - Running process : " & Process.Name
                Log " -> Ending process: " & Process.Name
                iRet = Process.Terminate()
            Next 'Process
            LogOnly " End Action: StopOSE"
            CloseOfficeApps
        End If
        'Wipe individual files detected earlier
        LogOnly " Removing left behind files"
        For Each sFile in arrDeleteFiles
            If oFso.FileExists(sFile) Then DeleteFile sFile
        Next 'File
    End If 'IsArray
    
    'Wipe Catalyst in commonfiles
    sFolder = sCommonProgramFiles & "\microsoft shared\OFFICE14\Office Setup Controller\"
    If EnumFolderNames(sFolder,arrSubFolders) Then
        For Each SubFolder in arrSubFolders
            sPath = sFolder & SubFolder
            If InStr(SubFolder,".")>0 Then sConfigName = UCase(Left(SubFolder,InStr(SubFolder,".")-1))Else sConfigName = UCase(Subfolder)
            If GetFolderPath(sPath) Then
                Set Folder = oFso.GetFolder(sPath)
                Set Files = Folder.Files
                fKeepFolder = False
                For Each File In Files
                    If (LCase(Right(File.Name,4))=".xml") AND (UCase(Left(File.Name,Len(sConfigName)))=sConfigName) Then
                        Set XmlFile = oFso.OpenTextFile(File,1)
                        sContents = XmlFile.ReadAll
                        Set XmlFile = Nothing
                        sProductCode = Mid(sContents,InStr(sContents,"ProductCode=")+Len("ProductCode=")+1,38)
                        If CheckDelete(sProductCode) Then DeleteFile File.Path Else fKeepFolder = True
                    End If
                Next 'File
                Set Files = Nothing
                Set Folder = Nothing
                If Not fKeepFolder Then DeleteFolder sPath
            End If 'GetFolderPath
        Next 'SubFolder
    End If 'EnumFolderNames
    
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
                'Handle C2R
                If InStr(File.Target,"CVH.EXE")>0 AND fRemoveC2R Then
                    If InStr(File.Target,"9014006")>0 Then fDeleteSC = True
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
    
    On Error Resume Next
    If oFso.FileExists(sScrubDir&"\CvhbsQuiet.vbs") Then oFso.DeleteFile sScrubDir&"\CvhbsQuiet.vbs",True
    If oFso.FolderExists(sScrubDir & "\ScrubTmp") Then oFso.DeleteFolder sScrubDir & "\ScrubTmp",True

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
                    If InStr(UCase(MspTargets(File.Path)),OFFICEID)>0 Then DeleteFile File.Path
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
    Dim Item, Name, Sku, key
    Dim hDefKey, sSubKeyName, sCurKey, sValue, sGuid
    Dim fkeep, fSystemComponent0, fPackages, fDisplayVersion
    Dim arrKeys, arrNames, arrTypes
    Dim iLoopCnt, iPos
    
    Log " Registry CleanUp"
    'Wipe registry data
    
    'User Profile settings
    RegDeleteKey HKCU,"Software\Policies\Microsoft\Office\14.0"
    If NOT fKeepUser Then
        RegDeleteKey HKCU,"Software\Microsoft\Office\14.0"
    End If 'fKeepUser
    
    'Computer specific settings
    If fRemoveAll Then
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Office\14.0"
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Office Test"
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform"
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform_Test"
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Office\Common","LastAccessInstall"
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Office\Common","MID"
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Office\Common\ActiveX Compatibility\{00024512-0000-0000-C000-000000000046}"
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Office\Common\OffDiag\Location","14"
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Office\Excel\Addins\Microsoft.PerformancePoint.Planning.Client.Excel"
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Office\InfoPath\Converters\Import\InfoPath.DesignerExcelImport\Versions","14.0"
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Office\InfoPath\Converters\Import\InfoPath.DesignerWordImport\Versions","14.0"
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Office\OneNote\Adapters","{456B0D0E-49DD-4C95-8DB6-175F54DE69A3}"
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Office\Outlook"
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Shared Tools\Text Converters\Export\MEWord12"
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Shared Tools\Text Converters\Export\Word12"
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Shared Tools\Text Converters\Export\Word97"
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Shared Tools\Text Converters\Import\MEWord12"
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Shared Tools\Text Converters\Import\Word12"
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Shared Tools\Text Converters\Import\Word97"
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Terminal Server\Install\Software\Microsoft\Office\14.0"
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Office\Common\OffDiag\Location","14"
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Run","GrooveMonitor"
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Run","LobiServer"
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Run","BCSSync"
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved","{42042206-2D85-11D3-8CFF-005004838597}"
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved","{993BE281-6695-4BA5-8A2A-7AACBFAAB69E}"
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved","{0006F045-0000-0000-C000-000000000046}"
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved","{C41662BB-1FA0-4CE0-8DC5-9B7F8279FF97}"
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved","{7CCA70DB-DE7A-4FB7-9B2B-52E2335A3B5A}"
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved","{506F4668-F13E-4AA1-BB04-B43203AB3CC0}"
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved","{D66DC78C-4F61-447F-942B-3FB6980118CF}"
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{B4F3A835-0E21-4959-BA22-42B3008E02FF}"
        'Groove Extensions 
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellExecuteHooks","{B5A7F190-DDA6-4420-B3BA-52453494E6CD}"
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved","{99FD978C-D287-4F50-827F-B2C658EDA8E7}"
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved","{AB5C5600-7E6E-4B06-9197-9ECEF74D31CC}"
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved","{920E6DB1-9907-4370-B3A0-BAFC03D81399}"
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved","{16F3DD56-1AF5-4347-846D-7C10C4192619}"
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved","{2916C86E-86A6-43FE-8112-43ABE6BF8DCC}"
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved","{72853161-30C5-4D22-B7F9-0BBC1D38A37E}"
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved","{6C467336-8281-4E60-8204-430CED96822D}"
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved","{2A541AE1-5BF6-4665-A8A3-CFA9672E4291}"
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved","{B5A7F190-DDA6-4420-B3BA-52453494E6CD}"
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved","{A449600E-1DC6-4232-B948-9BD794D62056}"
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved","{3D60EDA7-9AB4-4DA8-864C-D9B5F2E7281D}"
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved","{387E725D-DC16-4D76-B310-2C93ED4752A0}"
        RegDeleteKey HKLM,"SOFTWARE\Classes\*\shellex\ContextMenuHandlers\XXX Groove GFS Context Menu Handler XXX"
        RegDeleteKey HKLM,"SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\XXX Groove GFS Context Menu Handler XXX"
        RegDeleteKey HKLM,"SOFTWARE\Classes\Directory\shellex\ContextMenuHandlers\XXX Groove GFS Context Menu Handler XXX"
        RegDeleteKey HKLM,"SOFTWARE\Classes\Folder\ShellEx\ContextMenuHandlers\XXX Groove GFS Context Menu Handler XXX"
        RegDeleteKey HKLM,"SOFTWARE\Classes\Directory\Background\shellex\ContextMenuHandlers\XXX Groove GFS Context Menu Handler XXX"
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Groove Explorer Icon Overlay 1 (GFS Unread Stub)"
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Groove Explorer Icon Overlay 2 (GFS Stub)"
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Groove Explorer Icon Overlay 2.5 (GFS Unread Folder)"
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Groove Explorer Icon Overlay 3 (GFS Folder)"
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Groove Explorer Icon Overlay 4 (GFS Unread Mark)"
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{72853161-30C5-4D22-B7F9-0BBC1D38A37E}"
        
        RegDeleteKey HKLM,"SYSTEM\CurrentControlSet\Services\Outlook"
        
        'Win32Assemblies
        If RegEnumKey(HKCR,"Installer\Win32Assemblies\",arrKeys) Then
            For Each Item in arrKeys
                If InStr(UCase(Item),"OFFICE14")>0 Then RegDeleteKey HKCR,"Installer\Win32Assemblies\"&Item
            Next 'Item
        End If 'RegEnumKey
        'Groove blocks reinstall if it locates groove.exe over this key
        If RegKeyExists(HKCR,"GrooveFile\Shell\Open\Command\") Then
            sValue = ""
            RegReadValue HKCR,"GrooveFile\Shell\Open\Command\","",sValue,"REG_SZ"
            If InStr(sValue,"\Office14\")>0 Then RegDeleteKey HKCR,"GrooveFile"
        End If 'RegKeyExists
    End If 'fRemoveAll

    'Known Keypath settings
    For Each key in dicDelRegKey.Keys
        If Right(key,1) = "\" Then
            RegDeleteKey dicDelRegKey.Item(key),key
        Else
            iPos = InStrRev(Key,"\")
            If iPos > 0 Then RegDeleteValue dicDelRegKey.Item(key), Left(key,iPos - 1), Mid(key,iPos+1)
        End If
    Next

    'Add/Remove Programs
    RegWipeARP 

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
                    'Check if it's a Office 2010 key
                    If Right(sGuid,PRODLEN)=OFFICEID AND Mid(sGuid,4,2)="14" Then
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
            If fRemoveAll Then
                If (Mid(Item,26,PRODLEN)=OFFICEID AND Mid(Item,4,2)="14") OR _
                   LCase(Right(Item,7))="14.data" Then RegDeleteKey HKLM,sSubKeyName & Item
            Else
                If (Mid(Item,26,PRODLEN)=OFFICEID AND Mid(Item,4,2)="14") AND _
                   CheckDelete(UCase(Left(Item,38))) Then RegDeleteKey HKLM,sSubKeyName & Item
            End If
        Next 'Item
    End If 'RegEnumKey
    
    'Registration
    hDefKey = HKLM
    sSubKeyName = "SOFTWARE\Microsoft\Office\14.0\Registration\"
    If RegEnumKey(HKLM,sSubKeyName,arrKeys) Then
        For Each Item in arrKeys
            If CheckDelete(UCase(Left(Item,38))) Then RegDeleteKey HKLM,sSubKeyName & Item
        Next 'Item
    End If 'RegEnumKey
    
    'User Preconfigurations
    hDefKey = HKLM
    sSubKeyName = "SOFTWARE\Microsoft\Office\14.0\User Settings\"
    If RegEnumKey(HKLM,sSubKeyName,arrKeys) Then
        For Each Item in arrKeys
            If CheckDelete(UCase(Left(Item,38))) Then RegDeleteKey HKLM,sSubKeyName & Item
        Next 'Item
    End If 'RegEnumKey

    'Click2Run Cleanup
    RegWipeC2R 

    'Temporary entries in ARP
    TmpKeyCleanUp
End Sub 'RegWipeAll
'=======================================================================================================

'Clean up Add/Remove Programs registry
Sub RegWipeARP

    Dim Item, Name, Sku, key
    Dim sSubKeyName, sCurKey, sValue, sGuid
    Dim fkeep, fSystemComponent0, fPackages, fDisplayVersion
    Dim arrKeys

    'Add/Remove Programs
    sSubKeyName = REG_ARP
    If RegEnumKey(HKLM,sSubKeyName,arrKeys) Then
        For Each Item in arrKeys
            '*0FF1CE*
            If Len(Item)>37 Then
                sGuid = UCase(Left(Item,38))
                If Right(sGuid,PRODLEN)=OFFICEID AND Mid(sGuid,4,2)= OVERSIONMAJOR Then
                    If CheckDelete(sGuid) Then RegDeleteKey HKLM, sSubKeyName & Item
                End If 'Right(Item,PRODLEN)=OFFICEID
            End If 'Len(Item)>37
            
            'Config entries
            sCurKey = sSubKeyName & Item & "\"
            fSystemComponent0 = Not (RegReadValue(HKLM,sCurKey,"SystemComponent",sValue,"REG_DWORD") AND (sValue = "1"))
            fPackages = RegReadValue(HKLM,sCurKey,OPACKAGE,sValue,"REG_MULTI_SZ")
            fDisplayVersion = RegReadValue(HKLM,sCurKey,"DisplayVersion",sValue,"REG_SZ")
            If fDisplayVersion AND Len(sValue) > 1 Then
                fDisplayVersion = (Left(sValue,2) = OVERSIONMAJOR)
            End If
            If (fSystemComponent0 AND fPackages AND fDisplayVersion) OR (fSystemComponent0 AND fDisplayVersion AND InStr(UCase(Item),"CLICK2RUN")>0) Then
                fKeep = False
                If Not fRemoveAll Then
                    For Each Sku in dicKeepSku.Keys
                        If UCase(Item) =  OREGREF & Sku Then
                            fkeep = True
                            Exit For
                        End If
                    Next 'Sku
                End If
                If Not fkeep Then RegDeleteKey HKLM, sSubKeyName & Item
            End If
        Next 'Item
    End If 'RegEnumKey

End Sub 'RegWipeARP
'=======================================================================================================

'Clean up Click2Run specific registrations
Sub RegWipeC2R

    Dim Item
    Dim sSubKeyName
    Dim arrKeys

    'Click2Run Cleanup
    If fRemoveAll OR fRemoveC2R Then
        RegDeleteKey HKCU,"Software\Microsoft\Office\CVH"
        RegDeleteKey HKCU,"Software\Microsoft\Office\14.0\CVH"
        RegDeleteKey HKLM,"Software\Microsoft\Office\14.0\CVH"
        RegDeleteKey HKLM,"Software\Microsoft\Office\14.0\CVHSettings"
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Office\14.0\Common\InstallRoot\Virtual"

        'Control Panel Items
        RegDeleteKey HKLM,"Software\Microsoft\Windows\CurrentVersion\explorer\ControlPanel\NameSpace\{F9ACD2D6-09C8-4103-995C-912DE68DDE1E}"
        RegDeleteKey HKCR,"CLSID\{F9ACD2D6-09C8-4103-995C-912DE68DDE1E}"
        RegDeleteKey HKLM,"Software\Microsoft\Windows\CurrentVersion\explorer\ControlPanel\NameSpace\{005CB1F2-224F-4738-B051-91A96758F50C}"
        RegDeleteKey HKCR,"CLSID\{005CB1F2-224F-4738-B051-91A96758F50C}"

        sSubKeyName = "SOFTWARE\Microsoft\SoftGrid\4.5\Client\Packages\"
        If RegEnumKey(HKLM,sSubKeyName,arrKeys) Then
            For Each Item in arrKeys
                If CheckDelete(Item) Then RegDeleteKey HKLM,sSubKeyName & Item
            Next 'Item
        End If 'RegEnumKey
        If RegEnumKey(HKCU,sSubKeyName,arrKeys) Then
            For Each Item in arrKeys
                If CheckDelete(Item) Then RegDeleteKey HKLM,sSubKeyName & Item
            Next 'Item
        End If 'RegEnumKey
    End If

End Sub 'RegWipeC2R
'=======================================================================================================

'Clean up temporary registry keys
Sub TmpKeyCleanUp
    Dim TmpKey
    
    If fLogInitialized Then Log " Remove temporary registry entries"
    If IsArray(arrTmpSKUs) Then
        For Each TmpKey in arrTmpSKUs
            oReg.DeleteKey HKLM, REG_ARP & TmpKey
        Next 'Item
    End If 'IsArray
End Sub 'TmpKeyCleanUp

'=======================================================================================================
' Helper Functions
'=======================================================================================================

'End all running instances of applications that will be removed
Sub CloseOfficeApps
    Dim Processes, Process
    Dim fWait
    Dim iRet
    
    On Error Resume Next
    
    fWait = False
    Log " Doing Action: CloseOfficeApps"

    'OfficeVirt.exe needs to be shut down first
    Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process Where Name like 'officevirt%.exe'")
    For Each Process in Processes
        If dicApps.Exists(LCase(Process.Name)) Then
            Log " - End process " & Process.Name
            iRet = Process.Terminate()
            CheckError "CloseOfficeApps: " & "Process.Name"
            fWait = True
        End If
    Next 'Process

    Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process")
    For Each Process in Processes
        If dicApps.Exists(LCase(Process.Name)) Then
            Log " - End process " & Process.Name
            iRet = Process.Terminate()
            CheckError "CloseOfficeApps: " & "Process.Name"
            If Process.Name = "CVH.EXE" Then fWait = True
        End If
    Next 'Process
    If fWait Then
        wscript.sleep 10000
    End If
    LogOnly " End Action: CloseOfficeApps"
End Sub 'CloseOfficeApps
'=======================================================================================================

'CVHBS.exe has no true unattended option
'To ensure quiet automation does not break this dialog box handler monitors the process
Sub CvhbsDialogHandler

Dim CvhbsQuiet
Dim sRunCmd, sQuote

Set CvhbsQuiet = oFso.CreateTextFile(sScrubDir&"\CvhbsQuiet.vbs",True,True)
sQuote = "&chr(34)&"
CvhbsQuiet.WriteLine "On Error Resume Next"
CvhbsQuiet.WriteLine "Set oShell = CreateObject("&chr(34)&"WScript.Shell"&chr(34)&")"
CvhbsQuiet.WriteLine "Set oWmiLocal   = GetObject("&chr(34)&"winmgmts:\\.\root\cimv2"&chr(34)&")"
CvhbsQuiet.WriteLine "wscript.sleep 10000"
CvhbsQuiet.WriteLine "Do"
    CvhbsQuiet.WriteLine "Set Processes = oWmiLocal.ExecQuery("&chr(34)&"Select * From Win32_Process Where Name='cvhbs.exe'"&chr(34)&")"
    CvhbsQuiet.WriteLine "iCnt = Processes.Count"
    CvhbsQuiet.WriteLine "If iCnt > 0 Then"
        CvhbsQuiet.WriteLine "sCommand = "&chr(34)&"tasklist /FI "&chr(34)&sQuote&chr(34)&"WINDOWTITLE eq click*"&chr(34)&sQuote&chr(34)&" /FO CSV /NH"&chr(34)
        CvhbsQuiet.WriteLine "Set oExec = oShell.Exec(sCommand)"
        CvhbsQuiet.WriteLine "sCmdOut = oExec.StdOut.ReadAll()"
        CvhbsQuiet.WriteLine "Do While oExec.Status = 0"
             CvhbsQuiet.WriteLine "WScript.Sleep 200"
        CvhbsQuiet.WriteLine "Loop"

        CvhbsQuiet.WriteLine "If InStr(sCmdOut,"&chr(34)&","&chr(34)&")>0 Then"
            CvhbsQuiet.WriteLine "sCmdOut = Replace(sCmdOut,chr(34),"&chr(34)&chr(34)&")"
            CvhbsQuiet.WriteLine "arrCol = Split(sCmdOut,"&chr(34)&","&chr(34)&")"
                CvhbsQuiet.WriteLine "sPid = arrCol(1)"
                CvhbsQuiet.WriteLine "oShell.AppActivate sPID"
                CvhbsQuiet.WriteLine "oShell.SendKeys "&chr(34)&"{ENTER}"&chr(34)
        CvhbsQuiet.WriteLine "End If"

    CvhbsQuiet.WriteLine "End If"
    CvhbsQuiet.WriteLine "wscript.sleep 10000"
CvhbsQuiet.WriteLine "Loop While iCnt > 0"
CvhbsQuiet.Close

sRunCmd = "cscript "&chr(34)&sScrubDir&"\CvhbsQuiet.vbs"&chr(34)
oWShell.Run sRunCmd, 0, False

End Sub 'CvhbsDialogHandler

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

'Check if an Office 14 product is still registered with a SKU that stays on the computer
Function CheckDelete(sProductCode)
        
    'If it's a non Office 14 ProductCode exit with false right away
    CheckDelete = Right(sProductCode,PRODLEN) = OFFICEID
    If CheckDelete Then CheckDelete =  Mid(sProductCode,4,2) = "14"
    If Not CheckDelete Then Exit Function
    If dicKeepProd.Exists(UCase(sProductCode)) Then CheckDelete = False
End Function 'CheckDelete
'=======================================================================================================

'Returns a string with a list of ProductCodes from the summary information stream
Function MspTargets (sMspFile)
    Const MSIOPENDATABASEMODE_PATCHFILE = 32
    Const PID_TEMPLATE                  =  7
    
    Dim Msp
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
    Dim File
    Dim sFileName, sNewPath
    
    On Error Resume Next
    If oFso.FileExists(sFile) Then
        LogOnly " - Delete file: " & sFile
        If Not fDetectOnly Then oFso.DeleteFile sFile,True
        If Err <> 0 Then
            CheckError "DeleteFile"
            If fForce Then
                'Try to move the file and delete from there
                Set File = oFso.GetFile(sFile)
                sFileName = File.Name
                sNewPath = sScrubDir & "\ScrubTmp"
                Set File = Nothing
                If Not oFso.FolderExists(sNewPath) Then oFso.CreateFolder(sNewPath)
                'Move the file
                LogOnly " - Move file to: " & sNewPath & "\" & sFileName
                oFso.MoveFile sFile,sNewPath & "\" & sFileName
                If Err <> 0 Then 
                    CheckError "DeleteFile (move)"
                End If 'Err <> 0
            End If 'fForce
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
        sNewPath = sScrubDir & "\ScrubTmp"
        Set Folder = Nothing
        'Ensure we stay within the same drive
        If Not oFso.FolderExists(sNewPath) Then oFso.CreateFolder(sNewPath)
        'Move the folder
        LogOnly " - Moving folder to: " & sNewPath & "\" & sFolderName
        oFso.MoveFolder sFolder,sNewPath & "\" & sFolderName
        If Err <> 0 Then
            CheckError "DeleteFolder (move)"
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
    
    If dicKeepReg.Exists(LCase(sSubKeyName)) Then Exit Sub
    If f64 Then
        If dicKeepReg.Exists(LCase(Wow64Key(hDefKey, sSubKeyName))) Then Exit Sub
    End If
    If Len(sSubKeyName) > 1 Then
        If Right(sSubKeyName,1) = "\" Then sSubKeyName = Left(sSubKeyName,Len(sSubKeyName)-1)
    End If
    
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
    Dim iRet

    On Error Resume Next
    
    iRet = 0
    sQuery = "Select * From Win32_Service Where Name='" & sService & "'"
    Set Services = oWmiLocal.Execquery(sQuery)
    'Stop the service
    For Each Service in Services
        If UCase(Service.State) = "STARTED" Then iRet = Service.StopService
        If UCase(Service.State) = "RUNNING" Then iRet = Service.StopService

    Next 'Service
    StopService = (iRet = 0)
End Function 'StopService
'=======================================================================================================

'Delete a service
Sub DeleteService(sService)
    Dim Services, Service, Processes, Process
    Dim sQuery, sStates
    Dim iRet
    
    On Error Resume Next
    
    sStates = "STARTED;RUNNING"
    sQuery = "Select * From Win32_Service Where Name='" & sService & "'"
    Set Services = oWmiLocal.Execquery(sQuery)
    
    'Stop and delete the service
    For Each Service in Services
        Log " Found service " & sService & " in state " & Service.State
        If InStr(sStates,UCase(Service.State))>0 Then iRet = Service.StopService()
        'Ensure no more instances of the service are running
        Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process Where Name='" & sService & ".exe'")
        For Each Process in Processes
            iRet = Process.Terminate()
        Next 'Process
        If Not fDetectOnly Then 
            Log " - Deleting Service -> " & sService
            iRet = Service.Delete()
        Else
            Log " - Simulate deleting Service -> " & sService
        End If
    Next 'Service
    Set Services = Nothing
    Err.Clear

End Sub 'DeleteService
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
        
        Case "000F" : sReturn = "MONDO"
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
        Case "003D" : sReturn = "SINGLEIMAGE"
        Case "0043" : sReturn = "OFFICE32"
        Case "0044" : sReturn = "InfoPath"
        Case "0045" : sReturn = "XWEB"
        Case "0048" : sReturn = "OLC"
        Case "0049" : sReturn = "ACADEMIC"
        Case "004A" : sReturn = "OWC11"
        Case "0051" : sReturn = "VISPRO"
        Case "0052" : sReturn = "VisView"
        Case "0053" : sReturn = "VisStd"
        Case "0054" : sReturn = "VisMUI"
        Case "0055" : sReturn = "VisMUI"
        Case "0057" : sReturn = "VISIO"
        Case "0061" : sReturn = "CLICK2RUN"
        Case "0062" : sReturn = "CLICK2RUN"
        Case "0066" : sReturn = "CLICK2RUN"
        Case "006C" : sReturn = "CLICK2RUN"
        Case "006D" : sReturn = "CLICK2RUN"
        Case "006E" : sReturn = "Shared"
        Case "006F" : sReturn = "OFFICE"
        Case "0074" : sReturn = "STARTER"
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
        Case "011A" : sReturn = "SendASmile"
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

Sub LogOnly (sLog)
    LogStream.WriteLine sLog
End Sub 'Log
'=======================================================================================================

Sub CheckError(sModule)
    If Err <> 0 Then 
        LogOnly "   " & Now & " - " & sModule & " - Source: " & Err.Source & "; Err# (Hex): " & Hex( Err ) & _
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
        FindInstalledOProducts
        If dicInstalledSku.Count > 0 Then sDefault = Join(RemoveDuplicates(dicInstalledSku.Items),",") Else sDefault = "CLIENTALL"
        sDefault = InputBox("Enter a list of Office 2010 products to remove" & vbCrLf & vbCrLf & _
                "Examples:" & vbCrLf & _
                "CLIENTALL" & vbTab & "-> remove all Client products" & vbCrLf & _
                "ALL" & vbTab & vbTab & "-> remove all, including Server!" & vbCrLf & _
                "ProPlus,PrjPro" & vbTab & "-> remove ProPlus and Project" & vbCrLf &_
                "?" & vbTab & vbTab & "-> display Help", _
                "OffScrub10 - Office 2010 remover", _
                sDefault)
        If IsEmpty(sDefault) Then 'User cancelled
            Log "User cancelled. CleanUp & Exit."
            'Undo temporary entries created in ARP
            TmpKeyCleanUp
            wscript.quit 1602
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
    Select Case UCase(sArg0)
    
    Case "?"
        ShowSyntax
    
    Case "ALL"
        fRemoveAll = True
        fRemoveOSE = False
    
    Case "CLIENTSUITES"
        fRemoveCSuites = True
        fRemoveOSE = False
    
    Case "CLIENTSTANDALONE"
        fRemoveCSingle = True
        fRemoveOSE = False

    Case "CLIENTALL"
        fRemoveCSuites = True
        fRemoveCSingle = True
        fRemoveOSE = False
    
    Case "SERVER"
        fRemoveServer = True
        fRemoveOSE = False

    Case "ALL,OSE"
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
        
        Case "/FR","/FASTREMOVE"
            fBypass_Stage1 = True
            fSkipSD = True
        
        Case "/F","/FORCE"
            fForce = True
            fRemoveOSE = True
            fRemoveAppV = True
        
        Case "/K","/KEEPUSERSETTINGS"
            fKeepUser = True
        
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
        
        Case "/N","/NOCANCEL"
            fNoCancel = True
        
        Case "/O","/OSE"
            fRemoveOSE = True
        
        Case "/P","/PREVIEW","/DETECTONLY"
            fDetectOnly = True
        
        Case "/Q","/QUIET"
            fQuiet = True
        
        Case "/QND"
            fBypass_Stage1 = True
            fBypass_Stage2 = True
            fBypass_Stage3 = True
            fRemoveOSE = True
            fRemoveAll = True
            fSkipSD = True
            fForce = True
        
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

    Log "Microsoft Customer Support Services - Office 2010 Removal Utility" & vbCrLf & vbCrLf & _
                "Version: " & SCRIPTVERSION & vbCrLf & _
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
    TmpKeyCleanUp
    Wscript.Echo sErr & vbCrLf & _
             SCRIPTFILE & " V " & SCRIPTVERSION & vbCrLf & _
             "Copyright (c) Microsoft Corporation. All Rights Reserved" & vbCrLf & vbCrLf & _
             SCRIPTFILE & " helps to remove " & ONAME & " Server & Client products" & vbCrLf & _
             "when a regular uninstall is no longer possible" & vbCrLf & vbCrLf & _
             "Usage:" & vbTab & SCRIPTFILE & " [List of config ProductIDs] [Options]" & vbCrLf & vbCrLf & _
             vbTab & "/?                               ' Displays this help"& vbCrLf &_
             vbTab & "/Force                           ' Enforces file removal. May cause data loss!" & vbCrLf &_
             vbTab & "/SkipShortcutDetection           ' Does not search the local hard drives for shortcuts" & vbCrLf & _
             vbTab & "/Log [LogfolderPath]             ' Custom folder for log files" & vbCrLf & _
             vbTab & "/NoCancel                        ' Setup.exe and Msiexec.exe have no Cancel button" & vbCrLf &_
             vbTab & "/OSE                             ' Forces removal of the Office Source Engine service" & vbCrLf &_
             vbTab & "/Quiet                           ' Setup.exe and Msiexec.exe run quiet with no UI" & vbCrLf &_
             vbTab & "/Preview                         ' Run this script to preview what would get removed"& vbCrLf & vbCrLf & _
             "Examples:"& vbCrLf & _
             vbTab & SCRIPTFILE & " CLIENTALL         ' Remove all " & ONAME & " Client products" & vbCrLf &_
             vbTab & SCRIPTFILE & " ALL               ' Remove all " & ONAME & " Server & Client products" & vbCrLf &_
             vbTab & SCRIPTFILE & " ProPlus,PrjPro    ' Remove ProPlus and Project" & vbCrLf
    Wscript.Quit
End Sub 'ShowSyntax
'=======================================================================================================