
Imports System.IO
Imports System.Net.WebRequestMethods
Imports System.Runtime.InteropServices
Imports System.Security.AccessControl
Imports Autodesk.Connectivity.WebServices
Imports Autodesk.DataManagement.Client.Framework.Forms.Internal
'Imports Connectivity.Application.VaultBase
Imports Autodesk.DataManagement.Client.Framework.Library
Imports Connectivity.Application
Imports Connectivity.Application.VaultBase
Imports AWS = Autodesk.Connectivity.WebServices
Imports VCM = Autodesk.DataManagement.Client.Framework.Vault.Currency.Connections
'Imports VB = Connectivity.Application.VaultBase
Imports VDF = Autodesk.DataManagement.Client.Framework
Imports VFM = Autodesk.DataManagement.Client.Framework.Vault.Forms

Module Vault
    'Public Sub KillItWithFire(serverName As String, vaultName As String, userName As String, password As String, targetFolderPath As String)

    '    ' Create ServerIdentities - REQUIRED for Vault 2024+ credentials
    '    Dim serverIdentities As New AWS.ServerIdentities()
    '    serverIdentities.DataServer = serverName
    '    serverIdentities.FileServer = serverName

    '    ' Use ServerIdentities in credentials constructor
    '    Dim credentials As New ACW.UserPasswordCredentials(serverIdentities, vaultName, userName, password, AWS.LicensingAgent.Client)
    '    Dim serviceManager As New ACW.WebServiceManager(credentials)

    '    Try
    '        Dim targetFolder As ACW.DocumentService.Folder = serviceManager.DocumentService.GetFolderByPath(targetFolderPath)

    '        If targetFolder Is Nothing Then
    '            Console.WriteLine("Folder not found: " & targetFolderPath)
    '            Return
    '        End If

    '        DeleteFolderContentsRecursively(serviceManager.DocumentService, targetFolder.Id, serviceManager)

    '        Console.WriteLine("Deletion complete. All files and subfolders moved to Deleted area in " & targetFolderPath)
    '        Console.WriteLine("Next step: Manually purge in Vault Client (View > Deleted > select all > Purge).")

    '    Catch ex As Exception
    '        Console.WriteLine("Error: " & ex.Message)
    '    Finally
    '        serviceManager.Dispose()
    '    End Try

    'End Sub

    'Private Sub DeleteFolderContentsRecursively(docService As DocumentService, folderId As Long, serviceManager As WebServiceManager)

    '    ' Recurse into subfolders first (bottom-up deletion)
    '    Dim subFolders() As Folder = docService.GetFoldersByParentId(folderId, False)

    '    If subFolders IsNot Nothing Then
    '        For Each subFolder As Folder In subFolders
    '            DeleteFolderContentsRecursively(docService, subFolder.Id, serviceManager)

    '            Try
    '                ' Folders must be empty before deletion; recursion ensures this
    '                docService.DeleteFolderHierarchyUnconditional(subFolder.Id)
    '                Console.WriteLine("Deleted subfolder: " & subFolder.FullName)
    '            Catch ex As Exception
    '                Console.WriteLine("Failed to delete subfolder " & subFolder.FullName & ": " & ex.Message)
    '            End Try
    '        Next
    '    End If

    '    ' Delete all latest files in this folder
    '    Dim files() As File = docService.GetLatestFilesByFolderId(folderId, False)

    '    If files IsNot Nothing AndAlso files.Length > 0 Then
    '        For Each file As File In files
    '            Try
    '                ' DeleteFileFromFolder requires MasterId and parent FolderId
    '                docService.DeleteFileFromFolder(file.MasterId, folderId)
    '                Console.WriteLine("Deleted file: " & file.Name & " (MasterId: " & file.MasterId & ")")
    '            Catch ex As Exception
    '                Console.WriteLine("Failed to delete file " & file.Name & ": " & ex.Message)
    '                ' Optional: Try unconditional delete if available (admin bypass)
    '                Try
    '                    docService.DeleteFileFromFolderUnconditional(file.MasterId, folderId)
    '                    Console.WriteLine("Force-deleted file (unconditional): " & file.Name)
    '                Catch innerEx As Exception
    '                    Console.WriteLine("Unconditional delete also failed for " & file.Name & ": " & innerEx.Message)
    '                End Try
    '            End Try
    '        Next
    '    End If

    'End Sub

    Sub GetBigDXFs(serverName As String, vaultName As String, Optional ByVal lstFilePaths As List(Of String) = Nothing)

        If lstFilePaths Is Nothing Then
            lstFilePaths = New List(Of String)
        End If
        Dim mVltCon As VDF.Vault.Currency.Connections.Connection = VaultLogin(serverName, vaultName)

        If mVltCon Is Nothing Then
            MsgBox("Vault Login Failed - Exiting")
            Exit Sub
        End If

        Dim files As AWS.File
        ' Dim filePropDefs As New AWS.PropDef() = AWS.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE")
        Dim filePropDefs As PropDef() = mVltCon.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE")

        Dim fileExtPropDef As PropDef '= filePropDefs.DispName("File Extension")
        For Each def As PropDef In filePropDefs
            If def.DispName = "File Extension" Then
                fileExtPropDef = def
                Exit For
            End If
        Next
        Dim fileExtDXF As New SrchCond() With {
            .PropDefId = fileExtPropDef.Id,
            .PropTyp = PropertySearchType.SingleProperty,
            .SrchOper = 3, 'Equals
            .SrchRule = SearchRuleType.Must,
            .SrchTxt = "dxf"
        }

        Dim fileSizePropDef As PropDef '= filePropDefs.DispName("File Size")
        For Each def As PropDef In filePropDefs
            If def.DispName = "File Size" Then
                fileSizePropDef = def
                Exit For
            End If
        Next
        Dim fileSize As New SrchCond() With {
            .PropDefId = fileSizePropDef.Id,
            .PropTyp = PropertySearchType.SingleProperty,
            .SrchOper = 6, 'Greater Than
            .SrchRule = SearchRuleType.Must,
            .SrchTxt = "41940000"
        }

        Dim filePathPropDef As PropDef '= filePropDefs.DispName("File Size")
        For Each def As PropDef In filePropDefs
            If def.DispName = "Folder Path" Then
                filePathPropDef = def
                Exit For
            End If
        Next
        Dim filePath As New SrchCond() With {
            .PropDefId = filePathPropDef.Id,
            .PropTyp = PropertySearchType.SingleProperty,
            .SrchOper = 3, 'Equals
            .SrchRule = SearchRuleType.Must,
            .SrchTxt = "$/Designs/_Part Drawings - Published"
        }

        Dim chkInPropDef As PropDef '= filePropDefs.DispName("File Size")
        For Each def As PropDef In filePropDefs
            If def.DispName = "Folder Path" Then
                chkInPropDef = def
                Exit For
            End If
        Next
        Dim chkIn As New SrchCond() With {
            .PropDefId = chkInPropDef.Id,
            .PropTyp = PropertySearchType.SingleProperty,
            .SrchOper = 3, 'Equals
            .SrchRule = SearchRuleType.Must,
            .SrchTxt = "1/24/2026"
        }

        Dim searchCond As SrchCond() = New AWS.SrchCond() {fileExtDXF, fileSize, filePath, chkIn}

        Dim bookmark As String = vbNull
        Dim status As SrchStatus '= Nothing
        Dim totalResults As New List(Of AWS.File)()
        Dim fileIters As New List(Of VDF.Vault.Currency.Entities.FileIteration)
        Dim resultIter As VDF.Vault.Currency.Entities.FileIteration

        Do

            Dim results As AWS.File() = mVltCon.WebServiceManager.DocumentService.FindFilesBySearchConditions(searchCond, Nothing, Nothing, False, True, bookmark, status)
            If results IsNot Nothing Then
                totalResults.AddRange(results)

            Else
                'Exit While
            End If

        Loop Until totalResults.Count = status.TotalHits

        'MsgBox(status.TotalHits)

        'Dim numFiles As Integer = 5
        'totalResults.RemoveRange(numFiles, totalResults.Count - numFiles)

        MsgBox("Found " & totalResults.Count & " big DXF files.")

        If totalResults.Count = 0 Then
            MsgBox("No files to process - Exiting")
            VaultLogout(mVltCon)
            Exit Sub
        End If

        For Each file As AWS.File In totalResults

            fileIters.Add(New VDF.Vault.Currency.Entities.FileIteration(mVltCon, file))

        Next

        DownloadFiles(fileIters, mVltCon, lstFilePaths)

        'Replace this with the call to fix DXF's
        DXFFix(lstFilePaths)

        'Check in Files ////////////////////////////////////////////////////

        VaultLogout(mVltCon)

    End Sub

    Public Sub DownloadFiles(fileIters As ICollection(Of VDF.Vault.Currency.Entities.FileIteration), mvltCon As VDF.Vault.Currency.Connections.Connection, ByVal filePaths As List(Of String))

        ' download individual files to a temp location

        Dim settings As New VDF.Vault.Settings.AcquireFilesSettings(mvltCon)

        settings.LocalPath = Nothing ' New VDF.Currency.FolderPathAbsolute("c:\_vaultWIP\Designs\_Part Drawings - Published")

        For Each fileIter As VDF.Vault.Currency.Entities.FileIteration In fileIters

            settings.AddFileToAcquire(fileIter, VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download Or VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Checkout)

        Next

        Dim results As VDF.Vault.Results.AcquireFilesResults = mvltCon.FileManager.AcquireFiles(settings)

        For Each result As VDF.Vault.Results.FileAcquisitionResult In results.FileResults

            filePaths.Add(result.LocalPath.FullPath)

        Next

        MsgBox("Downloaded and Checked Out " & filePaths.Count & " files")

    End Sub


    Function GetVaultFile(VaultPath As String, mVltCon As VDF.Vault.Currency.Connections.Connection) As VDF.Vault.Currency.Entities.FileIteration

        VaultPath = VaultPath.Replace("C:\_vaultWIP\", "$/")

        'flip the slashes
        VaultPath = VaultPath.Replace("\", "/")

        Dim VaultPaths() As String = New String() {VaultPath}

        Dim wsFiles() As AWS.File = mVltCon.WebServiceManager.DocumentService.FindLatestFilesByPaths(VaultPaths)

        GetVaultFile = New VDF.Vault.Currency.Entities.FileIteration(mVltCon, wsFiles(0))

    End Function 'GetVaultFile

    Public Sub DownloadAssembly(topLevelAssembly As VDF.Vault.Currency.Entities.FileIteration, m_conn As VDF.Vault.Currency.Connections.Connection)

        ' download the latest version of the assembly to working folders
        Dim settings As New VDF.Vault.Settings.AcquireFilesSettings(m_conn)
        settings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeChildren = True
        settings.OptionsRelationshipGathering.FileRelationshipSettings.RecurseChildren = True
        settings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeParents = False
        settings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeRelatedDocumentation = True
        settings.OptionsRelationshipGathering.FileRelationshipSettings.VersionGatheringOption = VDF.Vault.Currency.VersionGatheringOption.Latest
        settings.AddFileToAcquire(topLevelAssembly, VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download)
        Dim results As VDF.Vault.Results.AcquireFilesResults

        Try
            If settings.IsValidConfiguration = True Then
                results = m_conn.FileManager.AcquireFiles(settings)
                MsgBox("Files Acquired")
            Else
                MsgBox("Settings Invalid")
            End If
        Catch ex As Exception
            MsgBox("Error during download: " & ex.Message & vbCr & "Please ensure this file has been checked into Vault before running this rule.")
        End Try

    End Sub 'DownloadAssembly

    Public Function VaultLogin(serverName As String, vaultName As String) As VDF.Vault.Currency.Connections.Connection

        Dim connection As VDF.Vault.Currency.Connections.Connection

        VFM.Library.Initialize()

        Try
            Dim loginSettings As New VFM.Settings.LoginSettings '= VFM.Settings.LoginSetting

            connection = VFM.Library.Login(loginSettings)

            If connection IsNot Nothing Then
                'MsgBox("Logged in to Vault: " & vaultName)
                Return connection
            End If
        Catch ex As Exception

        End Try

    End Function


    Public Sub VaultLogout(ByRef vltCon As VDF.Vault.Currency.Connections.Connection)
        VFM.Library.Logout(vltCon)
    End Sub

    Public Sub CheckInFileIters(fileIters As ICollection(Of VDF.Vault.Currency.Entities.FileIteration), mVltCon As VDF.Vault.Currency.Connections.Connection)
        For Each fileIter As VDF.Vault.Currency.Entities.FileIteration In fileIters
            'MsgBox("Checking in: " & fileIter.ServerPath.FullPath)
            'VDF.Vault.Services.Connection.IFileManager.CheckinFile(fileIter, "Checked in by Vault Fix Tool", False, VDF.Vault.Settings.CheckInFilesSettings.CheckInOptions.None,
        Next
    End Sub

    Public Sub FixVault(serverName As String, vaultName As String, Optional ByVal lstFilePaths As List(Of String) = Nothing)

        If lstFilePaths Is Nothing Then
            lstFilePaths = New List(Of String)
        End If
        Dim mVltCon As VDF.Vault.Currency.Connections.Connection = VaultLogin(serverName, vaultName)

        If mVltCon Is Nothing Then
            MsgBox("Vault Login Failed - Exiting")
            Exit Sub
        End If

        Dim files As AWS.File

        ' Dim filePropDefs As New AWS.PropDef() = AWS.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE")
        Dim filePropDefs As PropDef() = mVltCon.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE")

        Dim chkInPropDef As PropDef
        For Each def As PropDef In filePropDefs
            If def.DispName = "Checked In" Then
                chkInPropDef = def
                Exit For
            End If
        Next
        Dim chkIn As New SrchCond() With {
            .PropDefId = chkInPropDef.Id,
            .PropTyp = PropertySearchType.SingleProperty,
            .SrchOper = 6, 'Greater Than
            .SrchRule = SearchRuleType.Must,
            .SrchTxt = "11/18/2025"
        }

        Dim searchCond As SrchCond() = New AWS.SrchCond() {chkIn}

        Dim bookmark As String = vbNull
        Dim status As SrchStatus '= Nothing
        Dim totalResults As New List(Of AWS.File)()
        Dim fileIters As New List(Of VDF.Vault.Currency.Entities.FileIteration)
        Dim resultIter As VDF.Vault.Currency.Entities.FileIteration

        Do
            Dim results As AWS.File() = mVltCon.WebServiceManager.DocumentService.FindFilesBySearchConditions(searchCond, Nothing, Nothing, False, True, bookmark, status)
            If results IsNot Nothing Then
                totalResults.AddRange(results)
            Else
                'Exit While
            End If

        Loop Until totalResults.Count = status.TotalHits

        'MsgBox("Found " & status.TotalHits & " files. ")

        'Dim numFiles As Integer = 5
        'totalResults.RemoveRange(numFiles, totalResults.Count - numFiles)

        MsgBox("Found " & totalResults.Count & " delta files.")

        For Each file As AWS.File In totalResults

            fileIters.Add(New VDF.Vault.Currency.Entities.FileIteration(mVltCon, file))

        Next

        'Download Files and get file list////////////////////////////////////////////////////
        'DownloadFiles(fileIters, mVltCon, lstFilePaths)

        'Logout old Vault //////////////////////////////////////////////////
        ' VaultLogout(mVltCon)

        'Login New Vault //////////////////////////////////////////////////
        ' VaultLogin("https://vaultsandbox.tait.rocks/", vaultName)

        'Check out Existing Files ////////////////////////////////////////////////////

        'Check in 
        'Check in Files ////////////////////////////////////////////////////

        VaultLogout(mVltCon)

    End Sub
End Module
