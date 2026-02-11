
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
Imports Autodesk.DataManagement.Client.Framework.Vault.Currency.Entities

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
            If def.DispName = "Checked In" Then
                chkInPropDef = def
                Exit For
            End If
        Next
        Dim chkIn As New SrchCond() With {
            .PropDefId = chkInPropDef.Id,
            .PropTyp = PropertySearchType.SingleProperty,
            .SrchOper = 6, 'Greater than
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

        MsgBox("Found " & totalResults.Count & " big DXF files.")

        If totalResults.Count = 0 Then
            MsgBox("No files to process - Exiting")
            VaultLogout(mVltCon)
            Exit Sub
        End If

        For Each file As AWS.File In totalResults

            fileIters.Add(New VDF.Vault.Currency.Entities.FileIteration(mVltCon, file))

        Next


        Form1.ProgressBar1.Visible = True


        DownloadChkOutFiles(fileIters, mVltCon, lstFilePaths)


        DXFFix(lstFilePaths)

        Form1.ProgressBar1.Visible = False

        'Check in Files ////////////////////////////////////////////////////

        VaultLogout(mVltCon)

    End Sub

    Public Sub DownloadFiles(fileIters As ICollection(Of VDF.Vault.Currency.Entities.FileIteration), mvltCon As VDF.Vault.Currency.Connections.Connection, ByVal filePaths As List(Of String))

        ' download individual files to a temp location

        Dim settings As New VDF.Vault.Settings.AcquireFilesSettings(mvltCon)

        settings.LocalPath = Nothing ' New VDF.Currency.FolderPathAbsolute("c:\_vaultWIP\Designs\_Part Drawings - Published")

        For Each fileIter As VDF.Vault.Currency.Entities.FileIteration In fileIters

            settings.AddFileToAcquire(fileIter, VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download)

        Next

        Dim results As VDF.Vault.Results.AcquireFilesResults = mvltCon.FileManager.AcquireFiles(settings)

        For Each result As VDF.Vault.Results.FileAcquisitionResult In results.FileResults

            filePaths.Add(result.LocalPath.FullPath)

        Next

        MsgBox("Downloaded and Checked Out " & filePaths.Count & " files")

    End Sub

    Public Sub DownloadChkOutFiles(fileIters As ICollection(Of VDF.Vault.Currency.Entities.FileIteration), mvltCon As VDF.Vault.Currency.Connections.Connection, ByVal filePaths As List(Of String))

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

    Public Sub CheckoutFiles(fileIters As ICollection(Of VDF.Vault.Currency.Entities.FileIteration), mvltCon As VDF.Vault.Currency.Connections.Connection, ByVal filePaths As List(Of String))

        ' download individual files to a temp location

        Dim settings As New VDF.Vault.Settings.AcquireFilesSettings(mvltCon)

        settings.LocalPath = Nothing ' New VDF.Currency.FolderPathAbsolute("c:\_vaultWIP\Designs\_Part Drawings - Published")

        For Each fileIter As VDF.Vault.Currency.Entities.FileIteration In fileIters

            settings.AddFileToAcquire(fileIter, VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Checkout)

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

    Public Sub FixVault(serverName As String, vaultName As String, sDate As String, Optional ByVal lstFilePaths As List(Of String) = Nothing )

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

        Dim chkIn As New SrchCond()

        chkIn = MakeSrchCond("Checked In", 6, sDate, filePropDefs)

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

        MsgBox("Found " & totalResults.Count & " delta files.")

        For Each file As AWS.File In totalResults

            fileIters.Add(New VDF.Vault.Currency.Entities.FileIteration(mVltCon, file))

        Next

        'Download Files and get file list////////////////////////////////////////////////////
        'CheckoutFiles(fileIters, mVltCon, lstFilePaths) 'Working, disabled for testing

        'Logout old Vault //////////////////////////////////////////////////
        ' VaultLogout(mVltCon) 'Working, disabled for testing

        'Login New Vault //////////////////////////////////////////////////
        ' VaultLogin("https://vaultsandbox.tait.rocks/", vaultName) 'Working, disabled for testing

        'Check out Existing Files ////////////////////////////////////////////////////
        ' iterate through fileIters to check out all files
        ' ??? do the file iterations work across Vault? Might need to get new file iterations from files? Do Files work across vaults?  Need to do testing
        ' ??? How to account for files that were moved or renamed? Search and move in WIP before checkin?

        ' CheckoutFiles(fileIters, mVltCon, lstFilePaths) 'check out files in sandbox, do not download

        'Check in 
        ' Check in Files ////////////////////////////////////////////////////
        '   Sort Inventor Files and Other Files
        ' check in Other files directly

        VaultLogout(mVltCon)

    End Sub

    Sub SortFiles(ByVal fileIters As VDF.Vault.Currency.Entities.FileIteration)

    End Sub 'SortFiles

    Public Function MakeSrchCond(srchTxt As String, oper As String, propName As String, filePropDefs As PropDef()) As SrchCond

        'Find the Property definition by name
        Dim propDef As PropDef
        For Each def As PropDef In filePropDefs
            If def.DispName = propName Then
                propDef = def
                Exit For
            End If
        Next

        'Define the search condidion for a single property
        Dim srchCond As New SrchCond() With {
            .PropDefId = propDef.Id,
            .PropTyp = PropertySearchType.SingleProperty,
            .SrchOper = oper,
            .SrchRule = SearchRuleType.Must,
            .SrchTxt = srchTxt
        }

        Return srchCond

    End Function

    Public Sub FixVaultTest(serverName As String, vaultName As String, fileName As String, Optional ByVal lstFilePaths As List(Of String) = Nothing)

        If lstFilePaths Is Nothing Then
            lstFilePaths = New List(Of String)
        End If

        Dim mVltCon As VDF.Vault.Currency.Connections.Connection = VaultLogin(serverName, vaultName)

        If mVltCon Is Nothing Then
            MsgBox("Vault Login Failed - Exiting")
            Exit Sub
        End If

        Dim filePropDefs As PropDef() = mVltCon.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE")

        Dim deltaDate As New SrchCond()
        deltaDate = MakeSrchCond("Standard.ipt", "3", "File Name", filePropDefs)

        Dim searchCond As SrchCond() = New AWS.SrchCond() {deltaDate}
        Dim bookmark As String = vbNull
        Dim status As SrchStatus '= Nothing
        Dim totalResults As New List(Of AWS.File)()
        Dim fileIters As New List(Of VDF.Vault.Currency.Entities.FileIteration)
        'Dim resultIter As VDF.Vault.Currency.Entities.FileIteration
        Dim invResults As New List(Of AWS.File)

        Do
            Dim results As AWS.File() = mVltCon.WebServiceManager.DocumentService.FindFilesBySearchConditions(searchCond, Nothing, Nothing, False, True, bookmark, status)
            If results IsNot Nothing Then
                totalResults.AddRange(results)
            End If
        Loop Until totalResults.Count = status.TotalHits

        MsgBox("Found " & totalResults.Count & " delta files.")



        For Each file As AWS.File In totalResults

            fileIters.Add(New VDF.Vault.Currency.Entities.FileIteration(mVltCon, file))

        Next


        'Download all modified delta Files And get file list as strings////////////////////////////////////////////////////

        'DownloadFiles(fileIters, mVltCon, lstFilePaths) 'Working, disabled for testing



        'Logout old Vault //////////////////////////////////////////////////
        'VaultLogout(mVltCon) 'Working, disabled for testing

        'Login New Vault //////////////////////////////////////////////////
        'mVltCon = VaultLogin("https://vaultsandbox.tait.rocks/", vaultName) 'Working, disabled for testing

        'Check out Existing Files ////////////////////////////////////////////////////
        ' Extract file name without path for each file and add to DataTable: origPath, fileName, newPath
        '   orig path from lstFilePaths (gathered when files downloaded)
        '   file name extracted with extension from origPath
        '   new path obtained from folderID on file objec, using 


        Dim dtFileInfo As New DataTable()
        dtFileInfo.Columns.Add("origPath", GetType(String))
        dtFileInfo.Columns.Add("fileName", GetType(String))
        dtFileInfo.Columns.Add("newPath", GetType(String))
        dtFileInfo.Columns.Add("fileObject", GetType(AWS.File))
        dtFileInfo.Columns.Add("fileIter", GetType(VDF.Vault.Currency.Entities.FileIteration))

        Dim dtErrorTable As New DataTable()
        dtErrorTable.Columns.Add("origPath", GetType(String))
        dtErrorTable.Columns.Add("fileName", GetType(String))
        dtErrorTable.Columns.Add("newPath", GetType(String))
        dtErrorTable.Columns.Add("fileObject", GetType(AWS.File))
        dtErrorTable.Columns.Add("fileIter", GetType(VDF.Vault.Currency.Entities.FileIteration))

        Dim sFileName As String

        For Each fPath As String In lstFilePaths
            sFileName = Right(fPath, InStrRev(fPath, "/"))

            dtFileInfo.Rows.Add(fPath, sFileName, "")
        Next

        ' Search for each fileName, get AWS.File Object, Get FileIteration for File, write out exceptions or failed searches.  
        ' iterate through fileIters to check out all files

        Dim fName As New SrchCond()
        Dim srchResults As New List(Of AWS.File)()
        Dim srchFiles As New List(Of AWS.File)()

        Dim invFiles As New List(Of String)
        Dim otherFiles As New List(Of String)

        For Each row As DataRow In dtFileInfo.Rows

            sFileName = row("fileName").ToString

            fName = MakeSrchCond(sFileName, "3", "File Name", filePropDefs)
            searchCond = New AWS.SrchCond() {fName}
            bookmark = vbNull

            'Execute search and store results, 100 items at a time, loop until all are stored
            Do
                Dim results As AWS.File() = mVltCon.WebServiceManager.DocumentService.FindFilesBySearchConditions(searchCond, Nothing, Nothing, False, True, bookmark, status)
                If results IsNot Nothing Then
                    srchResults.AddRange(results)
                Else
                    'If results are nothing, its likely the file has been deleted or otherwise restricted beyond user permissions, write file info to text file for manual resolution
                    ' this shouldnt happen as admin, but just in case.  
                    '  "No Results for : " & row("origPath").ToString & " - " & row("fileName").ToString
                    dtErrorTable.Rows.Add(row)
                    Continue For
                End If
            Loop Until srchResults.Count = status.TotalHits

            'If srchResults > 1, write info to text for manual resolution
            If srchResults.Count > 1 Then
                'Write info to text file
                '  "Multiple Search Results for : " & row("origPath").ToString & " - & row("fileName").ToString & " - " & row("newPath").ToString
                'Write row to dtErrorTable
                dtErrorTable.Rows.Add(row)
                Continue For
            Else
                row("fileObject") = srchResults(0)
            End If

            'Get the folder object from document service, to extract new vault path
            Dim fldr As AWS.Folder = mVltCon.WebServiceManager.DocumentService.GetFolderById(srchResults(0).FolderId)
            row("newPath") = fldr.FullName


            'Begin Error Checking //////////////////////////////////////////////////////////////////

            'Check if original file path matches new file path
            If Not row("origPath").ToString = row("newPath").ToString & "/" & sFileName Then
                'export file info to text file for manual resolution
                'Write row to dtErrorTable
                dtErrorTable.Rows.Add(row)
                Continue For
            End If

            'check if files are locked
            If srchResults(0).Locked Then
                'export locked filenames to text file for manual resolution
                'Write row to dtErrorTable
                dtErrorTable.Rows.Add(row)
                Continue For
            End If


            'get fileIterations of files to prepare for checkout
            row("fileIter") = New VDF.Vault.Currency.Entities.FileIteration(mVltCon, srchResults(0))

            'Sort out Inventor files for checking in by application
            If sFileName.Contains(".ipt") OrElse sFileName.Contains(".iam") OrElse sFileName.Contains(".ipn") OrElse sFileName.Contains(".idw") OrElse sFileName.Contains(".dwg") Then
                invFiles.Add(sFileName)
            Else
                otherFiles.Add(sFileName)
            End If



        Next

        'delete rows in dtFileInfo if they exits in dtErrorTable
        For Each errRow As DataRow In dtErrorTable.Rows
            For Each row As DataRow In dtFileInfo.Rows
                If errRow("partName").ToString = row("partName").ToString Then
                    row.Delete()
                End If
            Next
        Next





        ' CheckoutFiles(fileIters, mVltCon, lstFilePaths) 'check out files in sandbox, do not download

        'Check in 
        ' Check in Files ////////////////////////////////////////////////////
        '   Sort Inventor Files and Other Files
        ' check in Other files directly

        VaultLogout(mVltCon)

    End Sub
End Module
