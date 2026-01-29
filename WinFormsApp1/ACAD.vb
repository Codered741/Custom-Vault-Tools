Imports System.Threading
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports System.IO
Imports System.Runtime.CompilerServices


Module Program

    Sub DXFFix(Optional oFiles As List(Of String) = Nothing)
        Dim acadApp As AcadApplication
        Try
            acadApp = GetObject(, "AutoCAD.Application")   ' attach if already running
            Thread.Sleep(5000)
            'MsgBox(acadApp.FullName)
        Catch ex As Exception
            ' Not running → start it manually
            Try
                acadApp = CreateObject("AutoCAD.Application")
                Thread.Sleep(5000)
            Catch createEx As Exception
                MessageBox.Show("Cannot start or find AutoCAD: " & createEx.Message)
                Return
            End Try
        End Try
        'Dim acadApp As AcadApplication
        'acadApp = CreateObject("AutoCAD.Application")

        If acadApp Is Nothing Then
            MessageBox.Show("AutoCAD COM object not available. Verify installation/ProgID/bitness.")
            Return
        End If

        acadApp.Visible = True

        If oFiles Is Nothing Then

            oFiles = New List(Of String)

            'Configure dialog for file selection
            Dim oFileDlg As New OpenFileDialog With {
                .Filter = "DXF Files (*.dxf)|*.dxf", '|All Files (*.*)|*.*"
                .Title = "Select File(s) to Import",
                .InitialDirectory = "C:\_vaultWIP\Designs\_Part Drawings - Published\",
                .Multiselect = (True)
            }

            'Open Dialog
            oFileDlg.ShowDialog()

            If oFileDlg.FileName <> "" Then

                For Each file In oFileDlg.FileNames
                    oFiles.Add(file)
                Next

            Else
                MessageBox.Show("No files were selected")
                Return
            End If

        End If

        Dim BlockName As New List(Of String) ' = "Borders TAIT MDE - PART_SUB-ASSY - BORDER"
        Dim dxfDoc As AcadDocument
        Dim blockTable As AcadBlocks
        Dim blockDef As AcadBlock
        Dim dxfMS As Object
        Dim sset As AcadSelectionSet = Nothing
        Dim saveName As String 'edited save name
        Dim fileName As String 'parsed file name
        Dim filePath As String 'parsed file path
        Dim fixedFiles As New List(Of String)

        For Each dxf In oFiles
            'MsgBox(dxf)
            Thread.Sleep(1000)
            dxfDoc = acadApp.Documents.Open(dxf)
            Thread.Sleep(2000)

            blockTable = dxfDoc.Blocks
            BlockName.Clear()

            For Each blockDef In blockTable
                If blockDef.Name Like "*Blocks*" Then
                    BlockName.Add(blockDef.Name)
                    'Exit For
                End If

                If blockDef.Name Like "*Border*" Then
                    BlockName.Add(blockDef.Name)
                End If
            Next

            If BlockName.Count = 0 Then
                GoTo skipFile
            End If

            For Each block As String In BlockName
                dxfDoc.SendCommand("-bedit" & vbCr & block & vbCr)
                Thread.Sleep(1000)


                dxfMS = dxfDoc.ModelSpace
                For Each entry In dxfMS
                    If TypeOf entry Is AcadOle Then
                        entry.Delete
                    End If
                Next

                'dxfDoc.SendCommand("select" & vbCr & "window" & vbCr & "405,275" & vbCr & "395,285" & vbCr)
                dxfDoc.SendCommand("bsave" & vbCr)
                dxfDoc.SendCommand("bclose" & vbCr)
                Thread.Sleep(250)
            Next

            dxfDoc.SendCommand("-purge" & vbCr & "a" & vbCr & vbCr & "n" & vbCr)

            Thread.Sleep(500)

            fileName = System.IO.Path.GetFileName(dxf)
            fileName = Left(fileName, fileName.Length - 4)


            filePath = Left(dxf, dxf.Length - (fileName.Length + 4))

            saveName = filePath & "Fixed\" & fileName

            fixedFiles.Add(saveName & ".dxf")

            dxfDoc.Export(saveName, "dxf", sset)
            'needs to export to separate folder, parse file string for file path

skipFile:
            dxfDoc.Close(False)
NextFile:
        Next

        Thread.Sleep(1000)
        acadApp.Quit()

        For i = 0 To fixedFiles.Count - 1
            If File.Exists(oFiles(i)) Then
                File.Delete(oFiles(i))
                Try
                    File.Move(fixedFiles(i), oFiles(i))
                    File.Delete(fixedFiles(i))
                Catch ex As Exception
                    MsgBox("Could not move fixed file " & fixedFiles(i) & " to " & oFiles(i) & vbCr & ex.Message)
                End Try
            Else
                MsgBox("Original file not found for replacement: " & oFiles(i))
            End If
        Next

        MsgBox("Done!")
    End Sub

    Sub GetAcadApp()
        Dim acadApp As AcadApplication
        Try
            acadApp = GetObject(, "AutoCAD.Application")   ' attach if already running
            Thread.Sleep(5000)
            MsgBox(acadApp.FullName)
        Catch ex As Exception
            ' Not running → start it manually
            Try
                acadApp = CreateObject("AutoCAD.Application")
            Catch createEx As Exception
                MessageBox.Show("Cannot start or find AutoCAD: " & createEx.Message)
                Return
            End Try
        End Try

        acadApp.Visible = True

    End Sub
End Module
