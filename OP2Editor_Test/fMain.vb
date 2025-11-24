Imports System.Diagnostics
Imports System.IO
Imports System.Media
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports Interop.OP2Editor
Imports Microsoft.Win32

' OP2EditorTester
' https://github.com/leviathan400/OP2EditorTester
'
' OP2Editor test application. Simple app to work with the OP2Editor for testing.
'
' Github for OP2Editor: https://github.com/OutpostUniverse/OP2Editor
'
' Outpost 2: Divided Destiny is a real-time strategy video game released in 1997.

Public Class fMain

    Public Version As String = "0.4.0"

    ' Outpost 2 - 1.3.6 folder
    Private OP2FolderPath As String = "C:\Users\User\Desktop\OP2-136"

    Private rm As IResourceManager
    Private map As MapFile

    Private Sub AppendToConsole(text As String)
        ' Append text to the console with a newline
        txtConsole.AppendText(text & vbCrLf)

        ' Scroll to the end of the text
        txtConsole.SelectionStart = txtConsole.Text.Length
        txtConsole.ScrollToCaret()
    End Sub

    Public Function FindRegisteredComDll(target As String) As List(Of String)
        Dim results As New List(Of String)()

        ' Search HKCR\CLSID and HKCR\WOW6432Node\CLSID
        Dim roots = {
        "HKEY_CLASSES_ROOT\CLSID",
        "HKEY_CLASSES_ROOT\WOW6432Node\CLSID"
    }

        For Each root In roots
            Using clsidKey As RegistryKey = Registry.ClassesRoot.OpenSubKey(If(root.Contains("WOW"), "WOW6432Node\CLSID", "CLSID"))
                If clsidKey Is Nothing Then Continue For

                For Each subKeyName In clsidKey.GetSubKeyNames()
                    Using subKey = clsidKey.OpenSubKey(subKeyName & "\InprocServer32")
                        If subKey Is Nothing Then Continue For

                        Dim dllPath = TryCast(subKey.GetValue(""), String)
                        If String.IsNullOrEmpty(dllPath) Then Continue For

                        ' Match by filename or path
                        If dllPath.IndexOf(target, StringComparison.OrdinalIgnoreCase) >= 0 Then
                            results.Add(dllPath)
                        End If
                    End Using
                Next
            End Using
        Next

        Return results.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
    End Function

    Public Function GetDllVersion(dllPath As String) As String
        Try
            Dim info = FileVersionInfo.GetVersionInfo(dllPath)

            Return $"FileVersion={info.FileVersion}, " &
               $"ProductVersion={info.ProductVersion}, " &
               $"Description={info.FileDescription}"
        Catch ex As Exception
            Return "(version info unavailable)"
        End Try
    End Function

    Private Sub fMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Debug.WriteLine("OP2EditorTester")
        Me.Icon = My.Resources.Mapper2_Icon

        TimerStartup.Enabled = True
    End Sub

    Private Sub TimerStartup_Tick(sender As Object, e As EventArgs) Handles TimerStartup.Tick
        FindCOMRegistration("OP2Editor")

        TimerStartup.Enabled = False
        btnLoadMap.Enabled = True
        btnExtractCLM.Enabled = True

        Dim player As New SoundPlayer(My.Resources.beep2)
        player.Play()
    End Sub

    Private Sub FindCOMRegistration(ByVal ComName As String)
        Dim paths = FindRegisteredComDll(ComName)

        If paths.Count = 0 Then
            AppendToConsole(ComName & " not registered.")
        Else
            AppendToConsole("Found " & ComName & " registration:")
            For Each p In paths
                AppendToConsole("" & p)
                AppendToConsole(GetDllVersion(p))
            Next
        End If
    End Sub

    Private Sub btnLoadMap_Click(sender As Object, e As EventArgs) Handles btnLoadMap.Click
        AppendToConsole("")

        Try
            Dim OpenFileDialog As New OpenFileDialog With {
            .Title = "Select Outpost 2 .MAP file",
            .Filter = "Outpost 2 Map Files (*.map)|*.map|All Files (*.*)|*.*",
            .InitialDirectory = OP2FolderPath
            }

            If OpenFileDialog.ShowDialog() <> DialogResult.OK Then
                AppendToConsole("Map load cancelled.")
                Exit Sub
            End If

            Dim selectedMapPath As String = OpenFileDialog.FileName
            AppendToConsole("Selected map: " & selectedMapPath)

            rm = New ResourceManagerClass()

            rm.RootPath = OP2FolderPath

            ' Add map .VOL archives
            AddVol(rm, IO.Path.Combine(OP2FolderPath, "maps.vol"))
            AddVol(rm, IO.Path.Combine(OP2FolderPath, "maps01.vol"))
            AddVol(rm, IO.Path.Combine(OP2FolderPath, "maps02.vol"))
            AddVol(rm, IO.Path.Combine(OP2FolderPath, "maps03.vol"))
            AddVol(rm, IO.Path.Combine(OP2FolderPath, "maps04.vol"))

            ' Load a map file
            Dim flags As MapLoadSaveFormat = CType(0, MapLoadSaveFormat)
            map = rm.LoadMapFile("C:\Users\User\Desktop\OP2-136\map1.map", flags)

            AppendToConsole("Loaded .map successfully.")

            InspectMap(map)
            LoadTileSet_Log(map)

            Marshal.FinalReleaseComObject(map)

        Catch ex As Exception
            AppendToConsole("Error: " & ex.ToString)
            AppendToConsole("Error: " & ex.Message)

            'Catch ex As COMException
            'MessageBox.Show($"COM 0x{ex.ErrorCode:X8}: {ex.Message}")
        Finally
            If rm IsNot Nothing Then Marshal.FinalReleaseComObject(rm) : rm = Nothing
        End Try
    End Sub

    Private Sub AddVol(rm As IResourceManager, path As String)
        If IO.File.Exists(path) Then
            Dim vol = rm.LoadVolFile(path, True)
            rm.AddArchiveToSearch(vol)
            Marshal.FinalReleaseComObject(vol)
            ''AppendToConsole("AddVol: " & path)
        End If
    End Sub

    Private Sub InspectMap(map As MapFile)
        Try
            ' Map info
            Dim w As Integer = map.TileWidth
            Dim h As Integer = map.TileHeight
            Dim atw As Boolean = (map.AroundTheWorld <> 0)
            AppendToConsole($"Size: {w}×{h}, AroundTheWorld={atw}")

            ' count cell types & sample a few tiles
            Dim counts As New Dictionary(Of Integer, Integer)
            For y = 0 To h - 1
                For x = 0 To w - 1
                    Dim ct As Integer = map.CellType(x, y)       ' raw cell-type value
                    If Not counts.ContainsKey(ct) Then counts(ct) = 0
                    counts(ct) += 1
                Next
            Next
            AppendToConsole("CellType histogram:")
            For Each kv In counts.OrderBy(Function(k) k.Key)
                AppendToConsole($"  {kv.Key} -> {kv.Value}")
            Next

            ' read tile indices (terrain graphic index)
            Dim tile00 As Integer = map.TileData(0, 0)
            Dim tile01 As Integer = map.TileData(0, 1)
            AppendToConsole($"TileData(0,0)={tile00}, TileData(0,1)={tile01}")

            ' Cell Type
            Dim cellTypeTest As Integer = map.CellType(0, 0)
            AppendToConsole($"CellType(0,0)={cellTypeTest}")

            ' groups (named tile selections)
            Dim nGroups As Integer = map.NumTileGroups
            AppendToConsole($"TileGroups: {nGroups}")
            For i = 0 To nGroups - 1
                AppendToConsole($"  {i}: {map.TileGroupName(i)}")
            Next

        Finally
            ' nothing here; caller should release map
        End Try
    End Sub

    Private Sub LoadTileSet_Log(map As MapFile)
        AppendToConsole("Tilesets in .map:")

        For Each t In GetTileSets(map)
            AppendToConsole($"{t.Name}  ({t.Count} tiles)")
        Next
    End Sub

    Function GetTileSets(map As MapFile) As List(Of (Name As String, Count As Integer))
        Dim list As New List(Of (String, Integer))()
        Dim tsm As TileSetManager = Nothing

        Try
            tsm = map.TileSetManager
            Dim n As Integer = tsm.NumTileSets

            For i = 0 To n - 1
                Dim ts As TileSet = Nothing
                Try
                    ' Some slots may be empty in this COM API – guard for Nothing
                    ts = tsm.TileSet(i)
                    If ts Is Nothing Then
                        ' Either a sparse slot or end-of-list – VB6 exited here
                        Continue For  ' or Exit For if you prefer to stop at first Nothing
                    End If

                    Dim name As String = tsm.TileSetName(i)
                    If String.IsNullOrEmpty(name) Then name = $"TileSet {i}"

                    Dim count As Integer = ts.NumTiles
                    list.Add((name, count))

                Finally
                    If ts IsNot Nothing Then Marshal.FinalReleaseComObject(ts)
                End Try
            Next

        Finally
            If tsm IsNot Nothing Then Marshal.FinalReleaseComObject(tsm)
        End Try

        Return list
    End Function

    Private Sub btnExtractCLM_Click(sender As Object, e As EventArgs) Handles btnExtractCLM.Click
        'Testing...

        rm = New ResourceManagerClass()

        ' Open file dialog to select the CLM file
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Title = "Select CLM File"
        openFileDialog.Filter = "CLM Files (*.clm)|*.clm|All Files (*.*)|*.*"
        openFileDialog.FileName = "op2.clm"

        If openFileDialog.ShowDialog() <> DialogResult.OK Then
            Return ' User cancelled
        End If

        Dim clmFilePath As String = openFileDialog.FileName
        AppendToConsole($"Selected CLM file: {clmFilePath}")

        ' Check if file exists
        If Not System.IO.File.Exists(clmFilePath) Then
            AppendToConsole("ERROR: File does not exist!")
            MessageBox.Show("File does not exist.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        AppendToConsole($"File exists. Size: {New System.IO.FileInfo(clmFilePath).Length} bytes")

        ' Folder browser dialog to select extraction destination
        Dim folderBrowser As New FolderBrowserDialog()
        folderBrowser.Description = "Select folder to extract CLM files to"
        folderBrowser.ShowNewFolderButton = True

        If folderBrowser.ShowDialog() <> DialogResult.OK Then
            Return ' User cancelled
        End If

        Dim extractPath As String = folderBrowser.SelectedPath
        AppendToConsole($"Extract to: {extractPath}")

        Try
            ' Load the CLM file
            AppendToConsole("Attempting to load CLM file...")

            Dim clmReader As ArchiveReader = Nothing

            Try
                clmReader = rm.LoadClmFile(clmFilePath, 0)
                AppendToConsole("LoadClmFile completed")
            Catch ex As Exception
                AppendToConsole($"ERROR in LoadClmFile: {ex.Message}")
                AppendToConsole($"Stack trace: {ex.StackTrace}")
                AppendToConsole($"Failed to load CLM file:{vbCrLf}{ex.Message}")
                Return
            End Try

            If clmReader Is Nothing Then
                AppendToConsole("ERROR: clmReader is Nothing!")
                AppendToConsole("Failed to load CLM file. The reader returned Nothing.")
                Return
            End If

            AppendToConsole("CLM file loaded successfully")

            ' Get number of files in the archive
            Dim numFiles As Long = clmReader.NumFiles
            AppendToConsole($"Number of files in archive: {numFiles}")

            If numFiles = 0 Then
                AppendToConsole("No files found in CLM archive.")
                Return
            End If

            ' Extract each file
            Dim extractedCount As Integer = 0
            For i As Long = 0 To numFiles - 1
                ' Get the file name
                Dim fileName As String = clmReader.fileName(i)
                AppendToConsole($"Extracting [{i + 1}/{numFiles}]: {fileName}")

                Try
                    ' Open stream to read the file from archive
                    Dim inStream As Interop.OP2Editor.StreamReader = clmReader.OpenStreamRead(fileName)

                    If inStream Is Nothing Then
                        AppendToConsole($"  - Failed to open read stream for: {fileName}")
                        Continue For
                    End If

                    ' Create output file path
                    Dim outputPath As String = System.IO.Path.Combine(extractPath, fileName)
                    'AppendToConsole($"  Output path: {outputPath}")

                    ' Open stream to write the file
                    Dim outStream As Interop.OP2Editor.StreamWriter = Nothing
                    Try
                        outStream = rm.OpenStreamWrite(outputPath)
                    Catch ex As Exception
                        AppendToConsole($"  - Failed to create output stream: {ex.Message}")
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(inStream)
                        Continue For
                    End Try

                    If outStream Is Nothing Then
                        AppendToConsole($"  - Output stream is Nothing for: {fileName}")
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(inStream)
                        Continue For
                    End If

                    ' Copy data from archive to file
                    Dim bufferSize As Integer = 4096
                    Dim buffer(bufferSize - 1) As Byte
                    Dim bytesRead As Integer = 0
                    Dim bytesWritten As Integer = 0
                    Dim totalBytesRead As Long = 0
                    Dim totalBytesWritten As Long = 0

                    ' Pin the buffer and get pointer
                    Dim handle As System.Runtime.InteropServices.GCHandle =
                    System.Runtime.InteropServices.GCHandle.Alloc(buffer,
                    System.Runtime.InteropServices.GCHandleType.Pinned)
                    Dim pBuffer As Integer = handle.AddrOfPinnedObject().ToInt32()

                    Try

                        ' Extract...
                        ' I had issues..

                        extractedCount += 1
                        AppendToConsole($"  - {fileName} ({totalBytesWritten} bytes)")

                    Catch ex As Exception
                        AppendToConsole($"  - {fileName}: {ex.Message}")
                    Finally
                        ' Free the pinned buffer
                        handle.Free()
                    End Try

                    ' Release COM objects
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(outStream)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(inStream)

                Catch ex As Exception
                    AppendToConsole($"  - Error extracting {fileName}: {ex.Message}")
                End Try
            Next

            ' Release the CLM reader
            System.Runtime.InteropServices.Marshal.ReleaseComObject(clmReader)

            ' Show success message
            AppendToConsole($"Extraction complete: {extractedCount}/{numFiles} files extracted")
            'MessageBox.Show($"Successfully extracted {extractedCount} of {numFiles} files to:{vbCrLf}{extractPath}",
            '               "Extraction Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            AppendToConsole($"EXCEPTION: {ex.Message}")
            AppendToConsole($"Stack trace: {ex.StackTrace}")
        End Try
    End Sub

End Class
