
Imports System.Diagnostics
Imports System.IO
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports System.Drawing
Imports Newtonsoft.Json
Imports System.Linq
Imports Guna.UI2.WinForms
Imports Guna.UI2.WinForms.Enums
Imports System.Security.Principal

Public Class Form1
    Inherits Form

    Private WithEvents cmdProcess As Process
    Private isProcessRunning As Boolean = False
    Private batFilePath As String = ""
    Private Const SETTINGS_FILE As String = "bat_settings.json"
    Private batFilesList As New List(Of BatFileInfo)

    ' Guna UI Controls
    Private txtBatPath As Guna2TextBox
    Private cmbBatFiles As Guna2ComboBox
    Private btnBrowse As Guna2Button
    Private btnSave As Guna2Button
    Private btnDelete As Guna2Button
    Private btnStart As Guna2Button
    Private btnStop As Guna2Button
    Private btnRestart As Guna2Button
    Private lblStatus As Guna2HtmlLabel
    Private txtCommand As Guna2TextBox
    Private btnSend As Guna2Button
    Private txtOutput As RichTextBox
    Private btnClear As Guna2Button
    Private btnMinimize As Guna2Button
    Private btnClose As Guna2Button

    ' Panels and Layout
    Private topPanel As Guna2Panel
    Private middlePanel As Guna2Panel
    Private outputPanel As Guna2Panel
    Private mainContainer As TableLayoutPanel
    Private titlePanel As Guna2Panel
    Private lblTitle As Guna2HtmlLabel
    Private lblAdmin As Guna2HtmlLabel

    ' Theme Colors
    Private ReadOnly primaryColor As Color = Color.FromArgb(94, 148, 255)
    Private ReadOnly secondaryColor As Color = Color.FromArgb(125, 137, 149)
    Private ReadOnly backgroundColor As Color = Color.FromArgb(23, 25, 35)
    Private ReadOnly surfaceColor As Color = Color.FromArgb(32, 35, 47)
    Private ReadOnly textColor As Color = Color.FromArgb(224, 224, 224)
    Private ReadOnly accentColor As Color = Color.FromArgb(255, 107, 107)
    Private ReadOnly successColor As Color = Color.FromArgb(46, 204, 113)

    Public Class BatFileInfo
        Public Property Name As String
        Public Property Path As String
        Public Property DateAdded As DateTime

        Public Sub New()
        End Sub

        Public Sub New(name As String, path As String)
            Me.Name = name
            Me.Path = path
            Me.DateAdded = DateTime.Now
        End Sub

        Public Overrides Function ToString() As String
            Return Name
        End Function
    End Class

    Public Sub New()
        CheckAdministratorPrivileges()
        SetupUI()
        LoadSavedBatFiles()
    End Sub

    Private Sub CheckAdministratorPrivileges()
        Dim identity = WindowsIdentity.GetCurrent()
        Dim principal = New WindowsPrincipal(identity)

        If Not principal.IsInRole(WindowsBuiltInRole.Administrator) Then
            ' Restart as administrator
            Try
                Dim startInfo As New ProcessStartInfo With {
                    .UseShellExecute = True,
                    .WorkingDirectory = Environment.CurrentDirectory,
                    .FileName = Application.ExecutablePath,
                    .Verb = "runas"
                }
                Process.Start(startInfo)
                Application.Exit()
            Catch ex As Exception
                MessageBox.Show("This application requires administrator privileges to run batch files properly.", "Administrator Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End If
    End Sub

    Private Sub SetupUI()
        ' Form Properties
        Me.Text = "Batch File Manager"
        Me.Size = New Size(1200, 800)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = backgroundColor
        Me.FormBorderStyle = FormBorderStyle.None
        Me.MinimumSize = New Size(1000, 700)

        CreateControls()
        SetupEvents()
        ApplyTheme()
        SetupDragForm()
    End Sub

    Private Sub SetupDragForm()
        ' Allow dragging the form
        AddHandler titlePanel.MouseDown, AddressOf OnTitleMouseDown
        AddHandler lblTitle.MouseDown, AddressOf OnTitleMouseDown
    End Sub

    Private Sub OnTitleMouseDown(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Left Then
            ' Release capture and send WM_NCLBUTTONDOWN
            ReleaseCapture()
            SendMessage(Me.Handle, &HA1, &H2, 0)
        End If
    End Sub

    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function ReleaseCapture() As Boolean
    End Function

    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As Integer, lParam As Integer) As Integer
    End Function

    Private Sub CreateControls()
        ' Main Container
        mainContainer = New TableLayoutPanel With {
        .Dock = DockStyle.Fill,
        .RowCount = 4,
        .ColumnCount = 1,
        .BackColor = backgroundColor,
        .Padding = New Padding(10)
    }
        mainContainer.RowStyles.Add(New RowStyle(SizeType.Absolute, 70))   ' Title
        mainContainer.RowStyles.Add(New RowStyle(SizeType.Absolute, 140))  ' Top Panel
        mainContainer.RowStyles.Add(New RowStyle(SizeType.Percent, 100))   ' Output Panel
        mainContainer.RowStyles.Add(New RowStyle(SizeType.Absolute, 70))   ' Middle Panel

        ' Title Panel
        titlePanel = New Guna2Panel With {.Dock = DockStyle.Fill, .FillColor = surfaceColor, .Margin = New Padding(0, 0, 0, 5)}
        lblTitle = New Guna2HtmlLabel With {.Text = "<b>⚡ Batch File Manager</b>", .Font = New Font("Segoe UI", 18, FontStyle.Bold), .ForeColor = primaryColor, .Location = New Point(20, 20), .Size = New Size(300, 30), .Cursor = Cursors.SizeAll}
        lblAdmin = New Guna2HtmlLabel With {.Text = "<b>🛡️ Administrator</b>", .Font = New Font("Segoe UI", 10, FontStyle.Bold), .ForeColor = successColor, .Location = New Point(320, 25), .Size = New Size(120, 20)}
        btnClose = New Guna2Button With {.Text = "✕", .Size = New Size(40, 30), .Location = New Point(1140, 20), .FillColor = accentColor, .Font = New Font("Segoe UI", 12, FontStyle.Bold), .BorderRadius = 5}
        btnMinimize = New Guna2Button With {.Text = "−", .Size = New Size(40, 30), .Location = New Point(1090, 20), .FillColor = secondaryColor, .Font = New Font("Segoe UI", 12, FontStyle.Bold), .BorderRadius = 5}
        titlePanel.Controls.AddRange({lblTitle, lblAdmin, btnMinimize, btnClose})

        ' Top Panel (File Management)
        topPanel = New Guna2Panel With {.Dock = DockStyle.Fill, .FillColor = surfaceColor, .Padding = New Padding(20), .Margin = New Padding(0, 0, 0, 5)}
        Dim pathLabel As New Guna2HtmlLabel With {.Text = "<b>📁 Batch File:</b>", .Font = New Font("Segoe UI", 11, FontStyle.Bold), .ForeColor = textColor, .Location = New Point(0, 15), .Size = New Size(100, 25)}
        txtBatPath = New Guna2TextBox With {.Location = New Point(110, 10), .Size = New Size(650, 36), .ReadOnly = True, .PlaceholderText = "Select a batch file..."}
        btnBrowse = New Guna2Button With {.Text = "📂 Browse", .Location = New Point(770, 10), .Size = New Size(100, 36)}
        lblStatus = New Guna2HtmlLabel With {.Text = "<b>🔴 Stopped</b>", .Font = New Font("Segoe UI", 11, FontStyle.Bold), .ForeColor = accentColor, .Location = New Point(880, 15), .Size = New Size(150, 25)}

        Dim savedLabel As New Guna2HtmlLabel With {.Text = "<b>💾 Saved:</b>", .Font = New Font("Segoe UI", 11, FontStyle.Bold), .ForeColor = textColor, .Location = New Point(0, 65), .Size = New Size(100, 25)}
        cmbBatFiles = New Guna2ComboBox With {.Location = New Point(110, 60), .Size = New Size(350, 36), .DropDownStyle = ComboBoxStyle.DropDownList}
        btnSave = New Guna2Button With {.Text = "💾 Save", .Location = New Point(470, 60), .Size = New Size(90, 36)}
        btnDelete = New Guna2Button With {.Text = "🗑️ Delete", .Location = New Point(580, 60), .Size = New Size(90, 36), .TextOffset = New Point(0, 5)}
        btnStart = New Guna2Button With {.Text = "▶️ Start", .Location = New Point(680, 60), .Size = New Size(90, 36)}
        btnStop = New Guna2Button With {.Text = "⏹️ Stop", .Location = New Point(780, 60), .Size = New Size(90, 36)}
        btnRestart = New Guna2Button With {.Text = "🔄 Restart", .Location = New Point(880, 60), .Size = New Size(100, 36)}

        topPanel.Controls.AddRange({pathLabel, txtBatPath, btnBrowse, lblStatus, savedLabel, cmbBatFiles, btnSave, btnDelete, btnStart, btnStop, btnRestart})

        ' Middle Panel (Command Input)
        middlePanel = New Guna2Panel With {.Dock = DockStyle.Fill, .FillColor = surfaceColor, .Padding = New Padding(20, 15, 20, 15), .Margin = New Padding(0, 0, 0, 5)}
        Dim cmdLabel As New Guna2HtmlLabel With {.Text = "<b>⌨️ Command:</b>", .Font = New Font("Segoe UI", 11, FontStyle.Bold), .ForeColor = textColor, .Location = New Point(0, 15), .Size = New Size(100, 25)}
        txtCommand = New Guna2TextBox With {.Location = New Point(110, 10), .Size = New Size(760, 36), .PlaceholderText = "Enter command to send to batch file..."}
        btnSend = New Guna2Button With {.Text = "📤 Send", .Location = New Point(880, 10), .Size = New Size(80, 36), .TextOffset = New Point(0, 4)}
        btnClear = New Guna2Button With {.Text = "🧹 Clear", .Location = New Point(970, 10), .Size = New Size(80, 36), .TextOffset = New Point(0, 5)}
        middlePanel.Controls.AddRange({cmdLabel, txtCommand, btnSend, btnClear})

        ' Output Panel
        outputPanel = New Guna2Panel With {.Dock = DockStyle.Fill, .FillColor = surfaceColor, .Padding = New Padding(20, 30, 20, 20)}
        Dim outputLabel As New Guna2HtmlLabel With {.Text = "<b>📺 Output Console</b>", .Font = New Font("Segoe UI", 12, FontStyle.Bold), .ForeColor = textColor, .Location = New Point(0, 0), .Size = New Size(200, 25)}
        txtOutput = New RichTextBox With {
        .Dock = DockStyle.Fill,
        .Margin = New Padding(0),
        .ReadOnly = True,
        .BackColor = Color.FromArgb(15, 17, 26),
        .ForeColor = Color.FromArgb(0, 255, 127),
        .Font = New Font("Consolas", 12, FontStyle.Regular), ' Increased font size
        .BorderStyle = BorderStyle.None,
        .ScrollBars = RichTextBoxScrollBars.Vertical
    }
        outputPanel.Controls.AddRange({outputLabel, txtOutput})

        ' Add panels to layout
        mainContainer.Controls.Add(titlePanel, 0, 0)
        mainContainer.Controls.Add(topPanel, 0, 1)
        mainContainer.Controls.Add(outputPanel, 0, 2)
        mainContainer.Controls.Add(middlePanel, 0, 3)

        Controls.Add(mainContainer)
    End Sub

    Private Sub ApplyTheme()
        ' Apply Guna2 styling to all controls
        ApplyTextBoxTheme(txtBatPath)
        ApplyTextBoxTheme(txtCommand)
        ApplyComboBoxTheme(cmbBatFiles)

        ' Apply button themes
        ApplyPrimaryButtonTheme(btnStart)
        ApplyDangerButtonTheme(btnStop)
        ApplySecondaryButtonTheme(btnRestart)
        ApplySecondaryButtonTheme(btnBrowse)
        ApplySecondaryButtonTheme(btnSave)
        ApplyDangerButtonTheme(btnDelete)
        ApplyPrimaryButtonTheme(btnSend)
        ApplySecondaryButtonTheme(btnClear)

        ' Window control buttons
        With btnClose
            .BorderRadius = 5
            .FillColor = accentColor
            .ForeColor = Color.White
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
            .HoverState.FillColor = Color.FromArgb(255, 127, 127)
        End With

        With btnMinimize
            .BorderRadius = 5
            .FillColor = secondaryColor
            .ForeColor = Color.White
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
            .HoverState.FillColor = Color.FromArgb(145, 157, 169)
        End With

        ' Panel styling
        For Each panel As Guna2Panel In {topPanel, middlePanel, outputPanel, titlePanel}
            panel.BorderRadius = 12
            panel.ShadowDecoration.Enabled = True
            panel.ShadowDecoration.Color = Color.Black
            panel.ShadowDecoration.Depth = 15
            panel.ShadowDecoration.Shadow = New Padding(0, 0, 3, 3)
        Next
    End Sub

    Private Sub ApplyTextBoxTheme(tb As Guna2TextBox)
        With tb
            .BorderRadius = 8
            .FillColor = backgroundColor
            .BorderColor = secondaryColor
            .ForeColor = textColor
            .Font = New Font("Segoe UI", 10)
            .PlaceholderForeColor = secondaryColor
            .FocusedState.BorderColor = primaryColor
            .HoverState.BorderColor = primaryColor
            .BorderThickness = 1
        End With
    End Sub

    Private Sub ApplyComboBoxTheme(cb As Guna2ComboBox)
        With cb
            .BorderRadius = 8
            .FillColor = backgroundColor
            .BorderColor = secondaryColor
            .ForeColor = textColor
            .Font = New Font("Segoe UI", 10)
            .FocusedColor = primaryColor
            .HoverState.BorderColor = primaryColor
            .ItemsAppearance.BackColor = backgroundColor
            .ItemsAppearance.ForeColor = textColor
            .ItemsAppearance.SelectedBackColor = primaryColor
            .ItemsAppearance.SelectedForeColor = Color.White
            .BorderThickness = 1
        End With
    End Sub

    Private Sub ApplyPrimaryButtonTheme(btn As Guna2Button)
        With btn
            .BorderRadius = 8
            .FillColor = primaryColor
            .ForeColor = Color.White
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
            .HoverState.FillColor = Color.FromArgb(114, 168, 255)
            .ShadowDecoration.Enabled = True
            .ShadowDecoration.Color = Color.FromArgb(50, primaryColor.R, primaryColor.G, primaryColor.B)
            .ShadowDecoration.Depth = 10
        End With
    End Sub

    Private Sub ApplySecondaryButtonTheme(btn As Guna2Button)
        With btn
            .BorderRadius = 8
            .FillColor = backgroundColor
            .BorderColor = secondaryColor
            .BorderThickness = 1
            .ForeColor = textColor
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
            .HoverState.FillColor = secondaryColor
            .HoverState.ForeColor = Color.White
        End With
    End Sub

    Private Sub ApplyDangerButtonTheme(btn As Guna2Button)
        With btn
            .BorderRadius = 8
            .FillColor = accentColor
            .ForeColor = Color.White
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
            .HoverState.FillColor = Color.FromArgb(255, 127, 127)
            .ShadowDecoration.Enabled = True
            .ShadowDecoration.Color = Color.FromArgb(50, accentColor.R, accentColor.G, accentColor.B)
            .ShadowDecoration.Depth = 10
        End With
    End Sub

    Private Sub SetupEvents()
        AddHandler btnBrowse.Click, Sub() BrowseFile()
        AddHandler btnSave.Click, Sub() SaveCurrentBatFile()
        AddHandler btnDelete.Click, Sub() DeleteSelectedBatFile()
        AddHandler btnStart.Click, Sub() StartProcess()
        AddHandler btnStop.Click, Sub() StopProcess()
        AddHandler btnRestart.Click, Sub() RestartProcess()
        AddHandler btnSend.Click, Sub() SendCommand(txtCommand.Text)
        AddHandler btnClear.Click, Sub() txtOutput.Clear()
        AddHandler cmbBatFiles.SelectedIndexChanged, Sub() LoadSelectedBatFile()
        AddHandler txtCommand.KeyDown, AddressOf OnCommandKeyDown
        AddHandler btnClose.Click, Sub() Me.Close()
        AddHandler btnMinimize.Click, Sub() Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub OnCommandKeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            SendCommand(txtCommand.Text)
            txtCommand.Clear()
        End If
    End Sub

    Private Sub LoadSavedBatFiles()
        Try
            If File.Exists(SETTINGS_FILE) Then
                Dim json = File.ReadAllText(SETTINGS_FILE)
                batFilesList = JsonConvert.DeserializeObject(Of List(Of BatFileInfo))(json)
                cmbBatFiles.Items.Clear()
                cmbBatFiles.Items.AddRange(batFilesList.ToArray())
            End If
        Catch ex As Exception
            AddToOutput($"Error loading saved files: {ex.Message}", True)
        End Try
    End Sub

    Private Sub BrowseFile()
        Using ofd As New OpenFileDialog()
            ofd.Filter = "Batch Files (*.bat)|*.bat|Command Files (*.cmd)|*.cmd|All Files (*.*)|*.*"
            ofd.Title = "Select Batch File"
            If ofd.ShowDialog() = DialogResult.OK Then
                batFilePath = ofd.FileName
                txtBatPath.Text = batFilePath
            End If
        End Using
    End Sub

    Private Sub SaveCurrentBatFile()
        If Not File.Exists(batFilePath) Then
            MessageBox.Show("Please select a valid batch file first.", "No File Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Show input dialog for custom name
        Dim customName = ShowInputDialog("Save Batch File", "Enter a name for this batch file:", Path.GetFileNameWithoutExtension(batFilePath))

        If String.IsNullOrWhiteSpace(customName) Then
            Return
        End If

        ' Check if name already exists
        If batFilesList.Any(Function(x) x.Name.Equals(customName, StringComparison.OrdinalIgnoreCase)) Then
            MessageBox.Show("A batch file with this name already exists. Please choose a different name.", "Name Exists", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Check if path already exists
        If batFilesList.Any(Function(x) x.Path = batFilePath) Then
            MessageBox.Show("This batch file path is already saved.", "Already Exists", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim newItem As New BatFileInfo(customName, batFilePath)
        batFilesList.Add(newItem)

        Try
            File.WriteAllText(SETTINGS_FILE, JsonConvert.SerializeObject(batFilesList, Formatting.Indented))
            cmbBatFiles.Items.Add(newItem)
            AddToOutput($"Batch file '{customName}' saved successfully!")
        Catch ex As Exception
            AddToOutput($"Error saving file: {ex.Message}", True)
        End Try
    End Sub

    Private Function ShowInputDialog(title As String, prompt As String, defaultValue As String) As String
        Using inputForm As New Form()
            inputForm.Text = title
            inputForm.Size = New Size(400, 150)
            inputForm.StartPosition = FormStartPosition.CenterParent
            inputForm.BackColor = backgroundColor
            inputForm.ForeColor = textColor
            inputForm.FormBorderStyle = FormBorderStyle.FixedDialog
            inputForm.MaximizeBox = False
            inputForm.MinimizeBox = False

            Dim lblPrompt As New Label() With {
                .Text = prompt,
                .Location = New Point(20, 20),
                .Size = New Size(350, 20),
                .ForeColor = textColor
            }

            Dim txtInput As New TextBox() With {
                .Text = defaultValue,
                .Location = New Point(20, 50),
                .Size = New Size(350, 25),
                .BackColor = surfaceColor,
                .ForeColor = textColor,
                .BorderStyle = BorderStyle.FixedSingle
            }

            Dim btnOK As New Button() With {
                .Text = "OK",
                .Location = New Point(215, 85),
                .Size = New Size(75, 25),
                .DialogResult = DialogResult.OK,
                .BackColor = primaryColor,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat
            }

            Dim btnCancel As New Button() With {
                .Text = "Cancel",
                .Location = New Point(295, 85),
                .Size = New Size(75, 25),
                .DialogResult = DialogResult.Cancel,
                .BackColor = secondaryColor,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat
            }

            inputForm.Controls.AddRange({lblPrompt, txtInput, btnOK, btnCancel})
            inputForm.AcceptButton = btnOK
            inputForm.CancelButton = btnCancel

            txtInput.SelectAll()
            txtInput.Focus()

            If inputForm.ShowDialog() = DialogResult.OK Then
                Return txtInput.Text.Trim()
            Else
                Return String.Empty
            End If
        End Using
    End Function

    Private Sub DeleteSelectedBatFile()
        If cmbBatFiles.SelectedItem IsNot Nothing Then
            Dim result = MessageBox.Show("Are you sure you want to remove this batch file from the list?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                Dim item = CType(cmbBatFiles.SelectedItem, BatFileInfo)
                batFilesList.Remove(item)
                cmbBatFiles.Items.Remove(item)
                File.WriteAllText(SETTINGS_FILE, JsonConvert.SerializeObject(batFilesList, Formatting.Indented))
                txtBatPath.Clear()
                batFilePath = ""
                AddToOutput($"Removed '{item.Name}' from saved files.")
            End If
        Else
            MessageBox.Show("Please select a batch file to delete.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Sub LoadSelectedBatFile()
        If cmbBatFiles.SelectedItem IsNot Nothing Then
            Dim item = CType(cmbBatFiles.SelectedItem, BatFileInfo)
            batFilePath = item.Path
            txtBatPath.Text = batFilePath

            ' Check if file still exists
            If Not File.Exists(batFilePath) Then
                AddToOutput($"Warning: File '{item.Name}' not found at path: {batFilePath}", True)
            End If
        End If
    End Sub

    Private Sub StartProcess()
        If Not File.Exists(batFilePath) Then
            MessageBox.Show("Please select a valid batch file first.", "No File Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If isProcessRunning Then
            MessageBox.Show("A process is already running. Stop it first.", "Process Running", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            Dim workingDirectory = Path.GetDirectoryName(batFilePath)

            cmdProcess = New Process With {
                .StartInfo = New ProcessStartInfo() With {
                    .FileName = batFilePath,
                    .WorkingDirectory = workingDirectory,
                    .RedirectStandardInput = True,
                    .RedirectStandardOutput = True,
                    .RedirectStandardError = True,
                    .UseShellExecute = False,
                    .CreateNoWindow = True,
                    .Verb = ""
                },
                .EnableRaisingEvents = True
            }

            AddHandler cmdProcess.OutputDataReceived, Sub(s, e)
                                                          If e.Data IsNot Nothing Then
                                                              txtOutput.Invoke(Sub()
                                                                                   AddToOutput(e.Data)
                                                                               End Sub)
                                                          End If
                                                      End Sub

            AddHandler cmdProcess.ErrorDataReceived, Sub(s, e)
                                                         If e.Data IsNot Nothing Then
                                                             txtOutput.Invoke(Sub()
                                                                                  AddToOutput($"ERROR: {e.Data}", True)
                                                                              End Sub)
                                                         End If
                                                     End Sub

            AddHandler cmdProcess.Exited, Sub()
                                              txtOutput.Invoke(Sub()
                                                                   isProcessRunning = False
                                                                   lblStatus.Text = "<b>🔴 Stopped</b>"
                                                                   lblStatus.ForeColor = accentColor
                                                                   AddToOutput("Process exited.")
                                                               End Sub)
                                          End Sub

            cmdProcess.Start()
            cmdProcess.BeginOutputReadLine()
            cmdProcess.BeginErrorReadLine()

            isProcessRunning = True
            lblStatus.Text = "<b>🟢 Running</b>"
            lblStatus.ForeColor = successColor

            AddToOutput($"Started: {Path.GetFileName(batFilePath)}")
            AddToOutput($"Working Directory: {workingDirectory}")

        Catch ex As Exception
            AddToOutput($"Error starting process: {ex.Message}", True)
        End Try
    End Sub
    Private Async Function StopProcess() As Task
        If cmdProcess IsNot Nothing AndAlso Not cmdProcess.HasExited Then
            Try
                Dim pid As Integer = cmdProcess.Id

                Await Task.Run(Sub()
                                   Dim killProcess As New Process()
                                   killProcess.StartInfo.FileName = "taskkill"
                                   killProcess.StartInfo.Arguments = $"/PID {pid} /T /F"
                                   killProcess.StartInfo.CreateNoWindow = True
                                   killProcess.StartInfo.UseShellExecute = False
                                   killProcess.StartInfo.RedirectStandardOutput = True
                                   killProcess.StartInfo.RedirectStandardError = True

                                   killProcess.Start()
                                   killProcess.WaitForExit()
                               End Sub)

                If Not cmdProcess.HasExited Then cmdProcess.Kill()
                cmdProcess.Dispose()

                isProcessRunning = False
                lblStatus.Text = "<b>🔴 Stopped</b>"
                lblStatus.ForeColor = accentColor
                AddToOutput("Process (and its children) terminated.")
            Catch ex As Exception
                AddToOutput($"Error stopping process: {ex.Message}", True)
            End Try
        Else
            MessageBox.Show("No process is currently running.", "No Process", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Function


    Private Async Sub RestartProcess()
        If isProcessRunning Then
            Await StopProcess()
            Await Task.Delay(2000)
        End If
        StartProcess()
    End Sub


    Private Sub SendCommand(cmd As String)
        If String.IsNullOrWhiteSpace(cmd) Then Return

        If cmdProcess IsNot Nothing AndAlso Not cmdProcess.HasExited Then
            Try
                cmdProcess.StandardInput.WriteLine(cmd)
                AddToOutput($"> {cmd}")
            Catch ex As Exception
                AddToOutput($"Error sending command: {ex.Message}", True)
            End Try
        Else
            MessageBox.Show("No process is running to send commands to.", "No Process", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Sub AddToOutput(text As String, Optional isError As Boolean = False)
        Dim timestamp = DateTime.Now.ToString("HH:mm:ss")

        txtOutput.SuspendLayout()

        txtOutput.SelectionStart = txtOutput.TextLength
        txtOutput.SelectionLength = 0

        If isError Then
            txtOutput.SelectionColor = accentColor
        Else
            txtOutput.SelectionColor = Color.FromArgb(0, 255, 127)
        End If

        txtOutput.AppendText($"[{timestamp}] {text}{vbCrLf}")

        ' Ensure scroll to bottom
        txtOutput.SelectionStart = txtOutput.Text.Length
        txtOutput.ScrollToCaret()

        txtOutput.ResumeLayout()
    End Sub


    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        If isProcessRunning Then
            Dim result = MessageBox.Show("A batch process is still running. Do you want to stop it and exit?", "Process Running", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                StopProcess()
            Else
                e.Cancel = True
                Return
            End If
        End If
        MyBase.OnFormClosing(e)
    End Sub

    ' Override WndProc to handle window dragging
    Protected Overrides Sub WndProc(ByRef m As Message)
        Const WM_NCHITTEST As Integer = &H84
        Const HTCLIENT As Integer = &H1
        Const HTCAPTION As Integer = &H2

        If m.Msg = WM_NCHITTEST Then
            MyBase.WndProc(m)
            If m.Result = New IntPtr(HTCLIENT) Then
                Dim pos = Me.PointToClient(New Point(m.LParam.ToInt32()))
                If titlePanel.Bounds.Contains(pos) Then
                    m.Result = New IntPtr(HTCAPTION)
                End If
            End If
        Else
            MyBase.WndProc(m)
        End If
    End Sub
End Class