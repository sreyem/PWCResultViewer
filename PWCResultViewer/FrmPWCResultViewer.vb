Imports System.Globalization
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Drawing.Imaging
Imports System.Windows.Forms.DataVisualization.Charting

Public Class FrmPWCResultViewer

#Region " Properties & Data Fields "

    ' Öffentlicher Zugriff auf das Chart für die Main.vb
    Public ReadOnly Property MainChartControl As Chart
        Get
            Return chartMain
        End Get
    End Property

    ' Data structure for daily records
    Private Structure DataRecord
        Public Property DateValue As DateTime
        Public Property WaterCol As Double
        Public Property Benthic As Double
    End Structure

    ' Store output file paths for each type
    Private ReadOnly filePaths As New Dictionary(Of String, String) From {
        {"Parent", Nothing},
        {"Daughter", Nothing},
        {"Granddaughter", Nothing}
    }

    ' Currently loaded output data per type
    Private ReadOnly allRecords As New Dictionary(Of String, List(Of DataRecord)) From {
        {"Parent", New List(Of DataRecord)()},
        {"Daughter", New List(Of DataRecord)()},
        {"Granddaughter", New List(Of DataRecord)()}
    }

    ' Currently selected active source key
    Private currentSourceKey As String = "Parent"

    ' UI Controls
    Private menuStrip1 As MenuStrip
    Private statusStrip1 As StatusStrip
    Private lblStatusFilePath As ToolStripStatusLabel
    Private pnlRadioContainer As Panel
    Private rbParent As RadioButton
    Private rbDaughter As RadioButton
    Private rbGranddaughter As RadioButton
    Private chkShowDetailChart As CheckBox
    Private chartSplitContainer As SplitContainer
    Private WithEvents chartMain As Chart
    Private WithEvents chartDetail As Chart

    ' Chart Title & Legends
    Private mainTitle As Title
    Private summaryLegend As Legend

    ' Dynamischer App-Titel aus My.Application.Info
    Private ReadOnly Property AppTitle As String
        Get
            Dim name As String = My.Application.Info.Title
            If String.IsNullOrEmpty(name) Then
                name = My.Application.Info.ProductName
            End If
            Dim version As String = My.Application.Info.Version.ToString()
            Return $"{name} v{version}"
        End Get
    End Property

    ' Status-Variable für den Umschalter (True = SW, False = Benthic)
    Private isSwMode As Boolean = True

    ' Farbdefinition für Benthic aus dem Benthic-Graphen auslesen oder festlegen
    Private ReadOnly colorBenthic As Drawing.Color = Drawing.Color.SaddleBrown ' Oder z. B. chartMain.Series("Benthic").Color

    Private WithEvents chartToolTip As New ToolTip()
    Private lastHoveredArea As String = ""

#End Region

#Region " Form Lifecycle & Reset "

    Private Sub Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Text = $"{AppTitle}"
        Me.Size = New Drawing.Size(width:=1400, height:=900)
        Me.AllowDrop = True

        ' 2. Fensterrand fixieren (Verhindert das Skalieren per Maus)
        Me.FormBorderStyle = FormBorderStyle.FixedSingle

        ' 3. Maximieren-Button deaktivieren (optional, aber empfohlen)
        Me.MaximizeBox = False

        ' 4. Minimale und maximale Größe auf den gleichen Wert setzen (Sicherheitsnetz)
        Me.MinimumSize = Me.Size
        Me.MaximumSize = Me.Size

        SetupMenuAndStatusStrip()
        SetupRadioButtons()
        SetupLayout()
        SetupMainChart()
        SetupDetailChart()
        CheckCMDArgs()

    End Sub

    Private Sub CheckCMDArgs()
        ' Befehlszeilen-Argumente auslesen
        Dim args() As String = Environment.GetCommandLineArgs()

        ' Parameter 2 parsen (UI anzeigen oder nicht - Standard ist False)
        Dim showUI As Boolean = False
        If args.Length > 2 Then
            Boolean.TryParse(args(2), showUI)
        End If

        ' args(0) ist der Pfad der EXE selbst. Das erste echte Argument ist args(1)
        If args.Length > 1 Then
            Dim inputPath As String = args(1).Trim(""""c) ' Anführungszeichen entfernen

            ' Fall 1: Pfad ist ein Ordner
            If Directory.Exists(inputPath) Then
                'Me.Show() ' Form anzeigen, damit das Chart gerendert werden kann
                'Application.DoEvents()

                ProcessDirectoryBatch(inputPath)

                ' Nach Batch-Verarbeitung die Anwendung automatisch beenden
                Me.Close()
                Me.Dispose()
                End

                ' Fall 2: Pfad ist eine einzelne Datei
            ElseIf File.Exists(inputPath) Then

                ProcessSelectedFile(inputPath)
                Dim gifPath As String = Path.ChangeExtension(inputPath, ".gif")
                SaveChartToGif(gifPath)
                Me.Close()
                Me.Dispose()
                End
            End If
        End If
    End Sub

#Region " Batch mode"

    ''' <summary>
    ''' Verarbeitet alle *.out- und *.csv-Dateien in einem Ordner und speichert automatisch die Charts als GIF.
    ''' </summary>
    Public Sub ProcessDirectoryBatch(directoryPath As String)

        If Not Directory.Exists(directoryPath) Then
            Console.WriteLine($"Directory not found: {directoryPath}")
            Exit Sub
        End If

        ' Alle passenden Dateien (.out und .csv) aus dem Ordner holen
        Dim files = Directory.GetFiles(directoryPath, "*.*") _
                             .Where(Function(f) f.EndsWith(".out", StringComparison.OrdinalIgnoreCase) OrElse
                                                f.EndsWith(".csv", StringComparison.OrdinalIgnoreCase)) _
                             .ToList()

        If files.Count = 0 Then
            Console.WriteLine($"No .out or .csv files found in directory: {directoryPath}")
            Exit Sub
        End If

        For Each file In files


            ProcessSelectedFile(file)
            Dim gifPath As String = Path.ChangeExtension(file, ".gif")
            SaveChartToGif(gifPath)
        Next

        'Try
        '    ' Datei laden und verarbeiten
        '    ProcessSelectedFile(file)

        '    ' Dateipfad für das GIF generieren (gleicher Ordner, gleicher Name mit .gif Extension)
        '    Dim gifPath As String = Path.Combine(directoryPath, Path.GetFileNameWithoutExtension(file) & ".gif")

        '    ' Grafik speichern
        '    SaveChartToGif(gifPath)

        '    Console.WriteLine($"Successfully generated GIF for: {Path.GetFileName(file)}")
        'Catch ex As Exception
        '    Console.WriteLine($"Error processing file {Path.GetFileName(file)}: {ex.Message}")
        'End Try
        'Next
    End Sub

    ''' <summary>
    ''' Hilfsmethode zum direkten Speichern des Charts als GIF an einem Zielpfad
    ''' </summary>
    Private Sub SaveChartToGif(outputPath As String)
        If chartMain.Width <= 0 OrElse chartMain.Height <= 0 Then Exit Sub

        Using bmp As New Bitmap(chartMain.Width, chartMain.Height)
            Using g As Graphics = Graphics.FromImage(bmp)
                g.Clear(Color.White)
                Dim bmpMain As New Bitmap(chartMain.Width, chartMain.Height)
                chartMain.DrawToBitmap(bmpMain, New Rectangle(0, 0, chartMain.Width, chartMain.Height))
                g.DrawImage(bmpMain, 0, 0)
            End Using
            bmp.Save(outputPath, ImageFormat.Gif)
        End Using
    End Sub

#End Region

    Private Sub ResetAppData()
        For Each key In filePaths.Keys.ToList()
            filePaths(key) = Nothing
        Next

        For Each key In allRecords.Keys.ToList()
            allRecords(key).Clear()
        Next

        RemoveHandler rbParent.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged
        RemoveHandler rbDaughter.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged
        RemoveHandler rbGranddaughter.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged

        rbParent.Enabled = False
        rbDaughter.Enabled = False
        rbGranddaughter.Enabled = False
        rbParent.Checked = True
        currentSourceKey = "Parent"

        AddHandler rbParent.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged
        AddHandler rbDaughter.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged
        AddHandler rbGranddaughter.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged

        chartMain.Series("Surface Water Max").Points.Clear()
        chartMain.Series("Benthic Max").Points.Clear()
        If chartMain.Series.IndexOf("BoxPlotSeries") >= 0 Then chartMain.Series("BoxPlotSeries").Points.Clear()
        If chartMain.Series.IndexOf("BoxPlotPoints") >= 0 Then chartMain.Series("BoxPlotPoints").Points.Clear()
        chartDetail.Series("Surface Water (Daily)").Points.Clear()
        chartDetail.Series("Benthic (Daily)").Points.Clear()
        summaryLegend.CustomItems.Clear()

        If mainTitle IsNot Nothing Then
            mainTitle.Text = "Yearly max concentrations of Parent"
        End If

        Me.Text = $"{AppTitle}"
        lblStatusFilePath.Text = "Drag & drop PWC file or File -> Open"


        If chartMain.Titles.FindByName("BoxPlotTitle") IsNot Nothing Then
            chartMain.Titles("BoxPlotTitle").Visible = False
        End If

    End Sub

#End Region

#Region " Setup Form & Controls "

    Private Sub SetupMenuAndStatusStrip()

        menuStrip1 = New MenuStrip() With {.Font = New Drawing.Font("Segoe UI", 12.0!)}

        ' File Menu
        Dim menuFile As New ToolStripMenuItem("File") With {.Font = New Drawing.Font("Segoe UI", 12.0!)}

        Dim menuOpen As New ToolStripMenuItem("Open...", Nothing, AddressOf MenuOpen_Click) With {
            .ShortcutKeys = Keys.Control Or Keys.O, .Font = New Drawing.Font("Segoe UI", 12.0!)
        }

        Dim menuProcessFolder As New ToolStripMenuItem("Process Folder...", Nothing, AddressOf MenuProcessFolder_Click) With {
            .ShortcutKeys = Keys.Control Or Keys.Shift Or Keys.O, .Font = New Drawing.Font("Segoe UI", 12.0!)
        }

        Dim menuSave As New ToolStripMenuItem("Save Chart as GIF...", Nothing, AddressOf MenuSaveAsGif_Click) With {
            .ShortcutKeys = Keys.Control Or Keys.S, .Font = New Drawing.Font("Segoe UI", 12.0!)
        }

        Dim menuReset As New ToolStripMenuItem("Reset", Nothing, AddressOf MenuReset_Click) With {
            .Font = New Drawing.Font("Segoe UI", 12.0!)
        }

        Dim menuExit As New ToolStripMenuItem("Exit", Nothing, Sub() Me.Close()) With {
            .Font = New Drawing.Font("Segoe UI", 12.0!)
        }

        Dim menuSaveYearMaxCsv As New ToolStripMenuItem("Save year max values as csv...", Nothing, AddressOf MenuSaveYearMaxCsv_Click) With {
            .Font = New Drawing.Font("Segoe UI", 12.0!)
        }

        menuFile.DropDownItems.Add(menuOpen)
        menuFile.DropDownItems.Add(menuProcessFolder)
        menuFile.DropDownItems.Add(menuSave)
        menuFile.DropDownItems.Add(menuSaveYearMaxCsv)
        menuFile.DropDownItems.Add(menuReset)
        menuFile.DropDownItems.Add(New ToolStripSeparator())
        menuFile.DropDownItems.Add(menuExit)

        ' Help Menu
        Dim menuHelp As New ToolStripMenuItem("Help") With {.Font = New Drawing.Font("Segoe UI", 12.0!)}
        Dim menuHowTo As New ToolStripMenuItem("How to generate PWC 3.XXXX *.out files...", Nothing, AddressOf MenuHelp_Click) With {
            .Font = New Drawing.Font("Segoe UI", 12.0!)
        }

        Dim menuShowDetailsHelp As New ToolStripMenuItem("Help on 'Show Details'", Nothing, AddressOf MenuShowDetailsHelp_Click) With {
            .Font = New Drawing.Font("Segoe UI", 12.0!)
        }
        menuHelp.DropDownItems.Add(menuHowTo)
        'menuHelp.DropDownItems.Add(menuShowDetailsHelp)



        menuHelp.DropDownItems.Add(New ToolStripSeparator()) ' Trennlinie zur besseren Übersicht

        menuStrip1.Items.Add(menuFile)
        menuStrip1.Items.Add(menuHelp)

        Me.MainMenuStrip = menuStrip1
        Me.Controls.Add(menuStrip1)

        ' Status Strip
        statusStrip1 = New StatusStrip()
        lblStatusFilePath = New ToolStripStatusLabel("Drag & drop PWC file or File -> Open") With {
            .Spring = True, .TextAlign = Drawing.ContentAlignment.MiddleLeft, .Font = New Drawing.Font("Segoe UI", 12.0!)
        }
        statusStrip1.Items.Add(lblStatusFilePath)
        Me.Controls.Add(statusStrip1)
    End Sub

    Private Sub MenuSaveYearMaxCsv_Click(sender As Object, e As EventArgs)
        Dim activeFile As String = filePaths(currentSourceKey)

        ' Prüfen, ob eine Datei geladen ist und Daten vorhanden sind
        If String.IsNullOrEmpty(activeFile) OrElse Not allRecords.ContainsKey(currentSourceKey) OrElse allRecords(currentSourceKey).Count = 0 Then
            MessageBox.Show("No data to save.", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        ' Dateinamen gemäß Anforderung auf *_YearMax.csv setzen
        Dim defaultFileName As String = Path.GetFileNameWithoutExtension(activeFile) & "_YearMax.csv"
        Dim initialDirectory As String = Path.GetDirectoryName(activeFile)

        Using sfd As New SaveFileDialog()
            sfd.Filter = "CSV Files (*.csv)|*.csv"
            sfd.FileName = defaultFileName
            sfd.InitialDirectory = initialDirectory

            If sfd.ShowDialog() = DialogResult.OK Then
                SaveYearMaxCsv(sfd.FileName)
            End If
        End Using
    End Sub

    Private Sub SaveYearMaxCsv(outputPath As String)
        Try
            Dim records = allRecords(currentSourceKey)
            If records.Count = 0 Then Exit Sub

            ' Gruppierung nach Jahr (identisch mit PlotMainChart)
            Dim yearlyGroups = records.GroupBy(Function(r) r.DateValue.Year).OrderBy(Function(g) g.Key).ToList()

            Using writer As New StreamWriter(outputPath, False, System.Text.Encoding.UTF8)
                ' Header wie gefordert: Year, Max Sw, Max Benthic
                writer.WriteLine("Year, Sw (ppb),Benthic (ppb)")

                For Each group In yearlyGroups
                    Dim yr As Integer = group.Key
                    Dim maxSw As Double = group.Max(Function(r) r.WaterCol)
                    Dim maxBenthic As Double = group.Max(Function(r) r.Benthic)

                    ' Speichern mit InvariantCulture (Punkt als Dezimaltrenner)
                    writer.WriteLine(String.Format(CultureInfo.InvariantCulture, "{0}, {1:F6}, {2:F6}", yr, maxSw, maxBenthic))
                Next
            End Using

            MessageBox.Show($"File saved at:{vbCrLf}{outputPath}", "Save Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show($"Error while saving the CSV file:{vbCrLf}{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub MenuShowDetailsHelp_Click(sender As Object, e As EventArgs)
        Dim helpText As String =
            "==========================================" & vbCrLf &
            "       HELP: Show Details Feature" & vbCrLf &
            "==========================================" & vbCrLf & vbCrLf &
            "The 'Show Details' functionality provides a second graphical view " &
            "to display the detailed yearly concentration referring to the selected year in the main chart." & vbCrLf &
            "Activate this by checking the 'Show Details' checkbox and select a year from the main chart by clicking on the item"

        MessageBox.Show(helpText, "Help - Show Details", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub SetupRadioButtons()
        pnlRadioContainer = New Panel() With {
            .Dock = DockStyle.Top,
            .Height = 45,
            .Padding = New Padding(10, 5, 10, 5)
        }

        Dim lblSelect As New Label() With {
            .Text = "Data Source:",
            .AutoSize = True,
            .Font = New Drawing.Font("Segoe UI", 12.0!, Drawing.FontStyle.Bold),
            .Location = New Drawing.Point(10, 10)
        }

        rbParent = New RadioButton() With {
            .Text = "Parent",
            .Location = New Drawing.Point(140, 8),
            .AutoSize = True,
            .Checked = True,
            .Enabled = False,
            .Font = New Drawing.Font("Segoe UI", 12.0!)
        }

        rbDaughter = New RadioButton() With {
            .Text = "Daughter",
            .Location = New Drawing.Point(260, 8),
            .AutoSize = True,
            .Enabled = False,
            .Font = New Drawing.Font("Segoe UI", 12.0!)
        }

        rbGranddaughter = New RadioButton() With {
            .Text = "Granddaughter",
            .Location = New Drawing.Point(390, 8),
            .AutoSize = True,
            .Enabled = False,
            .Font = New Drawing.Font("Segoe UI", 12.0!)
        }

        chkShowDetailChart = New CheckBox() With {
            .Text = "Show Details",
            .Location = New Drawing.Point(560, 8),
            .AutoSize = True,
            .Checked = False,
            .Font = New Drawing.Font("Segoe UI", 12.0!)
        }

        AddHandler rbParent.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged
        AddHandler rbDaughter.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged
        AddHandler rbGranddaughter.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged
        AddHandler chkShowDetailChart.CheckedChanged, AddressOf OnShowDetailChart_CheckedChanged

        pnlRadioContainer.Controls.Add(lblSelect)
        pnlRadioContainer.Controls.Add(rbParent)
        pnlRadioContainer.Controls.Add(rbDaughter)
        pnlRadioContainer.Controls.Add(rbGranddaughter)
        pnlRadioContainer.Controls.Add(chkShowDetailChart)

        Me.Controls.Add(pnlRadioContainer)
    End Sub



    Private Sub SetupLayout()
        chartSplitContainer = New SplitContainer() With {
            .Dock = DockStyle.Fill,
            .Orientation = Orientation.Horizontal,
            .SplitterDistance = 400,
            .Panel2Collapsed = True
        }
        Me.Controls.Add(chartSplitContainer)

        chartSplitContainer.BringToFront()
        pnlRadioContainer.BringToFront()
        menuStrip1.BringToFront()

        chartMain = New Chart() With {.Dock = DockStyle.Fill}
        chartSplitContainer.Panel1.Controls.Add(chartMain)

        chartDetail = New Chart() With {.Dock = DockStyle.Fill}
        chartSplitContainer.Panel2.Controls.Add(chartDetail)
    End Sub

    Private Sub SetupMainChart()
        Dim area As New ChartArea("MainArea")

        area.AxisX.Title = "Relative Year"
        area.AxisX.LabelStyle.Font = New Drawing.Font("Segoe UI", 12.0!)
        area.AxisX.TitleFont = New Drawing.Font("Segoe UI", 12.0!, Drawing.FontStyle.Bold)
        area.AxisX.Interval = 1
        area.AxisX.IsMarginVisible = True
        area.AxisX.MajorGrid.LineColor = Drawing.Color.LightGray

        area.AxisY.Title = "Yearly Max. Concentration (ppb)"
        area.AxisY.LabelStyle.Font = New Drawing.Font("Segoe UI", 12.0!)
        area.AxisY.TitleFont = New Drawing.Font("Segoe UI", 12.0!, Drawing.FontStyle.Bold)
        area.AxisY.LabelStyle.Format = "0.00"
        area.AxisY.MajorGrid.LineColor = Drawing.Color.LightGray
        chartMain.ChartAreas.Add(area)

        chartMain.Titles.Clear()
        mainTitle = New Title() With {
            .Name = "MainTitle",
            .Text = "Yearly max concentrations of Parent",
            .Font = New Drawing.Font("Segoe UI", 14.0!, Drawing.FontStyle.Bold),
            .Docking = Docking.Top
        }
        chartMain.Titles.Add(mainTitle)

        summaryLegend = New Legend("SummaryLegend") With {
            .Docking = Docking.Right,
            .Alignment = Drawing.StringAlignment.Near,
            .Title = "Result Overview",
            .TitleFont = New Drawing.Font("Segoe UI", 14.0!, Drawing.FontStyle.Bold),
            .BackColor = Drawing.Color.FromArgb(245, 245, 245),
            .BorderColor = Drawing.Color.LightGray,
            .TextWrapThreshold = 0,     ' Deaktiviert Zeilenumbrüche vollständig
            .IsTextAutoFit = True,      ' Passt Textgröße bei Platzmangel dynamisch an
            .AutoFitMinFontSize = 8     ' Minimale Schriftgröße fürs Auto-Fit
        }
        chartMain.Legends.Add(summaryLegend)

        Dim seriesWater As New Series("Surface Water Max") With {
            .ChartType = SeriesChartType.Point,
            .XValueType = ChartValueType.Double,
            .MarkerStyle = MarkerStyle.Circle,
            .MarkerSize = 10,
            .Color = Drawing.Color.RoyalBlue,
            .IsVisibleInLegend = False
        }

        Dim seriesBenthic As New Series("Benthic Max") With {
            .ChartType = SeriesChartType.Point,
            .XValueType = ChartValueType.Double,
            .MarkerStyle = MarkerStyle.Square,
            .MarkerSize = 10,
            .Color = Drawing.Color.DarkOrange,
            .IsVisibleInLegend = False
        }

        chartMain.Series.Add(seriesWater)
        chartMain.Series.Add(seriesBenthic)

        ' Überschrift für den Boxplot erstellen und an die BoxPlotArea binden
        Dim titleBoxPlot As New Title() With {
            .Name = "BoxPlotTitle",
            .Text = "Surface Water Yearly Max (ppb)",
            .DockedToChartArea = "BoxPlotArea", ' Bindet den Titel direkt an die Boxplot-Fläche
            .IsDockedInsideChartArea = False,   ' Platzierung knapp oberhalb des Plots
            .Font = New Drawing.Font("Segoe UI", 14.0!, Drawing.FontStyle.Bold),
            .ForeColor = Drawing.Color.Black,
            .Visible = False
        }

        ' Zur Chart-Control hinzufügen (vorherige Exemplare zur Sicherheit entfernen)
        If chartMain.Titles.FindByName("BoxPlotTitle") IsNot Nothing Then
            chartMain.Titles.Remove(chartMain.Titles("BoxPlotTitle"))
        End If
        chartMain.Titles.Add(titleBoxPlot)

        ' --- ChartArea & Serien für den Boxplot unter der Result Overview ---
        Dim areaBox As New ChartArea("BoxPlotArea")
        areaBox.Position.Auto = False
        areaBox.Position.X = 72
        areaBox.Position.Y = 55
        areaBox.Position.Width = 26
        areaBox.Position.Height = 40

        ' Achsen ausblenden und Ränder auf 0 setzen
        areaBox.AxisX.Enabled = AxisEnabled.False
        areaBox.AxisY.Enabled = AxisEnabled.False
        areaBox.InnerPlotPosition.Auto = False
        areaBox.InnerPlotPosition.X = 0
        areaBox.InnerPlotPosition.Y = 0
        areaBox.InnerPlotPosition.Width = 100
        areaBox.InnerPlotPosition.Height = 100

        chartMain.ChartAreas.Add(areaBox)

        ' Serie 1: Boxplot-Struktur (Transparent / Ohne Füllung)
        Dim seriesBoxPlot As New Series("BoxPlotSeries") With {
            .ChartType = SeriesChartType.BoxPlot,
            .ChartArea = "BoxPlotArea",
            .IsVisibleInLegend = False,
            .Color = Drawing.Color.FromArgb(0, 255, 255, 255), ' Transparentes Inneres
            .BorderColor = Drawing.Color.ForestGreen,           ' Rahmen & Whisker in Grün
            .BorderWidth = 2
        }

        seriesBoxPlot("BoxPlotTransparentColor") = "Transparent"
        seriesBoxPlot("BoxPlotWhiskerPercentile") = "0"
        seriesBoxPlot("BoxPlotShowMedian") = "True"
        seriesBoxPlot("BoxPlotShowAverage") = "False"

        ' Serie 2: Einzelpunkte
        Dim seriesBoxPoints As New Series("BoxPlotPoints") With {
            .ChartType = SeriesChartType.Point,
            .ChartArea = "BoxPlotArea",
            .IsVisibleInLegend = False,
            .MarkerStyle = MarkerStyle.Circle,
            .MarkerSize = 6,
            .Color = Drawing.Color.RoyalBlue
        }

        chartMain.Series.Add(seriesBoxPlot)
        chartMain.Series.Add(seriesBoxPoints)

    End Sub

    Private Sub SetupDetailChart()
        Dim area As New ChartArea("DetailArea")

        area.AxisX.Title = "Month (Daily values for selected year)"
        area.AxisX.LabelStyle.Font = New Drawing.Font("Segoe UI", 12.0!)
        area.AxisX.TitleFont = New Drawing.Font("Segoe UI", 12.0!, Drawing.FontStyle.Bold)
        area.AxisX.LabelStyle.Format = "MMM"
        area.AxisX.Interval = 1
        area.AxisX.IntervalType = DateTimeIntervalType.Months
        area.AxisX.MajorGrid.LineColor = Drawing.Color.LightGray

        area.AxisY.Title = "Daily Value (ppb)"
        area.AxisY.LabelStyle.Font = New Drawing.Font("Segoe UI", 12.0!)
        area.AxisY.TitleFont = New Drawing.Font("Segoe UI", 12.0!, Drawing.FontStyle.Bold)
        area.AxisY.LabelStyle.Format = "0.00"
        area.AxisY.MajorGrid.LineColor = Drawing.Color.LightGray

        area.CursorX.IsUserSelectionEnabled = True
        area.CursorX.IsUserEnabled = True
        area.AxisX.ScaleView.Zoomable = True
        area.AxisX.ScrollBar.IsPositionedInside = False
        area.AxisX.ScrollBar.Size = 12
        area.AxisX.ScrollBar.ButtonStyle = ScrollBarButtonStyles.All

        chartDetail.ChartAreas.Add(area)

        'Dim legend As New Legend("DetailLegend") With {
        '    .Docking = Docking.Bottom,
        '    .Font = New Drawing.Font("Segoe UI", 12.0!)
        '}
        'chartDetail.Legends.Add(legend)

        Dim seriesWater As New Series("Surface Water (Daily)") With {
            .ChartType = SeriesChartType.Line,
            .XValueType = ChartValueType.Date,
            .BorderWidth = 2,
            .Color = Drawing.Color.RoyalBlue
        }

        Dim seriesBenthic As New Series("Benthic (Daily)") With {
            .ChartType = SeriesChartType.Line,
            .XValueType = ChartValueType.Date,
            .BorderWidth = 2,
            .Color = Drawing.Color.DarkOrange
        }

        chartDetail.Series.Add(seriesWater)
        chartDetail.Series.Add(seriesBenthic)

    End Sub

#End Region

#Region " File Processing & Data I/O "

    Public Sub ProcessSelectedFile(filePath As String)
        ResetAppData()

        Dim fileName As String = Path.GetFileName(filePath)

        ' Strikte Trennung: Ist es eine Legacy-Datei oder eine PWC3-Datei?
        If IsLegacyFile(fileName) Then
            ProcessLegacyFile(filePath)
        Else
            ProcessPwc3File(filePath)
        End If

        DisplayCurrentSource()

        If chartMain.Titles.FindByName("BoxPlotTitle") IsNot Nothing Then
            chartMain.Titles("BoxPlotTitle").Visible = True
        End If

    End Sub

    Private Function IsLegacyFile(fileName As String) As Boolean
        Return fileName.EndsWith("_daily.csv", StringComparison.OrdinalIgnoreCase) OrElse
               fileName.IndexOf("_daily", StringComparison.OrdinalIgnoreCase) >= 0
    End Function

    ' ==========================================
    ' ROUTINE 1: Legacy PWC Processing (*_daily.csv)
    ' ==========================================
    Private Sub ProcessLegacyFile(filePath As String)
        Dim directoryPath As String = Path.GetDirectoryName(filePath)
        Dim fileName As String = Path.GetFileName(filePath)

        ' 1. Feststellen, welche Datei explizit aufgerufen wurde
        Dim targetSourceKey As String = "Parent"
        If fileName.IndexOf("_Degradate1_daily.csv", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
           fileName.IndexOf("Degradate1", StringComparison.OrdinalIgnoreCase) >= 0 Then
            targetSourceKey = "Daughter"
        ElseIf fileName.IndexOf("_Degradate2_daily.csv", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               fileName.IndexOf("Degradate2", StringComparison.OrdinalIgnoreCase) >= 0 Then
            targetSourceKey = "Granddaughter"
        End If

        ' 2. Suffixe bereinigen um das gemeinsame Präfix zu ermitteln
        Dim prefix As String = fileName
        Dim legacySuffixes() As String = {"_Parent_daily.csv", "_Degradate1_daily.csv", "_Degradate2_daily.csv", "_daily.csv"}

        For Each suffix In legacySuffixes
            If prefix.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) Then
                prefix = prefix.Substring(0, prefix.Length - suffix.Length)
                Exit For
            End If
        Next

        ' 3. Suchen aller zusammengehörigen Legacy-Dateien im Ordner
        filePaths("Parent") = FindMatchingFile(directoryPath, prefix, "_Parent_daily.csv")
        filePaths("Daughter") = FindMatchingFile(directoryPath, prefix, "_Degradate1_daily.csv")
        filePaths("Granddaughter") = FindMatchingFile(directoryPath, prefix, "_Degradate2_daily.csv")

        If String.IsNullOrEmpty(filePaths("Parent")) AndAlso targetSourceKey = "Parent" Then
            filePaths("Parent") = filePath
        End If

        ' 4. Daten einlesen
        For Each key In filePaths.Keys.ToList()
            If Not String.IsNullOrEmpty(filePaths(key)) Then
                LoadLegacyCsvFile(filePaths(key), allRecords(key))
            End If
        Next

        ' 5. RadioButtons basierend auf vorhandenen Daten aktivieren
        RemoveHandler rbParent.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged
        RemoveHandler rbDaughter.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged
        RemoveHandler rbGranddaughter.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged

        rbParent.Enabled = (allRecords("Parent").Count > 0)
        rbDaughter.Enabled = (allRecords("Daughter").Count > 0)
        rbGranddaughter.Enabled = (allRecords("Granddaughter").Count > 0)

        ' 6. Gezielt die aufgerufene Quelldatei als aktiv setzen
        If targetSourceKey = "Daughter" AndAlso rbDaughter.Enabled Then
            rbDaughter.Checked = True
            currentSourceKey = "Daughter"
        ElseIf targetSourceKey = "Granddaughter" AndAlso rbGranddaughter.Enabled Then
            rbGranddaughter.Checked = True
            currentSourceKey = "Granddaughter"
        ElseIf rbParent.Enabled Then
            rbParent.Checked = True
            currentSourceKey = "Parent"
        Else
            If rbDaughter.Enabled Then
                rbDaughter.Checked = True
                currentSourceKey = "Daughter"
            ElseIf rbGranddaughter.Enabled Then
                rbGranddaughter.Checked = True
                currentSourceKey = "Granddaughter"
            End If
        End If

        AddHandler rbParent.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged
        AddHandler rbDaughter.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged
        AddHandler rbGranddaughter.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged

        lblStatusFilePath.Text = $"Legacy File: {fileName}  |  Folder: {directoryPath}  |  Prefix: {prefix}"
    End Sub

    Private Sub LoadLegacyCsvFile(filePath As String, targetList As List(Of DataRecord))
        Try
            Dim currentDate As New DateTime(1980, 1, 1)
            Const multiplier As Double = 1000000.0

            Using reader As New StreamReader(filePath)
                ' Exakt die ersten 5 Zeilen überspringen (Header/Metadaten)
                For i As Integer = 1 To 5
                    If reader.EndOfStream Then Exit Sub
                    reader.ReadLine()
                Next

                ' Ab Zeile 6 Daten einlesen
                While Not reader.EndOfStream
                    Dim line As String = reader.ReadLine()
                    If String.IsNullOrWhiteSpace(line) Then Continue While

                    Dim parts() As String = line.TrimEnd(","c).Split(","c)

                    If parts.Length >= 3 Then
                        Dim waterVal As Double
                        Dim benthicVal As Double

                        If Double.TryParse(parts(1).Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, waterVal) AndAlso
                           Double.TryParse(parts(2).Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, benthicVal) Then

                            targetList.Add(New DataRecord With {
                                .DateValue = currentDate,
                                .WaterCol = waterVal * multiplier,
                                .Benthic = benthicVal * multiplier
                            })

                            currentDate = currentDate.AddDays(1)
                        End If
                    End If
                End While
            End Using

        Catch ex As Exception
            MessageBox.Show($"Error reading Legacy CSV file {Path.GetFileName(filePath)}:{vbCrLf}{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' ==========================================
    ' ROUTINE 2: PWC 3 Standard Processing (*.out)
    ' ==========================================
    Private Sub ProcessPwc3File(filePath As String)
        Dim directoryPath As String = Path.GetDirectoryName(filePath)
        Dim fileName As String = Path.GetFileName(filePath)

        Dim targetSourceKey As String = "Parent"
        If fileName.IndexOf("daughter", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
           fileName.IndexOf("deg1", StringComparison.OrdinalIgnoreCase) >= 0 Then
            targetSourceKey = "Daughter"
        ElseIf fileName.IndexOf("granddaughter", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               fileName.IndexOf("deg2", StringComparison.OrdinalIgnoreCase) >= 0 Then
            targetSourceKey = "Granddaughter"
        End If

        Dim prefix As String = fileName
        Dim knownSuffixes() As String = {"_parent_Pond.out", "_daughter_Pond.out", "_granddaughter_Pond.out", "_summary.txt", "_summary_Deg1.txt", "_summary_Deg2.txt"}

        For Each suffix In knownSuffixes
            If prefix.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) Then
                prefix = prefix.Substring(0, prefix.Length - suffix.Length)
                Exit For
            End If
        Next

        filePaths("Parent") = FindMatchingFile(directoryPath, prefix, "_parent_Pond.out")
        filePaths("Daughter") = FindMatchingFile(directoryPath, prefix, "_daughter_Pond.out")
        filePaths("Granddaughter") = FindMatchingFile(directoryPath, prefix, "_granddaughter_Pond.out")

        For Each key In filePaths.Keys.ToList()
            If Not String.IsNullOrEmpty(filePaths(key)) Then
                LoadPwc3OutFile(filePaths(key), allRecords(key))
            End If
        Next

        RemoveHandler rbParent.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged
        RemoveHandler rbDaughter.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged
        RemoveHandler rbGranddaughter.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged

        rbParent.Enabled = (allRecords("Parent").Count > 0)
        rbDaughter.Enabled = (allRecords("Daughter").Count > 0)
        rbGranddaughter.Enabled = (allRecords("Granddaughter").Count > 0)

        If targetSourceKey = "Daughter" AndAlso rbDaughter.Enabled Then
            rbDaughter.Checked = True
            currentSourceKey = "Daughter"
        ElseIf targetSourceKey = "Granddaughter" AndAlso rbGranddaughter.Enabled Then
            rbGranddaughter.Checked = True
            currentSourceKey = "Granddaughter"
        ElseIf rbParent.Enabled Then
            rbParent.Checked = True
            currentSourceKey = "Parent"
        Else
            If rbDaughter.Enabled Then
                rbDaughter.Checked = True
                currentSourceKey = "Daughter"
            ElseIf rbGranddaughter.Enabled Then
                rbGranddaughter.Checked = True
                currentSourceKey = "Granddaughter"
            End If
        End If

        AddHandler rbParent.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged
        AddHandler rbDaughter.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged
        AddHandler rbGranddaughter.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged

        lblStatusFilePath.Text = $"PWC3 File: {fileName}  |  Folder: {directoryPath}  |  Prefix: {prefix}"
    End Sub

    Private Sub LoadPwc3OutFile(filePath As String, targetList As List(Of DataRecord))
        Try
            Dim currentDate As New DateTime(1980, 1, 1)
            Const multiplier As Double = 1000000.0

            Using reader As New StreamReader(filePath)
                Dim header As String = reader.ReadLine()

                While Not reader.EndOfStream
                    Dim line As String = reader.ReadLine()
                    If String.IsNullOrWhiteSpace(line) Then Continue While

                    Dim parts() As String = line.TrimEnd(","c).Split(","c)

                    If parts.Length >= 3 Then
                        Dim waterVal As Double
                        Dim benthicVal As Double

                        If Double.TryParse(parts(1).Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, waterVal) AndAlso
                           Double.TryParse(parts(2).Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, benthicVal) Then

                            targetList.Add(New DataRecord With {
                                .DateValue = currentDate,
                                .WaterCol = waterVal * multiplier,
                                .Benthic = benthicVal * multiplier
                            })

                            currentDate = currentDate.AddDays(1)
                        End If
                    End If
                End While
            End Using

        Catch ex As Exception
            MessageBox.Show($"Error reading PWC3 file {Path.GetFileName(filePath)}:{vbCrLf}{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function FindMatchingFile(dir As String, prefix As String, suffix As String) As String
        Dim exact As String = Path.Combine(dir, prefix & suffix)
        If File.Exists(exact) Then Return exact

        Dim matches = Directory.GetFiles(dir, $"{prefix}*{suffix}")
        If matches.Length > 0 Then Return matches(0)

        Return Nothing
    End Function

    ' ==========================================
    ' SUMMARY FILE PARSING (PWC 3 & LEGACY)
    ' ==========================================
    Private Sub LoadSummaryFileForCurrentSource(mainFilePath As String)
        summaryLegend.CustomItems.Clear()

        ' Farbige Legenden-Symbole
        Dim itemWater As New LegendItem()

        'itemWater.Color = Drawing.Color.RoyalBlue
        ''itemWater.MarkerStyle = MarkerStyle.Circle
        'itemWater.Cells.Add(New LegendCell(LegendCellType.SeriesSymbol, "") With {
        '    .SeriesSymbolSize = New Drawing.Size(12, 12)
        '})

        itemWater.Cells.Add(New LegendCell(LegendCellType.Text, "Concentration in Surface Water   (ppb)") With {
            .Font = New Drawing.Font("Segoe UI", 11.0!, Drawing.FontStyle.Bold),
            .ForeColor = Drawing.Color.RoyalBlue,
            .Alignment = Drawing.ContentAlignment.MiddleLeft
        })

        Dim itemBenthic As New LegendItem()

        'itemBenthic.Color = Drawing.Color.DarkOrange
        ''itemBenthic.MarkerStyle = MarkerStyle.Square
        'itemBenthic.Cells.Add(New LegendCell(LegendCellType.SeriesSymbol, "") With {
        '    .SeriesSymbolSize = New Drawing.Size(12, 12)
        '})
        itemBenthic.Cells.Add(New LegendCell(LegendCellType.Text, "                             Benthic System (ppb)") With {
            .Font = New Drawing.Font("Segoe UI", 11.0!, Drawing.FontStyle.Bold),
            .ForeColor = Drawing.Color.DarkOrange,
            .Alignment = Drawing.ContentAlignment.MiddleLeft
        })

        summaryLegend.CustomItems.Add(itemWater)
        summaryLegend.CustomItems.Add(itemBenthic)


        ' STRIKTE TRENNUNG: Legacy vs PWC3 Summary
        If IsLegacyFile(Path.GetFileName(mainFilePath)) Then
            LoadLegacySummaryFile(mainFilePath)
        Else
            LoadPwc3SummaryFile(mainFilePath)
        End If
    End Sub

#Region "Legacy Summary Processing"

    ' ==========================================
    ' ROUTINE: Legacy Summary Processing (*.txt)
    ' ==========================================
    Private Sub LoadLegacySummaryFile(mainFilePath As String)
        Dim directoryPath As String = Path.GetDirectoryName(mainFilePath)
        Dim fileName As String = Path.GetFileName(mainFilePath)

        Dim prefix As String = fileName
        Dim legacySuffixes() As String = {"_Parent_daily.csv", "_Degradate1_daily.csv", "_Degradate2_daily.csv", "_daily.csv"}

        For Each suffix In legacySuffixes
            If prefix.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) Then
                prefix = prefix.Substring(0, prefix.Length - suffix.Length)
                Exit For
            End If
        Next

        Dim summarySuffix As String = "_Parent.txt"
        Select Case currentSourceKey
            Case "Daughter"
                summarySuffix = "_Degradate1.txt"
            Case "Granddaughter"
                summarySuffix = "_Degradate2.txt"
        End Select

        Dim expectedSummaryName As String = $"{prefix}{summarySuffix}"
        Dim summaryFilePath As String = Path.Combine(directoryPath, expectedSummaryName)

        If Not File.Exists(summaryFilePath) Then
            Dim searchPattern As String = $"{prefix}*{summarySuffix}"
            Dim matchingFiles = Directory.GetFiles(directoryPath, searchPattern)
            If matchingFiles.Length > 0 Then
                summaryFilePath = matchingFiles(0)
            Else
                AddSummaryLegendItem("Status:", $"{summarySuffix} missing")
                Exit Sub
            End If
        End If

        Try
            Dim lines() As String = File.ReadAllLines(summaryFilePath)
            Dim dictValues As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            For Each line As String In lines
                ' Strippe Windows-Zeilenumbrüche (\r) und Trimme
                Dim trimmedLine As String = line.Replace(vbCr, "").Trim()

                If trimmedLine.Contains("="c) Then
                    ' Nur am ERSTEN Gleichheitszeichen trennen
                    Dim parts() As String = trimmedLine.Split(New Char() {"="c}, 2)
                    If parts.Length >= 2 Then
                        Dim key As String = parts(0).Trim()
                        Dim val As String = parts(1).Trim()

                        ' Falls Duplikate vorkommen, nur den ersten Schlüssel speichern
                        If Not dictValues.ContainsKey(key) Then
                            dictValues.Add(key, val)
                        End If
                    End If
                End If
            Next



            AddSummaryLegendItem("--------", "-----------")

            ' --- Top-Konzentrationen ausgeben ---
            AddLegacyValueToSummary(dictValues, "SW       1-d avg:", "1-d avg 1-in-10.0", "")
            AddLegacyValueToSummary(dictValues, "SW       4-d avg:", "4-d avg 1-in-10.0", "")
            AddLegacyValueToSummary(dictValues, "SW      21-d avg:", "21-d avg 1-in-10.0", "")
            AddLegacyValueToSummary(dictValues, "SW      60-d avg:", "60-d avg 1-in-10.0", "")
            AddLegacyValueToSummary(dictValues, "SW      90-d avg:", "90-d avg 1-in-10.0", "")

            AddLegacyValueToSummary(dictValues, "Benthic  1-d avg:", "Benthic Pore Water 1-d   avg 1-in-10.0", "")
            AddLegacyValueToSummary(dictValues, "        21-d avg:", "Benthic Pore Water 21-d avg 1-in-10.0", "")

            ' --- Einträge (Erosion, Runoff, Drift) ---
            AddSummaryLegendItem("Entry Paths", "")

            AddLegacyValueToSummary(dictValues, "Runoff :", "Due to Runoff", "%", isPercentage:=True)
            AddLegacyValueToSummary(dictValues, "Erosion:", "Due to Erosion", "%", isPercentage:=True)
            AddLegacyValueToSummary(dictValues, "Drift  :", "Due to Drift", "%", isPercentage:=True)

            '' --- Inputs (Halbwertszeiten DT50) ---
            'AddSummaryLegendItem("Inputs", " ----- ")

            'AddLegacyValueToSummary(dictValues, "SW      DT50:", "water column half Life", "d")
            'AddLegacyValueToSummary(dictValues, "Benthic DT50:", "benthic Half Life", "d")

        Catch ex As Exception
            MessageBox.Show($"Error reading Legacy summary file ({Path.GetFileName(summaryFilePath)}):{vbCrLf}{ex.Message}", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    ' ==========================================
    ' HILFSMETHODE: Value Search & Matching
    ' ==========================================
    Private Sub AddLegacyValueToSummary(dict As Dictionary(Of String, String), label As String, keyName As String, unit As String, Optional isPercentage As Boolean = False)
        ' Normalisiert Whitespaces (NBSP \u00A0 & multiple Leerzeichen/Tabs -> Einzel-Space)
        Dim CleanString = Function(s As String) Regex.Replace(s.Replace(ChrW(160), " "c), "\s+", " ").Trim()
        Dim targetClean As String = CleanString(keyName)

        For Each kvp In dict
            Dim keyClean As String = CleanString(kvp.Key)
            Dim valClean As String = CleanString(kvp.Value)

            ' Fall 1: Suchbegriff steht LINKS vom '=' (z. B. "Due to Runoff = 0.9950")
            If keyClean.IndexOf(targetClean, StringComparison.OrdinalIgnoreCase) >= 0 Then
                ExtractAndAddValue(kvp.Value, label, unit, isPercentage)
                Exit Sub
            End If

            ' Fall 2: Suchbegriff steht RECHTS vom '=' (z. B. "54.00 = water column half Life")
            If valClean.IndexOf(targetClean, StringComparison.OrdinalIgnoreCase) >= 0 Then
                ExtractAndAddValue(kvp.Key, label, unit, isPercentage)
                Exit Sub
            End If
        Next

        ' Falls nicht gefunden
        AddSummaryLegendItem(label, "N/A")
    End Sub

    ' ==========================================
    ' HILFSMETHODE: Zahl parsen & formatieren
    ' ==========================================
    Private Sub ExtractAndAddValue(rawText As String, label As String, unit As String, isPercentage As Boolean)
        ' Extrahiert den ersten numerischen Wert aus dem Text
        Dim firstToken As String = Regex.Split(rawText.Trim(), "\s+")(0).Replace("ppb", "").Replace("%", "").Trim()
        Dim parsedVal As Double

        If Double.TryParse(firstToken, NumberStyles.Float, CultureInfo.InvariantCulture, parsedVal) Then
            If isPercentage Then
                Dim pctVal As Double = If(parsedVal <= 1.0 AndAlso parsedVal > 0, parsedVal * 100.0, parsedVal)
                AddSummaryLegendItem(label, $"{pctVal:F2} %")
            Else
                If unit = "d" Then
                    AddSummaryLegendItem(label, $"{parsedVal:F1} {unit}")
                Else
                    AddSummaryLegendItem(label, $"{parsedVal:F3} {unit}")
                End If
            End If
        Else
            AddSummaryLegendItem(label, "N/A")
        End If
    End Sub

#End Region

    ' --- ROUTINE B: Standard PWC 3 Summary Processing ---
    Private Sub LoadPwc3SummaryFile(mainFilePath As String)
        Dim directoryPath As String = Path.GetDirectoryName(mainFilePath)
        Dim fileNameWithoutExt As String = Path.GetFileNameWithoutExtension(mainFilePath)

        Dim prefix As String = fileNameWithoutExt
        If prefix.Contains("_") Then
            prefix = prefix.Split("_"c)(0)
        End If

        Dim summarySuffix As String = "_summary.txt"
        Select Case currentSourceKey
            Case "Daughter"
                summarySuffix = "_summary_Deg1.txt"
            Case "Granddaughter"
                summarySuffix = "_summary_Deg2.txt"
        End Select

        Dim expectedSummaryName As String = $"{prefix}{summarySuffix}"
        Dim summaryFilePath As String = Path.Combine(directoryPath, expectedSummaryName)

        If Not File.Exists(summaryFilePath) Then
            Dim searchPattern As String = $"{prefix}*{summarySuffix}"
            Dim matchingFiles = Directory.GetFiles(directoryPath, searchPattern)
            If matchingFiles.Length > 0 Then
                summaryFilePath = matchingFiles(0)
            Else
                AddSummaryLegendItem("Status:", $"{summarySuffix} missing")
                Exit Sub
            End If
        End If

        Try
            Dim lines() As String = File.ReadAllLines(summaryFilePath)
            Dim headerLine As String = Nothing
            Dim dataLine As String = Nothing

            Dim targetRowStart As String = fileNameWithoutExt
            Dim firstUnderscoreIdx As Integer = targetRowStart.IndexOf("_"c)
            If firstUnderscoreIdx >= 0 AndAlso firstUnderscoreIdx < targetRowStart.Length - 1 Then
                targetRowStart = targetRowStart.Substring(firstUnderscoreIdx + 1)
            End If

            Const pondSuffix As String = "_pond"
            If targetRowStart.EndsWith(pondSuffix, StringComparison.OrdinalIgnoreCase) Then
                targetRowStart = targetRowStart.Substring(0, targetRowStart.Length - pondSuffix.Length)
            End If

            If currentSourceKey = "Daughter" Then
                targetRowStart = targetRowStart.Replace("daughter", "deg1")
            ElseIf currentSourceKey = "Granddaughter" Then
                targetRowStart = targetRowStart.Replace("granddaughter", "deg2")
            End If

            For Each line As String In lines
                Dim trimmedLine As String = line.Trim()

                If trimmedLine.Contains("1-d avg") OrElse trimmedLine.Contains("Runoff") Then
                    headerLine = trimmedLine
                End If

                If trimmedLine.StartsWith(targetRowStart, StringComparison.OrdinalIgnoreCase) Then
                    dataLine = trimmedLine
                End If
            Next

            If headerLine IsNot Nothing AndAlso dataLine IsNot Nothing Then
                AddSummaryLegendItem("--------", "-----------")
                AddSummaryLegendItem("SW       1-d avg:", GetValueFromHeader(headerLine, dataLine, "1-d avg", isPercentage:=False)) ' & " ppb")
                AddSummaryLegendItem("         4-d avg:", GetValueFromHeader(headerLine, dataLine, "4-d avg", isPercentage:=False)) ' & " ppb")
                AddSummaryLegendItem("        21-d avg:", GetValueFromHeader(headerLine, dataLine, "21-d avg", isPercentage:=False)) ' & " ppb")
                AddSummaryLegendItem("Benthic  1-d avg:", GetValueFromHeader(headerLine, dataLine, "B 1-day", isPercentage:=False)) ' & " ppb")
                AddSummaryLegendItem("        21-d avg:", GetValueFromHeader(headerLine, dataLine, "B 21-d avg", isPercentage:=False)) ' & " ppb")
                AddSummaryLegendItem("Entry Paths", " ----- ")
                AddSummaryLegendItem("Runoff :", GetValueFromHeader(headerLine, dataLine, "Runoff Frac", isPercentage:=True))
                AddSummaryLegendItem("Erosion:", GetValueFromHeader(headerLine, dataLine, "Erosn Frac", isPercentage:=True))
                AddSummaryLegendItem("Drift  :", GetValueFromHeader(headerLine, dataLine, "Drift Frac", isPercentage:=True))
            Else
                AddSummaryLegendItem("Status:", "Data row not found")
            End If

        Catch ex As Exception
            MessageBox.Show($"Error reading summary file ({Path.GetFileName(summaryFilePath)}):{vbCrLf}{ex.Message}", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Function GetValueFromHeader(headerLine As String, dataLine As String, keyName As String, Optional isPercentage As Boolean = False) As String
        Try
            Dim headers As String()
            Dim values As String()

            If headerLine.Contains(","c) Then
                headers = headerLine.Split(","c).Select(Function(s) s.Trim()).ToArray()
                values = dataLine.Split(","c).Select(Function(s) s.Trim()).ToArray()
            Else
                headers = Regex.Split(headerLine.Trim(), "\s{2,}|\t").Select(Function(s) s.Trim()).ToArray()
                values = Regex.Split(dataLine.Trim(), "\s{2,}|\t").Select(Function(s) s.Trim()).ToArray()
            End If

            For index As Integer = 0 To headers.Length - 1
                If headers(index).Equals(keyName, StringComparison.OrdinalIgnoreCase) OrElse
                   headers(index).Replace(" ", "").Equals(keyName.Replace(" ", ""), StringComparison.OrdinalIgnoreCase) Then

                    If index < values.Length Then
                        Dim rawVal As String = values(index)
                        Dim parsedVal As Double

                        If Double.TryParse(rawVal, NumberStyles.Float, CultureInfo.InvariantCulture, parsedVal) Then

                            If isPercentage Then
                                Dim pctVal As Double = parsedVal * 100.0
                                Return pctVal.ToString("F2", CultureInfo.InvariantCulture) & " %"
                            End If

                            If Math.Abs(parsedVal) < 0.001 AndAlso parsedVal <> 0 Then
                                Return parsedVal.ToString("0.000E+00", CultureInfo.InvariantCulture)
                            Else
                                Return parsedVal.ToString("F4", CultureInfo.InvariantCulture)
                            End If

                        End If
                        Return rawVal
                    End If
                End If
            Next
        Catch
        End Try

        Return "N/A"
    End Function

#End Region

#Region " Chart Rendering & Plotting "

    Private Sub DisplayCurrentSource()
        If mainTitle IsNot Nothing Then
            mainTitle.Text = $"Yearly max concentrations of {currentSourceKey}"
        End If

        PlotMainChart()

        Dim activeFile As String = filePaths(currentSourceKey)
        If Not String.IsNullOrEmpty(activeFile) Then
            LoadSummaryFileForCurrentSource(activeFile)
            Me.Text = $"PWC Viewer - {currentSourceKey} ({Path.GetFileName(activeFile)})"
        Else
            summaryLegend.CustomItems.Clear()
            Me.Text = "PWC Viewer - No File Loaded"
            chartMain.Update()
        End If
    End Sub



    Private Sub SetBoxPlotVisible(visible As Boolean)
        ' 1. Serien ein-/ausblenden
        If chartMain.Series.FindByName("BoxPlotSeries") IsNot Nothing Then
            chartMain.Series("BoxPlotSeries").Enabled = visible
        End If
        If chartMain.Series.FindByName("BoxPlotPoints") IsNot Nothing Then
            chartMain.Series("BoxPlotPoints").Enabled = visible
        End If

        ' 2. ChartArea ein-/ausblenden
        If chartMain.ChartAreas.FindByName("BoxPlotArea") IsNot Nothing Then
            chartMain.ChartAreas("BoxPlotArea").Visible = visible
        End If

        ' 3. Titel ein-/ausblenden
        Dim title = chartMain.Titles.FindByName("BoxPlotTitle")
        If title IsNot Nothing Then
            title.Visible = visible
        End If
    End Sub

    Private Sub PlotMainChart()
        Dim records = allRecords(currentSourceKey)

        Dim seriesWater = chartMain.Series("Surface Water Max")
        Dim seriesBenthic = chartMain.Series("Benthic Max")

        seriesWater.Points.Clear()
        seriesBenthic.Points.Clear()

        Dim seriesBoxPlot = chartMain.Series("BoxPlotSeries")
        Dim seriesBoxPoints = chartMain.Series("BoxPlotPoints")
        seriesBoxPlot.Points.Clear()
        seriesBoxPoints.Points.Clear()

        If records.Count = 0 Then Exit Sub

        Dim yearlyGroups = records.GroupBy(Function(r) r.DateValue.Year).OrderBy(Function(g) g.Key).ToList()
        If yearlyGroups.Count = 0 Then Exit Sub

        Dim startYear As Integer = yearlyGroups.First().Key

        Dim yearlyMaxWater = yearlyGroups.Select(Function(g) New With {
            .RelativeYear = g.Key - startYear,
            .ActualDate = g.OrderByDescending(Function(r) r.WaterCol).First().DateValue,
            .Value = g.OrderByDescending(Function(r) r.WaterCol).First().WaterCol
        }).ToList()

        Dim yearlyMaxBenthic = yearlyGroups.Select(Function(g) New With {
            .RelativeYear = g.Key - startYear,
            .ActualDate = g.OrderByDescending(Function(r) r.Benthic).First().DateValue,
            .Value = g.OrderByDescending(Function(r) r.Benthic).First().Benthic
        }).ToList()

        For Each item In yearlyMaxWater
            Dim ptIndex As Integer = seriesWater.Points.AddXY(item.RelativeYear, item.Value)
            seriesWater.Points(ptIndex).ToolTip = $"SW Max ({currentSourceKey}){vbCrLf}Relative Year: {item.RelativeYear} (Actual: {item.ActualDate:dd.MM.yyyy}){vbCrLf}Value: {item.Value:F4} ppb{vbCrLf}(Click for detail view)"
            seriesWater.Points(ptIndex).Tag = item.ActualDate
        Next

        For Each item In yearlyMaxBenthic
            Dim ptIndex As Integer = seriesBenthic.Points.AddXY(item.RelativeYear, item.Value)
            seriesBenthic.Points(ptIndex).ToolTip = $"Benthic Max ({currentSourceKey}){vbCrLf}Relative Year: {item.RelativeYear} (Actual: {item.ActualDate:dd.MM.yyyy}){vbCrLf}Value: {item.Value:F4} ppb{vbCrLf}(Click for detail view)"
            seriesBenthic.Points(ptIndex).Tag = item.ActualDate
        Next

        ' --- BEFÜLLUNG UND STYLING DES BOXPLOTS ---
        If Not chkShowDetailChart.Checked Then
            SetBoxPlotVisible(True)

            ' 1. Daten und Design basierend auf dem aktuellen Modus (SW vs. Benthic) festlegen
            Dim values As List(Of Double)
            Dim currentTitle As String
            Dim pointColor As Drawing.Color
            Dim valuePrefix As String

            If isSwMode Then
                ' SW-Modus
                values = yearlyMaxWater.Select(Function(x) x.Value).OrderBy(Function(v) v).ToList()
                currentTitle = "Surface Water Yearly max (ppb)"
                valuePrefix = "SW Max"

                ' Dynamisch die Farbe der SW-Serie auslesen (mit Fallback, falls der Name abweicht)
                Dim targetSeries = chartMain.Series.FindByName("SW") ' <-- Falls deine Serie anders heißt, hier anpassen (z.B. "Water")
                If targetSeries IsNot Nothing Then
                    pointColor = targetSeries.Color
                Else
                    pointColor = Drawing.Color.RoyalBlue ' Standardfarbe, falls "SW" nicht gefunden wird
                End If
            Else
                ' Benthic-Modus
                values = yearlyMaxBenthic.Select(Function(x) x.Value).OrderBy(Function(v) v).ToList()
                currentTitle = "Benthic Yearly Max (ppb)"
                valuePrefix = "Benthic Max"

                ' Auch für Benthic optional prüfen, ob eine Serie existiert, oder die feste Farbe nutzen
                Dim benthicSeries = chartMain.Series.FindByName("Benthic")
                If benthicSeries IsNot Nothing Then
                    pointColor = benthicSeries.Color
                Else
                    pointColor = Drawing.Color.SandyBrown ' Hellbraun / Benthic-Ton
                End If
            End If

            ' 2. Titel des Boxplots aktualisieren
            Dim boxTitle = chartMain.Titles.FindByName("BoxPlotTitle")
            If boxTitle IsNot Nothing Then
                boxTitle.Text = currentTitle
            End If

            ' 3. Datenpunkte und Serien zurücksetzen & befüllen
            seriesBoxPoints.Points.Clear()

            If values.Count > 0 Then
                For Each val As Double In values
                    Dim ptIdx As Integer = seriesBoxPoints.Points.AddXY(1.0, val)
                    seriesBoxPoints.Points(ptIdx).ToolTip = $"{valuePrefix}: {val:F4} ppb"
                Next

                ' Boxplot an Punkte-Serie knüpfen
                seriesBoxPlot("BoxPlotSeries") = "BoxPlotPoints"
                seriesBoxPlot.Points.Clear()

                ' Farben zuweisen (Punkte in der aktiven Farbe, Box-Rahmen in Grün)
                seriesBoxPoints.Color = pointColor
                seriesBoxPlot.BorderColor = Drawing.Color.Black
                seriesBoxPlot.Color = Drawing.Color.FromArgb(0, 255, 255, 255)

                ' 4. Min, Max, Median Beschriftungen setzen (rechtsbündig mit Leerzeichen)
                Dim minVal As Double = values.First()
                Dim maxVal As Double = values.Last()
                Dim medianVal As Double
                Dim count As Integer = values.Count

                If count Mod 2 = 0 Then
                    medianVal = (values(count \ 2 - 1) + values(count \ 2)) / 2.0
                Else
                    medianVal = values(count \ 2)
                End If

                Dim labelFont As New Drawing.Font("Segoe UI", 12.0!, Drawing.FontStyle.Bold)

                Dim minPoint = seriesBoxPoints.Points.FirstOrDefault(Function(p) p.YValues(0) = minVal)
                If minPoint IsNot Nothing Then
                    minPoint.Label = $"              Min: {minVal:F2}"
                    minPoint.Font = labelFont
                    minPoint.LabelForeColor = Drawing.Color.Black
                    minPoint("LabelStyle") = "Right"
                End If

                Dim maxPoint = seriesBoxPoints.Points.FirstOrDefault(Function(p) p.YValues(0) = maxVal)
                If maxPoint IsNot Nothing Then
                    maxPoint.Label = $"              Max: {maxVal:F2}"
                    maxPoint.Font = labelFont
                    maxPoint.LabelForeColor = Drawing.Color.Black
                    maxPoint("LabelStyle") = "Right"
                End If

                Dim medianPoint = seriesBoxPoints.Points.OrderBy(Function(p) Math.Abs(p.YValues(0) - medianVal)).FirstOrDefault()
                If medianPoint IsNot Nothing AndAlso medianPoint IsNot minPoint AndAlso medianPoint IsNot maxPoint Then
                    medianPoint.Label = $"              Med: {medianVal:F2}"
                    medianPoint.Font = labelFont
                    medianPoint.LabelForeColor = Drawing.Color.Black
                    medianPoint("LabelStyle") = "Right"
                End If

                ' X-Achsen-Limits zur Platzierung sichern
                Dim boxArea = chartMain.ChartAreas(seriesBoxPoints.ChartArea)
                boxArea.AxisX.Minimum = 0.0
                boxArea.AxisX.Maximum = 2.5
            End If
        End If

        chartMain.ChartAreas("MainArea").RecalculateAxesScale()
        chartMain.ChartAreas("BoxPlotArea").RecalculateAxesScale()

        If yearlyGroups.Any() Then
            UpdateDetailChart(yearlyGroups.First().Key)
        End If

    End Sub

    Private Sub UpdateDetailChart(selectedYear As Integer)
        Dim records = allRecords(currentSourceKey)

        Dim seriesWater = chartDetail.Series("Surface Water (Daily)")
        Dim seriesBenthic = chartDetail.Series("Benthic (Daily)")

        seriesWater.Points.Clear()
        seriesBenthic.Points.Clear()

        Dim yearData = records.Where(Function(r) r.DateValue.Year = selectedYear) _
                              .OrderBy(Function(r) r.DateValue) _
                              .ToList()

        For Each record In yearData
            Dim pWaterIdx As Integer = seriesWater.Points.AddXY(record.DateValue, record.WaterCol)
            seriesWater.Points(pWaterIdx).ToolTip = $"Water Col ({currentSourceKey}){vbCrLf}Date: {record.DateValue:dd.MM.yyyy}{vbCrLf}Value: {record.WaterCol:F4} ppb"

            Dim pBenthicIdx As Integer = seriesBenthic.Points.AddXY(record.DateValue, record.Benthic)
            seriesBenthic.Points(pBenthicIdx).ToolTip = $"Benthic ({currentSourceKey}){vbCrLf}Date: {record.DateValue:dd.MM.yyyy}{vbCrLf}Value: {record.Benthic:F4} ppb"
        Next

        Dim area = chartDetail.ChartAreas("DetailArea")
        area.AxisX.ScaleView.ZoomReset(0)
        area.AxisX.Title = $"Daily Values for Year {selectedYear} ({currentSourceKey})"
        area.AxisX.IntervalType = DateTimeIntervalType.Months
        area.AxisX.Interval = 1
        area.RecalculateAxesScale()
    End Sub

    Private Sub AddSummaryLegendItem(label As String, value As String)
        Dim item As New LegendItem()

        Dim combinedText As String

        ' Trenner / Abschnitte (z. B. "**** Concentrations" oder "Entries") ohne Ausrichtung darstellen
        If String.IsNullOrEmpty(value) OrElse value = " ----- " Then
            combinedText = label
        Else
            ' Label auf z. B. 15 Zeichen mit Leerzeichen auffüllen, 
            ' gefolgt von 4 festen Leerzeichen Abstand zum Wert
            Const labelWidth As Integer = 7
            combinedText = $"{label.PadRight(labelWidth)}    {value}"
        End If

        Dim cellCombined As New LegendCell(LegendCellType.Text, combinedText) With {
            .Alignment = Drawing.ContentAlignment.MiddleLeft,
            .Font = New Drawing.Font("Courier New", 12.0!, Drawing.FontStyle.Bold)
        }

        item.Cells.Add(cellCombined)
        summaryLegend.CustomItems.Add(item)
    End Sub

    'Private Sub AddSummaryLegendItem(label As String, value As String)
    '    Dim item As New LegendItem()

    '    ' Trenner / Kategorien (wie z. B. "Entries") ohne Leerzeichen behandeln
    '    Dim combinedText As String
    '    If String.IsNullOrEmpty(value) OrElse value = " ----- " Then
    '        combinedText = label
    '    Else
    '        ' Label und Value mit exakt 4 Leerzeichen verbinden
    '        combinedText = $"{label}    {value}"
    '    End If

    '    Dim cellCombined As New LegendCell(LegendCellType.Text, combinedText) With {
    '        .Alignment = Drawing.ContentAlignment.MiddleLeft,
    '        .Font = New Drawing.Font("Courier New", 12.0!, Drawing.FontStyle.Bold)
    '    }

    '    item.Cells.Add(cellCombined)
    '    summaryLegend.CustomItems.Add(item)
    'End Sub

    'Private Sub AddSummaryLegendItem(label As String, value As String)
    '    Dim item As New LegendItem()

    '    Dim cellLabel As New LegendCell(LegendCellType.Text, label) With {
    '        .Alignment = Drawing.ContentAlignment.MiddleLeft,
    '        .Font = New Drawing.Font("Courier New", 12.0!, Drawing.FontStyle.Bold)
    '    }

    '    Dim cellValue As New LegendCell(LegendCellType.Text, value) With {
    '        .Alignment = Drawing.ContentAlignment.MiddleRight,
    '        .Font = New Drawing.Font("Courier New", 12.0!, Drawing.FontStyle.Regular)
    '    }

    '    item.Cells.Add(cellLabel)
    '    item.Cells.Add(cellValue)

    '    summaryLegend.CustomItems.Add(item)
    'End Sub

#End Region

#Region " Event Handlers & User Interaction "

    ' --- Menu Events ---

    Private Sub MenuOpen_Click(sender As Object, e As EventArgs)
        Using ofd As New OpenFileDialog()
            ofd.Filter = "PWC Output Files (*.out;*.csv)|*.out;*.csv|All Files (*.*)|*.*"
            ofd.Title = "Select PWC Output File"

            If ofd.ShowDialog() = DialogResult.OK Then
                ProcessSelectedFile(ofd.FileName)
            End If

        End Using
    End Sub

    Private Sub MenuReset_Click(sender As Object, e As EventArgs)
        ResetAppData()
    End Sub

    Private Sub MenuSaveAsGif_Click(sender As Object, e As EventArgs)
        Dim activeFile As String = filePaths(currentSourceKey)
        If String.IsNullOrEmpty(activeFile) Then
            MessageBox.Show("Please open a valid data file first.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Using sfd As New SaveFileDialog()
            sfd.Filter = "GIF Image|*.gif"
            sfd.Title = "Save Charts as GIF"
            sfd.FileName = Path.GetFileNameWithoutExtension(activeFile) & ".gif"

            If sfd.ShowDialog() = DialogResult.OK Then
                Try
                    Dim totalWidth As Integer = chartMain.Width
                    Dim totalHeight As Integer = chartMain.Height

                    Using bmp As New Bitmap(totalWidth, totalHeight)
                        Using g As Graphics = Graphics.FromImage(bmp)
                            g.Clear(Color.White)

                            Dim bmpMain As New Bitmap(chartMain.Width, chartMain.Height)
                            chartMain.DrawToBitmap(bmpMain, New Rectangle(0, 0, chartMain.Width, chartMain.Height))
                            g.DrawImage(bmpMain, 0, 0)
                        End Using

                        bmp.Save(sfd.FileName, ImageFormat.Gif)
                        MessageBox.Show("Saved as: " & Path.GetFileName(sfd.FileName), "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End Using
                Catch ex As Exception
                    MessageBox.Show("Error saving: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End Using
    End Sub

    Private Sub MenuProcessFolder_Click(sender As Object, e As EventArgs)
        Using fbd As New FolderBrowserDialog()
            fbd.Description = "Select Folder Containing PWC Files (*.out / *.csv)"
            fbd.ShowNewFolderButton = False
            fbd.SelectedPath = "C:\"


            If fbd.ShowDialog() = DialogResult.OK Then
                ' Ruft die bereits erstellte Batch-Methode für den gewählten Ordner auf
                ProcessDirectoryBatch(fbd.SelectedPath)


                If chartMain.Titles.FindByName("BoxPlotTitle") IsNot Nothing Then
                    chartMain.Titles("BoxPlotTitle").Visible = True
                End If

                MessageBox.Show("Batch processing completed!", "Process Folder", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End Using
    End Sub

    Private Sub MenuHelp_Click(sender As Object, e As EventArgs)
        Using helpForm As New Form()

            helpForm.Text = "PWC 3.X Output Generation"
            helpForm.Size = New Drawing.Size(750, 900)
            helpForm.StartPosition = FormStartPosition.CenterParent
            helpForm.FormBorderStyle = FormBorderStyle.FixedDialog
            helpForm.MaximizeBox = False
            helpForm.MinimizeBox = False

            ' 1. OK-Button ganz unten fixieren
            Dim btnClose As New Button() With {
                .Text = "OK",
                .DialogResult = DialogResult.OK,
                .Dock = DockStyle.Bottom,
                .Height = 40,
                .Font = New Drawing.Font("Segoe UI", 11.0!, Drawing.FontStyle.Bold)
            }
            helpForm.Controls.Add(btnClose)

            ' 2. Scrollbares Haupt-Panel erstellen
            Dim scrollPanel As New Panel() With {
                .Dock = DockStyle.Fill,
                .AutoScroll = True,
                .Padding = New Padding(15)
            }
            helpForm.Controls.Add(scrollPanel)

            ' --- ERSTER ABSCHNITT (Text + Bild 1) ---

            Dim lblToggleOutput As New Label() With {
                .Text = "To generate the required *.out files in PWC 3, first enable 'Optional Output' by clicking" & vbCrLf & " 'More Tabs -> Toggle More Outputs'" & vbCrLf,
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .MaximumSize = New Drawing.Size(680, 0), ' Verhindert abgeschnittenen Text
                .Padding = New Padding(0, 5, 0, 10),
                .Font = New Drawing.Font("Segoe UI", 11.0!, Drawing.FontStyle.Regular)
            }

            Dim picToggleOutput As New PictureBox() With {
                .Dock = DockStyle.Top,
                .Height = 300,
                .SizeMode = PictureBoxSizeMode.AutoSize,
                .Padding = New Padding(0, 0, 0, 15)
            }
            ' Erstes Bild laden (falls vorhanden)
            If File.Exists("ToggleOutput.png") Then
                picToggleOutput.Load("ToggleOutput.png")
            End If


            ' --- ZWEITER ABSCHNITT (Zusätzlicher Text + Bild 2) ---

            Dim lblOptionalOutput As New Label() With {
                .Text = "Then check the box 'Print daily waterbody output ...'  before running your simulation." & vbCrLf & "",
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .MaximumSize = New Drawing.Size(680, 0),
                .Padding = New Padding(0, 10, 0, 10),
                .Font = New Drawing.Font("Segoe UI", 11.0!, Drawing.FontStyle.Regular)
            }

            Dim picOptionalOutput As New PictureBox() With {
                .Dock = DockStyle.Top,
                .Height = 300,
                .SizeMode = PictureBoxSizeMode.AutoSize,
                .Padding = New Padding(0, 0, 0, 15)
            }
            ' Zweites Bild laden (Name der neuen Datei anpassen!)
            If File.Exists("OptionalOutput.png") Then
                picOptionalOutput.Load("OptionalOutput.png")
            End If


            ' 3. Controls IN UMGEKEHRTER REIHENFOLGE hinzufügen (wegen DockStyle.Top)
            scrollPanel.Controls.Add(picOptionalOutput)
            scrollPanel.Controls.Add(lblOptionalOutput)
            scrollPanel.Controls.Add(picToggleOutput)
            scrollPanel.Controls.Add(lblToggleOutput)

            ' Scroll-Panel nach oben bringen
            scrollPanel.BringToFront()

            helpForm.ShowDialog(Me)
        End Using
    End Sub


    'Private Sub MenuHelp_Click(sender As Object, e As EventArgs)
    '    Using helpForm As New Form()
    '        helpForm.Text = "PWC Output Generation Help"
    '        helpForm.Size = New Drawing.Size(700, 550)
    '        helpForm.StartPosition = FormStartPosition.CenterParent
    '        helpForm.FormBorderStyle = FormBorderStyle.FixedDialog
    '        helpForm.MaximizeBox = False
    '        helpForm.MinimizeBox = False

    '        Dim lblToggleOutput As New Label() With {
    '            .Text = "To generate the required *.out files in PWC 3, please enabled 'Optional Output' by clicking 'More Tabs -> Toggle More Outputs'" & vbCrLf,
    '            .Dock = DockStyle.Top,
    '            .Height = 60,
    '            .Padding = New Padding(15, 10, 15, 5),
    '            .Font = New Drawing.Font("Segoe UI", 11.0!, Drawing.FontStyle.Regular)
    '        }

    '        Dim lblInfo As New Label() With {
    '            .Text = "Then check the box 'Print daily waterbody output ...'  before running your simulation.",
    '            .Dock = DockStyle.Top,
    '            .Height = 60,
    '            .Padding = New Padding(15, 10, 15, 5),
    '            .Font = New Drawing.Font("Segoe UI", 11.0!, Drawing.FontStyle.Regular)
    '        }

    '        Dim picToggleOutput As New PictureBox() With {
    '            .Dock = DockStyle.Fill,
    '            .SizeMode = PictureBoxSizeMode.Zoom,
    '            .Padding = New Padding(10)
    '        }
    '        picToggleOutput.Load("ToggleOutput.png")

    '        Dim picScreenshot As New PictureBox() With {
    '            .Dock = DockStyle.Fill,
    '            .SizeMode = PictureBoxSizeMode.Zoom,
    '            .Padding = New Padding(10)
    '        }
    '        picScreenshot.Load("HelpCreateOutFiles.png")

    '        Dim btnClose As New Button() With {
    '            .Text = "OK",
    '            .DialogResult = DialogResult.OK,
    '            .Dock = DockStyle.Bottom,
    '            .Height = 40,
    '            .Font = New Drawing.Font("Segoe UI", 11.0!, Drawing.FontStyle.Bold)
    '        }

    '        helpForm.Controls.Add(lblToggleOutput)
    '        helpForm.Controls.Add(picToggleOutput)
    '        helpForm.Controls.Add(lblInfo)
    '        helpForm.Controls.Add(picScreenshot)

    '        helpForm.Controls.Add(btnClose)

    '        helpForm.ShowDialog(Me)
    '    End Using
    'End Sub

    ' --- Control & View State Events ---

    Private Sub OnShowDetailChart_CheckedChanged(sender As Object, e As EventArgs)
        If chartSplitContainer IsNot Nothing Then
            chartSplitContainer.Panel2Collapsed = Not chkShowDetailChart.Checked
            SetBoxPlotVisible(Not chkShowDetailChart.Checked)
        End If
    End Sub

    Private Sub OnSourceRadioButton_CheckedChanged(sender As Object, e As EventArgs)
        Dim rb As RadioButton = CType(sender, RadioButton)
        If rb.Checked Then
            If rb Is rbParent Then currentSourceKey = "Parent"
            If rb Is rbDaughter Then currentSourceKey = "Daughter"
            If rb Is rbGranddaughter Then currentSourceKey = "Granddaughter"

            DisplayCurrentSource()
        End If
    End Sub

    ' --- Drag & Drop Events ---

    Private Sub Form_DragEnter(sender As Object, e As DragEventArgs) Handles MyBase.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub Form_DragDrop(sender As Object, e As DragEventArgs) Handles MyBase.DragDrop
        Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())

        If files IsNot Nothing AndAlso files.Length > 0 Then
            ProcessSelectedFile(files(0))
        End If
    End Sub

    ' --- Chart Interaction Events ---

    Private Sub chartMain_MouseMove(sender As Object, e As MouseEventArgs) Handles chartMain.MouseMove

        Dim result As HitTestResult = chartMain.HitTest(e.X, e.Y)

        If result IsNot Nothing Then
            Dim isOverBoxPlot As Boolean = False

            ' Prüfen, ob die Maus über der BoxPlotArea, den Serien oder dem Title ist
            If (result.ChartArea IsNot Nothing AndAlso result.ChartArea.Name = "BoxPlotArea") OrElse
           (result.Series IsNot Nothing AndAlso (result.Series.Name = "BoxPlotSeries" OrElse result.Series.Name = "BoxPlotPoints")) OrElse
           (result.Object IsNot Nothing AndAlso TypeOf result.Object Is Title AndAlso CType(result.Object, Title).Name = "BoxPlotTitle") Then

                isOverBoxPlot = True
            End If

            ' Cursor & Tooltip anpassen
            If isOverBoxPlot Then
                chartMain.Cursor = Cursors.Hand ' Hand-Symbol zur Verdeutlichung der Klickbarkeit

                If lastHoveredArea <> "BoxPlotArea" Then
                    chartToolTip.SetToolTip(chartMain, "Click to switch Surface Water <-> Benthic System")
                    lastHoveredArea = "BoxPlotArea"
                End If
            Else
                If lastHoveredArea = "BoxPlotArea" Then
                    chartMain.Cursor = Cursors.Default
                    chartToolTip.RemoveAll()
                    lastHoveredArea = ""
                End If
            End If

            If result.ChartElementType = ChartElementType.DataPoint Then
                chartMain.Cursor = Cursors.Hand
            Else
                chartMain.Cursor = Cursors.Default
            End If
        End If


    End Sub

    Private Sub chartDetail_MouseMove(sender As Object, e As MouseEventArgs) Handles chartDetail.MouseMove
        Dim result As HitTestResult = chartDetail.HitTest(e.X, e.Y)
        If result.ChartElementType = ChartElementType.DataPoint Then
            chartDetail.Cursor = Cursors.Hand
        Else
            chartDetail.Cursor = Cursors.Default
        End If
    End Sub

    Private Sub chartMain_MouseClick(sender As Object, e As MouseEventArgs) Handles chartMain.MouseClick
        Dim result As HitTestResult = chartMain.HitTest(e.X, e.Y)
        ' Prüfen, ob mit der linken Maustaste geklickt wurde
        If e.Button = MouseButtons.Left Then
            ' Prüfen, ob der Klick auf die BoxPlotArea oder die Serien des Boxplots ging
            If result IsNot Nothing AndAlso
          (result.ChartArea?.Name = "BoxPlotArea" OrElse
           (result.Series IsNot Nothing AndAlso (result.Series.Name = "BoxPlotSeries" OrElse result.Series.Name = "BoxPlotPoints"))) Then

                ' Modus umschalten (SW <-> Benthic)
                isSwMode = Not isSwMode

                ' Chart mit den neuen Daten & Farben aktualisieren
                PlotMainChart()
                Exit Sub
            End If
        End If


        If result.ChartElementType = ChartElementType.DataPoint Then
            Dim point As DataPoint = result.Series.Points(result.PointIndex)
            If point.Tag IsNot Nothing AndAlso TypeOf point.Tag Is DateTime Then
                Dim realDate As DateTime = CType(point.Tag, DateTime)
                chkShowDetailChart.Checked = True
                UpdateDetailChart(realDate.Year)
            End If
        End If
    End Sub

    Private Sub chartDetail_MouseWheel(sender As Object, e As MouseEventArgs) Handles chartDetail.MouseWheel
        Dim area As ChartArea = chartDetail.ChartAreas("DetailArea")
        Try
            If e.Delta < 0 Then
                area.AxisX.ScaleView.ZoomReset(0)
            ElseIf e.Delta > 0 Then
                Dim xMin As Double = area.AxisX.ScaleView.ViewMinimum
                Dim xMax As Double = area.AxisX.ScaleView.ViewMaximum
                Dim xMouse As Double = area.AxisX.PixelPositionToValue(e.X)

                Dim newSpan As Double = (xMax - xMin) / 2
                area.AxisX.ScaleView.Zoom(xMouse - newSpan / 2, xMouse + newSpan / 2)
            End If
        Catch
        End Try
    End Sub

    Private Sub chartDetail_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles chartDetail.MouseDoubleClick
        chartDetail.ChartAreas("DetailArea").AxisX.ScaleView.ZoomReset(0)
    End Sub



#End Region

End Class