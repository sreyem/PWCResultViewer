

Imports System.Globalization
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Drawing.Imaging
Imports System.Windows.Forms.DataVisualization.Charting

''' <summary>
''' Main form class for viewing PWC (Pesticide Water Calculator) output results,
''' displaying yearly maximum concentrations, box plots, and detailed daily values.
''' </summary>
Public Class FrmPWCResultViewer

#Region " Properties & Data Fields "

    ''' <summary>
    ''' Provides public access to the main chart control for external integration (e.g., Main.vb).
    ''' </summary>
    Public ReadOnly Property MainChartControl As Chart
        Get
            Return chartMain
        End Get
    End Property

    ''' <summary>
    ''' Structure representing daily data entries for surface water and benthic concentrations.
    ''' </summary>
    Private Structure DataRecord
        Public Property DateValue As DateTime
        Public Property WaterCol As Double
        Public Property Benthic As Double
    End Structure

    ''' <summary>
    ''' Stores loaded output file paths mapped by compound source type (Parent, Daughter, Granddaughter).
    ''' </summary>
    Private ReadOnly filePaths As New Dictionary(Of String, String) From {
        {"Parent", Nothing},
        {"Daughter", Nothing},
        {"Granddaughter", Nothing}
    }

    ''' <summary>
    ''' Stores lists of parsed data records mapped by compound source type.
    ''' </summary>
    Private ReadOnly allRecords As New Dictionary(Of String, List(Of DataRecord)) From {
        {"Parent", New List(Of DataRecord)()},
        {"Daughter", New List(Of DataRecord)()},
        {"Granddaughter", New List(Of DataRecord)()}
    }

    ''' <summary>
    ''' Currently selected active data source compound key (default: "Parent").
    ''' </summary>
    Private currentSourceKey As String = "Parent"

    ' UI Elements
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

    ' Chart Titles and Legends
    Private mainTitle As Title
    Private summaryLegend As Legend

    ''' <summary>
    ''' Retrieves dynamic application title and version from application metadata.
    ''' </summary>
    Private ReadOnly Property AppTitle As String
        Get
            ' Retrieve title from metadata; fall back to product name if title is empty
            Dim name As String = My.Application.Info.Title
            If String.IsNullOrEmpty(name) Then
                name = My.Application.Info.ProductName
            End If
            ' Retrieve version string
            Dim version As String = My.Application.Info.Version.ToString()
            Return $"{name} v{version}"
        End Get
    End Property

    ''' <summary>
    ''' Mode switch status for the box plot (True = Surface Water, False = Benthic System).
    ''' </summary>
    Private isSwMode As Boolean = True

    ''' <summary>
    ''' Default color definition for Benthic visualization.
    ''' </summary>
    Private ReadOnly colorBenthic As Drawing.Color = Drawing.Color.SaddleBrown

    ' Tooltip and hover tracking controls
    Private WithEvents chartToolTip As New ToolTip()
    Private lastHoveredArea As String = ""

#End Region

#Region " Form Lifecycle & Initialization "

    ''' <summary>
    ''' Event handler triggered when the form is initially loaded.
    ''' </summary>
    Private Sub Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Set application title and form size
        Me.Text = $"{AppTitle}"
        Me.Size = New Drawing.Size(width:=1400, height:=900)
        Me.AllowDrop = True

        ' Lock window border style to prevent manual resizing by mouse
        Me.FormBorderStyle = FormBorderStyle.FixedSingle

        ' Disable maximize box to enforce fixed layout size
        Me.MaximizeBox = False

        ' Set minimum and maximum window boundaries to current size
        Me.MinimumSize = Me.Size
        Me.MaximumSize = Me.Size

        ' Initialize UI components and layout
        SetupMenuAndStatusStrip()
        SetupRadioButtons()
        SetupLayout()
        SetupMainChart()
        SetupDetailChart()

        ' Check and evaluate command-line parameters
        CheckCMDArgs()
    End Sub

    ''' <summary>
    ''' Parses and executes command-line arguments for automated or batch operations.
    ''' </summary>
    Private Sub CheckCMDArgs()
        ' Retrieve command line argument array
        Dim args() As String = Environment.GetCommandLineArgs()

        ' Parse second argument (UI display flag - defaults to False)
        Dim showUI As Boolean = False
        If args.Length > 2 Then
            Boolean.TryParse(args(2), showUI)
        End If

        ' Verify if a target path argument is passed (args(0) is the executable path itself)
        If args.Length > 1 Then
            ' Clean quotes from argument path string
            Dim inputPath As String = args(1).Trim(""""c)

            ' Case 1: Input path points to a directory
            If Directory.Exists(inputPath) Then
                ' Execute batch process for directory
                ProcessDirectoryBatch(inputPath)

                ' Close and release application resources upon batch completion
                Me.Close()
                Me.Dispose()
                End

                ' Case 2: Input path points to a single file
            ElseIf File.Exists(inputPath) Then
                ' Process single file, export graph as GIF, and exit application
                ProcessSelectedFile(inputPath)
                Dim gifPath As String = Path.ChangeExtension(inputPath, ".gif")
                SaveChartToGif(gifPath)
                Me.Close()
                Me.Dispose()
                End
            End If
        End If
    End Sub

#Region " Batch Mode Operations "

    ''' <summary>
    ''' Processes all *.out and *.csv files in a specified directory and exports charts to GIF format.
    ''' </summary>
    ''' <param name="directoryPath">Target directory path to scan for output files.</param>
    Public Sub ProcessDirectoryBatch(directoryPath As String)
        ' Verify directory existence
        If Not Directory.Exists(directoryPath) Then
            Console.WriteLine($"Directory not found: {directoryPath}")
            Exit Sub
        End If

        ' Fetch all eligible files matching .out or .csv extensions
        Dim files = Directory.GetFiles(directoryPath, "*.*") _
                             .Where(Function(f) f.EndsWith(".out", StringComparison.OrdinalIgnoreCase) OrElse
                                                f.EndsWith(".csv", StringComparison.OrdinalIgnoreCase)) _
                             .ToList()

        ' Check if any matching files were retrieved
        If files.Count = 0 Then
            Console.WriteLine($"No .out or .csv files found in directory: {directoryPath}")
            Exit Sub
        End If

        ' Iterate through each discovered file and export output GIF
        For Each file In files
            ProcessSelectedFile(file)
            Dim gifPath As String = Path.ChangeExtension(file, ".gif")
            SaveChartToGif(gifPath)
        Next
    End Sub

    ''' <summary>
    ''' Saves the main chart control as a GIF image file at the specified output path.
    ''' </summary>
    ''' <param name="outputPath">Target destination file path for the GIF image.</param>
    Private Sub SaveChartToGif(outputPath As String)
        ' Prevent exporting if main chart dimension properties are invalid
        If chartMain.Width <= 0 OrElse chartMain.Height <= 0 Then Exit Sub

        ' Create bitmap instance and render chart content onto graphic canvas
        Using bmp As New Bitmap(chartMain.Width, chartMain.Height)
            Using g As Graphics = Graphics.FromImage(bmp)
                g.Clear(Color.White)
                Dim bmpMain As New Bitmap(chartMain.Width, chartMain.Height)
                chartMain.DrawToBitmap(bmpMain, New Rectangle(0, 0, chartMain.Width, chartMain.Height))
                g.DrawImage(bmpMain, 0, 0)
            End Using
            ' Save generated bitmap image using GIF format
            bmp.Save(outputPath, ImageFormat.Gif)
        End Using
    End Sub

#End Region

    ''' <summary>
    ''' Resets application state, clearing data containers, chart series, and active UI elements.
    ''' </summary>
    Private Sub ResetAppData()
        ' Reset file path values in tracking dictionary
        For Each key In filePaths.Keys.ToList()
            filePaths(key) = Nothing
        Next

        ' Clear all loaded records from internal storage
        For Each key In allRecords.Keys.ToList()
            allRecords(key).Clear()
        Next

        ' Temporarily detach radio button event handlers during UI reset
        RemoveHandler rbParent.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged
        RemoveHandler rbDaughter.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged
        RemoveHandler rbGranddaughter.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged

        ' Reset radio button controls state
        rbParent.Enabled = False
        rbDaughter.Enabled = False
        rbGranddaughter.Enabled = False
        rbParent.Checked = True
        currentSourceKey = "Parent"

        ' Re-attach event handlers for radio buttons
        AddHandler rbParent.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged
        AddHandler rbDaughter.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged
        AddHandler rbGranddaughter.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged

        ' Clear chart series data points
        chartMain.Series("Surface Water Max").Points.Clear()
        chartMain.Series("Benthic Max").Points.Clear()
        If chartMain.Series.IndexOf("BoxPlotSeries") >= 0 Then chartMain.Series("BoxPlotSeries").Points.Clear()
        If chartMain.Series.IndexOf("BoxPlotPoints") >= 0 Then chartMain.Series("BoxPlotPoints").Points.Clear()
        chartDetail.Series("Surface Water (Daily)").Points.Clear()
        chartDetail.Series("Benthic (Daily)").Points.Clear()
        summaryLegend.CustomItems.Clear()

        ' Reset main title text
        If mainTitle IsNot Nothing Then
            mainTitle.Text = "Yearly max concentrations of Parent"
        End If

        ' Reset form title and status label text
        Me.Text = $"{AppTitle}"
        lblStatusFilePath.Text = "Drag & drop PWC file or File -> Open"

        ' Hide box plot title if it exists
        If chartMain.Titles.FindByName("BoxPlotTitle") IsNot Nothing Then
            chartMain.Titles("BoxPlotTitle").Visible = False
        End If
    End Sub

#End Region

#Region " Setup Form & Controls "

    ''' <summary>
    ''' Builds and configures the main menu bar structure and status strip controls.
    ''' </summary>
    Private Sub SetupMenuAndStatusStrip()
        ' Initialize main menu bar control
        menuStrip1 = New MenuStrip() With {.Font = New Drawing.Font("Segoe UI", 12.0!)}

        ' Construct 'File' menu header and items
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

        ' Populate 'File' dropdown items
        menuFile.DropDownItems.Add(menuOpen)
        menuFile.DropDownItems.Add(menuProcessFolder)
        menuFile.DropDownItems.Add(menuSave)
        menuFile.DropDownItems.Add(menuSaveYearMaxCsv)
        menuFile.DropDownItems.Add(menuReset)
        menuFile.DropDownItems.Add(New ToolStripSeparator())
        menuFile.DropDownItems.Add(menuExit)

        ' Construct 'Help' menu header and items
        Dim menuHelp As New ToolStripMenuItem("Help") With {.Font = New Drawing.Font("Segoe UI", 12.0!)}
        Dim menuHowTo As New ToolStripMenuItem("How to generate PWC 3.XXXX *.out files...", Nothing, AddressOf MenuHelp_Click) With {
            .Font = New Drawing.Font("Segoe UI", 12.0!)
        }

        Dim menuShowDetailsHelp As New ToolStripMenuItem("Help on 'Show Details'", Nothing, AddressOf MenuShowDetailsHelp_Click) With {
            .Font = New Drawing.Font("Segoe UI", 12.0!)
        }
        menuHelp.DropDownItems.Add(menuHowTo)
        menuHelp.DropDownItems.Add(New ToolStripSeparator())

        ' Register top-level menus to menu strip
        menuStrip1.Items.Add(menuFile)
        menuStrip1.Items.Add(menuHelp)

        Me.MainMenuStrip = menuStrip1
        Me.Controls.Add(menuStrip1)

        ' Initialize Status Strip control
        statusStrip1 = New StatusStrip()
        lblStatusFilePath = New ToolStripStatusLabel("Drag & drop PWC file or File -> Open") With {
            .Spring = True, .TextAlign = Drawing.ContentAlignment.MiddleLeft, .Font = New Drawing.Font("Segoe UI", 12.0!)
        }
        statusStrip1.Items.Add(lblStatusFilePath)
        Me.Controls.Add(statusStrip1)
    End Sub

    ''' <summary>
    ''' Handles user menu request to export calculated annual maximum concentrations to a CSV file.
    ''' </summary>
    Private Sub MenuSaveYearMaxCsv_Click(sender As Object, e As EventArgs)
        Dim activeFile As String = filePaths(currentSourceKey)

        ' Validate loaded data availability prior to saving
        If String.IsNullOrEmpty(activeFile) OrElse Not allRecords.ContainsKey(currentSourceKey) OrElse allRecords(currentSourceKey).Count = 0 Then
            MessageBox.Show("No data to save.", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        ' Setup default output filename and directory path
        Dim defaultFileName As String = Path.GetFileNameWithoutExtension(activeFile) & "_YearMax.csv"
        Dim initialDirectory As String = Path.GetDirectoryName(activeFile)

        ' Show save file dialog to prompt user for target path
        Using sfd As New SaveFileDialog()
            sfd.Filter = "CSV Files (*.csv)|*.csv"
            sfd.FileName = defaultFileName
            sfd.InitialDirectory = initialDirectory

            If sfd.ShowDialog() = DialogResult.OK Then
                SaveYearMaxCsv(sfd.FileName)
            End If
        End Using
    End Sub

    ''' <summary>
    ''' Exports calculated annual maximum surface water and benthic values to a CSV file.
    ''' </summary>
    ''' <param name="outputPath">Output file destination path.</param>
    Private Sub SaveYearMaxCsv(outputPath As String)
        Try
            Dim records = allRecords(currentSourceKey)
            If records.Count = 0 Then Exit Sub

            ' Group records by calendar year ordered sequentially
            Dim yearlyGroups = records.GroupBy(Function(r) r.DateValue.Year).OrderBy(Function(g) g.Key).ToList()

            ' Write output CSV data using invariant culture formatting
            Using writer As New StreamWriter(outputPath, False, System.Text.Encoding.UTF8)
                ' Header line: Year, Sw (ppb), Benthic (ppb)
                writer.WriteLine("Year, Sw (ppb),Benthic (ppb)")

                For Each group In yearlyGroups
                    Dim yr As Integer = group.Key
                    Dim maxSw As Double = group.Max(Function(r) r.WaterCol)
                    Dim maxBenthic As Double = group.Max(Function(r) r.Benthic)

                    ' Format and write values using period decimal separator
                    writer.WriteLine(String.Format(CultureInfo.InvariantCulture, "{0}, {1:F6}, {2:F6}", yr, maxSw, maxBenthic))
                Next
            End Using

            MessageBox.Show($"File saved at:{vbCrLf}{outputPath}", "Save Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show($"Error while saving the CSV file:{vbCrLf}{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Displays help information dialog for the 'Show Details' functionality.
    ''' </summary>
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

    ''' <summary>
    ''' Creates and places input data source selection radio buttons and options panel.
    ''' </summary>
    Private Sub SetupRadioButtons()
        ' Initialize container panel
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

        ' Parent Radio Button
        rbParent = New RadioButton() With {
            .Text = "Parent",
            .Location = New Drawing.Point(140, 8),
            .AutoSize = True,
            .Checked = True,
            .Enabled = False,
            .Font = New Drawing.Font("Segoe UI", 12.0!)
        }

        ' Daughter Radio Button
        rbDaughter = New RadioButton() With {
            .Text = "Daughter",
            .Location = New Drawing.Point(260, 8),
            .AutoSize = True,
            .Enabled = False,
            .Font = New Drawing.Font("Segoe UI", 12.0!)
        }

        ' Granddaughter Radio Button
        rbGranddaughter = New RadioButton() With {
            .Text = "Granddaughter",
            .Location = New Drawing.Point(390, 8),
            .AutoSize = True,
            .Enabled = False,
            .Font = New Drawing.Font("Segoe UI", 12.0!)
        }

        ' Show Details CheckBox
        chkShowDetailChart = New CheckBox() With {
            .Text = "Show Details",
            .Location = New Drawing.Point(560, 8),
            .AutoSize = True,
            .Checked = False,
            .Font = New Drawing.Font("Segoe UI", 12.0!)
        }

        ' Register event listeners for controls
        AddHandler rbParent.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged
        AddHandler rbDaughter.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged
        AddHandler rbGranddaughter.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged
        AddHandler chkShowDetailChart.CheckedChanged, AddressOf OnShowDetailChart_CheckedChanged

        ' Add controls to container panel
        pnlRadioContainer.Controls.Add(lblSelect)
        pnlRadioContainer.Controls.Add(rbParent)
        pnlRadioContainer.Controls.Add(rbDaughter)
        pnlRadioContainer.Controls.Add(rbGranddaughter)
        pnlRadioContainer.Controls.Add(chkShowDetailChart)

        Me.Controls.Add(pnlRadioContainer)
    End Sub

    ''' <summary>
    ''' Configures the split container layout structure for primary and detail chart controls.
    ''' </summary>
    Private Sub SetupLayout()
        chartSplitContainer = New SplitContainer() With {
            .Dock = DockStyle.Fill,
            .Orientation = Orientation.Horizontal,
            .SplitterDistance = 400,
            .Panel2Collapsed = True
        }
        Me.Controls.Add(chartSplitContainer)

        ' Adjust z-order positioning
        chartSplitContainer.BringToFront()
        pnlRadioContainer.BringToFront()
        menuStrip1.BringToFront()

        ' Instantiate and attach main chart control to top panel
        chartMain = New Chart() With {.Dock = DockStyle.Fill}
        chartSplitContainer.Panel1.Controls.Add(chartMain)

        ' Instantiate and attach detail chart control to bottom panel
        chartDetail = New Chart() With {.Dock = DockStyle.Fill}
        chartSplitContainer.Panel2.Controls.Add(chartDetail)
    End Sub

    ''' <summary>
    ''' Sets up chart areas, axes, titles, legends, and series for the main chart control.
    ''' </summary>
    Private Sub SetupMainChart()
        Dim area As New ChartArea("MainArea")

        ' Configure X-Axis styling
        area.AxisX.Title = "Relative Year"
        area.AxisX.LabelStyle.Font = New Drawing.Font("Segoe UI", 12.0!)
        area.AxisX.TitleFont = New Drawing.Font("Segoe UI", 12.0!, Drawing.FontStyle.Bold)
        area.AxisX.Interval = 1
        area.AxisX.IsMarginVisible = True
        area.AxisX.MajorGrid.LineColor = Drawing.Color.LightGray

        ' Configure Y-Axis styling
        area.AxisY.Title = "Yearly Max. Concentration (ppb)"
        area.AxisY.LabelStyle.Font = New Drawing.Font("Segoe UI", 12.0!)
        area.AxisY.TitleFont = New Drawing.Font("Segoe UI", 12.0!, Drawing.FontStyle.Bold)
        area.AxisY.LabelStyle.Format = "0.00"
        area.AxisY.MajorGrid.LineColor = Drawing.Color.LightGray
        chartMain.ChartAreas.Add(area)

        ' Clear and set main chart title
        chartMain.Titles.Clear()
        mainTitle = New Title() With {
            .Name = "MainTitle",
            .Text = "Yearly max concentrations of Parent",
            .Font = New Drawing.Font("Segoe UI", 14.0!, Drawing.FontStyle.Bold),
            .Docking = Docking.Top
        }
        chartMain.Titles.Add(mainTitle)

        ' Configure summary result overview legend
        summaryLegend = New Legend("SummaryLegend") With {
            .Docking = Docking.Right,
            .Alignment = Drawing.StringAlignment.Near,
            .Title = "Result Overview",
            .TitleFont = New Drawing.Font("Segoe UI", 14.0!, Drawing.FontStyle.Bold),
            .BackColor = Drawing.Color.FromArgb(245, 245, 245),
            .BorderColor = Drawing.Color.LightGray,
            .TextWrapThreshold = 0,
            .IsTextAutoFit = True,
            .AutoFitMinFontSize = 8
        }
        chartMain.Legends.Add(summaryLegend)

        ' Surface Water maximum points series configuration
        Dim seriesWater As New Series("Surface Water Max") With {
            .ChartType = SeriesChartType.Point,
            .XValueType = ChartValueType.Double,
            .MarkerStyle = MarkerStyle.Circle,
            .MarkerSize = 10,
            .Color = Drawing.Color.RoyalBlue,
            .IsVisibleInLegend = False
        }

        ' Benthic maximum points series configuration
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

        ' Configure Title for BoxPlot chart area
        Dim titleBoxPlot As New Title() With {
            .Name = "BoxPlotTitle",
            .Text = "Surface Water Yearly Max (ppb)",
            .DockedToChartArea = "BoxPlotArea",
            .IsDockedInsideChartArea = False,
            .Font = New Drawing.Font("Segoe UI", 14.0!, Drawing.FontStyle.Bold),
            .ForeColor = Drawing.Color.Black,
            .Visible = False
        }

        ' Remove duplicate title instance if present
        If chartMain.Titles.FindByName("BoxPlotTitle") IsNot Nothing Then
            chartMain.Titles.Remove(chartMain.Titles("BoxPlotTitle"))
        End If
        chartMain.Titles.Add(titleBoxPlot)

        ' Configure ChartArea for BoxPlot overlay
        Dim areaBox As New ChartArea("BoxPlotArea")
        areaBox.Position.Auto = False
        areaBox.Position.X = 72
        areaBox.Position.Y = 55
        areaBox.Position.Width = 26
        areaBox.Position.Height = 40

        ' Hide axes and margins for clean overlay appearance
        areaBox.AxisX.Enabled = AxisEnabled.False
        areaBox.AxisY.Enabled = AxisEnabled.False
        areaBox.InnerPlotPosition.Auto = False
        areaBox.InnerPlotPosition.X = 0
        areaBox.InnerPlotPosition.Y = 0
        areaBox.InnerPlotPosition.Width = 100
        areaBox.InnerPlotPosition.Height = 100

        chartMain.ChartAreas.Add(areaBox)

        ' Series 1: BoxPlot structure settings
        Dim seriesBoxPlot As New Series("BoxPlotSeries") With {
            .ChartType = SeriesChartType.BoxPlot,
            .ChartArea = "BoxPlotArea",
            .IsVisibleInLegend = False,
            .Color = Drawing.Color.FromArgb(0, 255, 255, 255),
            .BorderColor = Drawing.Color.ForestGreen,
            .BorderWidth = 2
        }

        seriesBoxPlot("BoxPlotTransparentColor") = "Transparent"
        seriesBoxPlot("BoxPlotWhiskerPercentile") = "0"
        seriesBoxPlot("BoxPlotShowMedian") = "True"
        seriesBoxPlot("BoxPlotShowAverage") = "False"

        ' Series 2: Individual data points series
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

    ''' <summary>
    ''' Configures chart area, series, and interaction parameters for the detailed daily chart control.
    ''' </summary>
    Private Sub SetupDetailChart()
        Dim area As New ChartArea("DetailArea")

        ' Configure X-Axis properties
        area.AxisX.Title = "Month (Daily values for selected year)"
        area.AxisX.LabelStyle.Font = New Drawing.Font("Segoe UI", 12.0!)
        area.AxisX.TitleFont = New Drawing.Font("Segoe UI", 12.0!, Drawing.FontStyle.Bold)
        area.AxisX.LabelStyle.Format = "MMM"
        area.AxisX.Interval = 1
        area.AxisX.IntervalType = DateTimeIntervalType.Months
        area.AxisX.MajorGrid.LineColor = Drawing.Color.LightGray

        ' Configure Y-Axis properties
        area.AxisY.Title = "Daily Value (ppb)"
        area.AxisY.LabelStyle.Font = New Drawing.Font("Segoe UI", 12.0!)
        area.AxisY.TitleFont = New Drawing.Font("Segoe UI", 12.0!, Drawing.FontStyle.Bold)
        area.AxisY.LabelStyle.Format = "0.00"
        area.AxisY.MajorGrid.LineColor = Drawing.Color.LightGray

        ' Enable interactive zooming and scrolling features
        area.CursorX.IsUserSelectionEnabled = True
        area.CursorX.IsUserEnabled = True
        area.AxisX.ScaleView.Zoomable = True
        area.AxisX.ScrollBar.IsPositionedInside = False
        area.AxisX.ScrollBar.Size = 12
        area.AxisX.ScrollBar.ButtonStyle = ScrollBarButtonStyles.All

        chartDetail.ChartAreas.Add(area)

        ' Daily surface water concentration line series
        Dim seriesWater As New Series("Surface Water (Daily)") With {
            .ChartType = SeriesChartType.Line,
            .XValueType = ChartValueType.Date,
            .BorderWidth = 2,
            .Color = Drawing.Color.RoyalBlue
        }

        ' Daily benthic concentration line series
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

    ''' <summary>
    ''' Inspects, branches, and processes a user-selected file path according to file type specifications.
    ''' </summary>
    ''' <param name="filePath">Target file path string to process.</param>
    Public Sub ProcessSelectedFile(filePath As String)
        ' Clear previous state data
        ResetAppData()

        Dim fileName As String = Path.GetFileName(filePath)

        ' Determine file format type (Legacy CSV vs. standard PWC3 output)
        If IsLegacyFile(fileName) Then
            ProcessLegacyFile(filePath)
        Else
            ProcessPwc3File(filePath)
        End If

        ' Render data views for current active target source
        DisplayCurrentSource()

        ' Enable box plot title display
        If chartMain.Titles.FindByName("BoxPlotTitle") IsNot Nothing Then
            chartMain.Titles("BoxPlotTitle").Visible = True
        End If
    End Sub

    ''' <summary>
    ''' Checks if the given filename corresponds to legacy PWC CSV format patterns.
    ''' </summary>
    ''' <param name="fileName">Name of the target file.</param>
    ''' <returns>True if legacy CSV file pattern matches; otherwise False.</returns>
    Private Function IsLegacyFile(fileName As String) As Boolean
        Return fileName.EndsWith("_daily.csv", StringComparison.OrdinalIgnoreCase) OrElse
               fileName.IndexOf("_daily", StringComparison.OrdinalIgnoreCase) >= 0
    End Function

    ' ==========================================
    ' ROUTINE 1: Legacy PWC Processing (*_daily.csv)
    ' ==========================================

    ''' <summary>
    ''' Handles file discovery and data ingestion for legacy PWC CSV output files.
    ''' </summary>
    ''' <param name="filePath">Primary legacy file path.</param>
    Private Sub ProcessLegacyFile(filePath As String)
        Dim directoryPath As String = Path.GetDirectoryName(filePath)
        Dim fileName As String = Path.GetFileName(filePath)

        ' Determine target compound key based on file naming convention
        Dim targetSourceKey As String = "Parent"
        If fileName.IndexOf("_Degradate1_daily.csv", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
           fileName.IndexOf("Degradate1", StringComparison.OrdinalIgnoreCase) >= 0 Then
            targetSourceKey = "Daughter"
        ElseIf fileName.IndexOf("_Degradate2_daily.csv", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               fileName.IndexOf("Degradate2", StringComparison.OrdinalIgnoreCase) >= 0 Then
            targetSourceKey = "Granddaughter"
        End If

        ' Strip suffixes to identify common file prefix name
        Dim prefix As String = fileName
        Dim legacySuffixes() As String = {"_Parent_daily.csv", "_Degradate1_daily.csv", "_Degradate2_daily.csv", "_daily.csv"}

        For Each suffix In legacySuffixes
            If prefix.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) Then
                prefix = prefix.Substring(0, prefix.Length - suffix.Length)
                Exit For
            End If
        Next

        ' Locate matching companion compound legacy files within target directory
        filePaths("Parent") = FindMatchingFile(directoryPath, prefix, "_Parent_daily.csv")
        filePaths("Daughter") = FindMatchingFile(directoryPath, prefix, "_Degradate1_daily.csv")
        filePaths("Granddaughter") = FindMatchingFile(directoryPath, prefix, "_Degradate2_daily.csv")

        ' Fallback assignment if parent compound path is unassigned
        If String.IsNullOrEmpty(filePaths("Parent")) AndAlso targetSourceKey = "Parent" Then
            filePaths("Parent") = filePath
        End If

        ' Parse records from all discovered valid file paths
        For Each key In filePaths.Keys.ToList()
            If Not String.IsNullOrEmpty(filePaths(key)) Then
                LoadLegacyCsvFile(filePaths(key), allRecords(key))
            End If
        Next

        ' Detach event handlers during option state configuration
        RemoveHandler rbParent.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged
        RemoveHandler rbDaughter.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged
        RemoveHandler rbGranddaughter.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged

        ' Enable radio buttons where data was parsed
        rbParent.Enabled = (allRecords("Parent").Count > 0)
        rbDaughter.Enabled = (allRecords("Daughter").Count > 0)
        rbGranddaughter.Enabled = (allRecords("Granddaughter").Count > 0)

        ' Set selected radio button to match target source key
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

        ' Re-attach event listeners
        AddHandler rbParent.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged
        AddHandler rbDaughter.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged
        AddHandler rbGranddaughter.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged

        ' Update status bar label display
        lblStatusFilePath.Text = $"Legacy File: {fileName}  |  Folder: {directoryPath}  |  Prefix: {prefix}"
    End Sub

    ''' <summary>
    ''' Reads and parses records from a single legacy CSV output file into target record container.
    ''' </summary>
    ''' <param name="filePath">Target legacy CSV file path.</param>
    ''' <param name="targetList">Data record list destination container.</param>
    Private Sub LoadLegacyCsvFile(filePath As String, targetList As List(Of DataRecord))
        Try
            Dim currentDate As New DateTime(1980, 1, 1)
            Const multiplier As Double = 1000000.0

            Using reader As New StreamReader(filePath)
                ' Skip top 5 metadata header lines
                For i As Integer = 1 To 5
                    If reader.EndOfStream Then Exit Sub
                    reader.ReadLine()
                Next

                ' Parse daily records line by line from line 6 onwards
                While Not reader.EndOfStream
                    Dim line As String = reader.ReadLine()
                    If String.IsNullOrWhiteSpace(line) Then Continue While

                    Dim parts() As String = line.TrimEnd(","c).Split(","c)

                    If parts.Length >= 3 Then
                        Dim waterVal As Double
                        Dim benthicVal As Double

                        ' Parse numerical values using invariant culture
                        If Double.TryParse(parts(1).Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, waterVal) AndAlso
                           Double.TryParse(parts(2).Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, benthicVal) Then

                            targetList.Add(New DataRecord With {
                                .DateValue = currentDate,
                                .WaterCol = waterVal * multiplier,
                                .Benthic = benthicVal * multiplier
                            })

                            ' Increment calendar date sequentially by 1 day
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
    ' ROUTINE 2: Standard PWC 3 Processing (*.out)
    ' ==========================================

    ''' <summary>
    ''' Handles file discovery and data parsing for standard PWC 3 *.out output files.
    ''' </summary>
    ''' <param name="filePath">Target PWC 3 file path.</param>
    Private Sub ProcessPwc3File(filePath As String)
        Dim directoryPath As String = Path.GetDirectoryName(filePath)
        Dim fileName As String = Path.GetFileName(filePath)

        ' Identify target compound key based on file naming structure
        Dim targetSourceKey As String = "Parent"
        If fileName.IndexOf("daughter", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
           fileName.IndexOf("deg1", StringComparison.OrdinalIgnoreCase) >= 0 Then
            targetSourceKey = "Daughter"
        ElseIf fileName.IndexOf("granddaughter", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               fileName.IndexOf("deg2", StringComparison.OrdinalIgnoreCase) >= 0 Then
            targetSourceKey = "Granddaughter"
        End If

        ' Clean known file suffixes to establish shared file prefix name
        Dim prefix As String = fileName
        Dim knownSuffixes() As String = {"_parent_Pond.out", "_daughter_Pond.out", "_granddaughter_Pond.out", "_summary.txt", "_summary_Deg1.txt", "_summary_Deg2.txt"}

        For Each suffix In knownSuffixes
            If prefix.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) Then
                prefix = prefix.Substring(0, prefix.Length - suffix.Length)
                Exit For
            End If
        Next

        ' Scan directory for related compound standard output files
        filePaths("Parent") = FindMatchingFile(directoryPath, prefix, "_parent_Pond.out")
        filePaths("Daughter") = FindMatchingFile(directoryPath, prefix, "_daughter_Pond.out")
        filePaths("Granddaughter") = FindMatchingFile(directoryPath, prefix, "_granddaughter_Pond.out")

        ' Parse records for each detected output file path
        For Each key In filePaths.Keys.ToList()
            If Not String.IsNullOrEmpty(filePaths(key)) Then
                LoadPwc3OutFile(filePaths(key), allRecords(key))
            End If
        Next

        ' Detach event handlers prior to updating state
        RemoveHandler rbParent.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged
        RemoveHandler rbDaughter.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged
        RemoveHandler rbGranddaughter.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged

        ' Configure control availability based on loaded data
        rbParent.Enabled = (allRecords("Parent").Count > 0)
        rbDaughter.Enabled = (allRecords("Daughter").Count > 0)
        rbGranddaughter.Enabled = (allRecords("Granddaughter").Count > 0)

        ' Set active compound source radio button selection
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

        ' Re-attach event listeners
        AddHandler rbParent.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged
        AddHandler rbDaughter.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged
        AddHandler rbGranddaughter.CheckedChanged, AddressOf OnSourceRadioButton_CheckedChanged

        ' Update status strip path label
        lblStatusFilePath.Text = $"PWC3 File: {fileName}  |  Folder: {directoryPath}  |  Prefix: {prefix}"
    End Sub

    ''' <summary>
    ''' Loads and parses daily water column and benthic records from a PWC 3 *.out file.
    ''' </summary>
    ''' <param name="filePath">Path to target standard PWC 3 output file.</param>
    ''' <param name="targetList">Data record destination collection.</param>
    Private Sub LoadPwc3OutFile(filePath As String, targetList As List(Of DataRecord))
        Try
            Dim currentDate As New DateTime(1980, 1, 1)
            Const multiplier As Double = 1000000.0

            Using reader As New StreamReader(filePath)
                ' Read header line
                Dim header As String = reader.ReadLine()

                ' Iterate through data lines
                While Not reader.EndOfStream
                    Dim line As String = reader.ReadLine()
                    If String.IsNullOrWhiteSpace(line) Then Continue While

                    Dim parts() As String = line.TrimEnd(","c).Split(","c)

                    If parts.Length >= 3 Then
                        Dim waterVal As Double
                        Dim benthicVal As Double

                        ' Parse numerical concentrations using invariant format rules
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

    ''' <summary>
    ''' Searches directory for files matching specified prefix and suffix patterns.
    ''' </summary>
    ''' <param name="dir">Directory path to search within.</param>
    ''' <param name="prefix">File prefix pattern string.</param>
    ''' <param name="suffix">File suffix pattern string.</param>
    ''' <returns>Matching full file path string if found; otherwise Nothing.</returns>
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

    ''' <summary>
    ''' Loads associated summary metadata for the current active source and updates the summary legend.
    ''' </summary>
    ''' <param name="mainFilePath">Loaded primary data file path.</param>
    Private Sub LoadSummaryFileForCurrentSource(mainFilePath As String)
        summaryLegend.CustomItems.Clear()

        ' Build visual Surface Water header item for summary legend
        Dim itemWater As New LegendItem()
        itemWater.Cells.Add(New LegendCell(LegendCellType.Text, "Concentration in Surface Water   (ppb)") With {
            .Font = New Drawing.Font("Segoe UI", 11.0!, Drawing.FontStyle.Bold),
            .ForeColor = Drawing.Color.RoyalBlue,
            .Alignment = Drawing.ContentAlignment.MiddleLeft
        })

        ' Build visual Benthic System header item for summary legend
        Dim itemBenthic As New LegendItem()
        itemBenthic.Cells.Add(New LegendCell(LegendCellType.Text, "                             Benthic System (ppb)") With {
            .Font = New Drawing.Font("Segoe UI", 11.0!, Drawing.FontStyle.Bold),
            .ForeColor = Drawing.Color.DarkOrange,
            .Alignment = Drawing.ContentAlignment.MiddleLeft
        })

        summaryLegend.CustomItems.Add(itemWater)
        summaryLegend.CustomItems.Add(itemBenthic)

        ' Branch execution based on legacy vs standard PWC 3 file format
        If IsLegacyFile(Path.GetFileName(mainFilePath)) Then
            LoadLegacySummaryFile(mainFilePath)
        Else
            LoadPwc3SummaryFile(mainFilePath)
        End If
    End Sub

#Region " Legacy Summary Processing "

    ''' <summary>
    ''' Parses legacy text summary files (*.txt) and attaches metrics to the result overview legend.
    ''' </summary>
    ''' <param name="mainFilePath">Path to the active legacy data file.</param>
    Private Sub LoadLegacySummaryFile(mainFilePath As String)
        Dim directoryPath As String = Path.GetDirectoryName(mainFilePath)
        Dim fileName As String = Path.GetFileName(mainFilePath)

        ' Strip legacy file suffixes to establish search prefix
        Dim prefix As String = fileName
        Dim legacySuffixes() As String = {"_Parent_daily.csv", "_Degradate1_daily.csv", "_Degradate2_daily.csv", "_daily.csv"}

        For Each suffix In legacySuffixes
            If prefix.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) Then
                prefix = prefix.Substring(0, prefix.Length - suffix.Length)
                Exit For
            End If
        Next

        ' Select target summary suffix according to active compound source key
        Dim summarySuffix As String = "_Parent.txt"
        Select Case currentSourceKey
            Case "Daughter"
                summarySuffix = "_Degradate1.txt"
            Case "Granddaughter"
                summarySuffix = "_Degradate2.txt"
        End Select

        ' Resolve summary file path
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
            ' Parse key-value pairs separated by equal sign '='
            Dim lines() As String = File.ReadAllLines(summaryFilePath)
            Dim dictValues As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            For Each line As String In lines
                ' Strip carriage return line breaks and trim spaces
                Dim trimmedLine As String = line.Replace(vbCr, "").Trim()

                If trimmedLine.Contains("="c) Then
                    ' Split string on first equal sign occurrence
                    Dim parts() As String = trimmedLine.Split(New Char() {"="c}, 2)
                    If parts.Length >= 2 Then
                        Dim key As String = parts(0).Trim()
                        Dim val As String = parts(1).Trim()

                        ' Store unique keys in lookup dictionary
                        If Not dictValues.ContainsKey(key) Then
                            dictValues.Add(key, val)
                        End If
                    End If
                End If
            Next

            AddSummaryLegendItem("--------", "-----------")

            ' Extract surface water average concentrations
            AddLegacyValueToSummary(dictValues, "SW       1-d avg:", "1-d avg 1-in-10.0", "")
            AddLegacyValueToSummary(dictValues, "SW       4-d avg:", "4-d avg 1-in-10.0", "")
            AddLegacyValueToSummary(dictValues, "SW      21-d avg:", "21-d avg 1-in-10.0", "")
            AddLegacyValueToSummary(dictValues, "SW      60-d avg:", "60-d avg 1-in-10.0", "")
            AddLegacyValueToSummary(dictValues, "SW      90-d avg:", "90-d avg 1-in-10.0", "")

            ' Extract benthic pore water average concentrations
            AddLegacyValueToSummary(dictValues, "Benthic  1-d avg:", "Benthic Pore Water 1-d   avg 1-in-10.0", "")
            AddLegacyValueToSummary(dictValues, "        21-d avg:", "Benthic Pore Water 21-d avg 1-in-10.0", "")

            ' Extract entry path percentages
            AddSummaryLegendItem("Entry Paths", "")

            AddLegacyValueToSummary(dictValues, "Runoff :", "Due to Runoff", "%", isPercentage:=True)
            AddLegacyValueToSummary(dictValues, "Erosion:", "Due to Erosion", "%", isPercentage:=True)
            AddLegacyValueToSummary(dictValues, "Drift  :", "Due to Drift", "%", isPercentage:=True)

        Catch ex As Exception
            MessageBox.Show($"Error reading Legacy summary file ({Path.GetFileName(summaryFilePath)}):{vbCrLf}{ex.Message}", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    ''' <summary>
    ''' Helper routine to search legacy key-value dictionary and add parsed metric to summary legend.
    ''' </summary>
    Private Sub AddLegacyValueToSummary(dict As Dictionary(Of String, String), label As String, keyName As String, unit As String, Optional isPercentage As Boolean = False)
        ' Helper function normalizing whitespace character sequences
        Dim CleanString = Function(s As String) Regex.Replace(s.Replace(ChrW(160), " "c), "\s+", " ").Trim()
        Dim targetClean As String = CleanString(keyName)

        ' Search matching key-value pairs in dictionary
        For Each kvp In dict
            Dim keyClean As String = CleanString(kvp.Key)
            Dim valClean As String = CleanString(kvp.Value)

            ' Search Case 1: Search key located left of '=' sign
            If keyClean.IndexOf(targetClean, StringComparison.OrdinalIgnoreCase) >= 0 Then
                ExtractAndAddValue(kvp.Value, label, unit, isPercentage)
                Exit Sub
            End If

            ' Search Case 2: Search key located right of '=' sign
            If valClean.IndexOf(targetClean, StringComparison.OrdinalIgnoreCase) >= 0 Then
                ExtractAndAddValue(kvp.Key, label, unit, isPercentage)
                Exit Sub
            End If
        Next

        ' Append N/A fallback entry if key is absent
        AddSummaryLegendItem(label, "N/A")
    End Sub

    ''' <summary>
    ''' Helper routine parsing numerical values from raw text and displaying formatted summary legend items.
    ''' </summary>
    Private Sub ExtractAndAddValue(rawText As String, label As String, unit As String, isPercentage As Boolean)
        ' Extract leading numerical token string
        Dim firstToken As String = Regex.Split(rawText.Trim(), "\s+")(0).Replace("ppb", "").Replace("%", "").Trim()
        Dim parsedVal As Double

        ' Try parsing numeric double floating-point value
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

    ''' <summary>
    ''' Parses standard PWC 3 summary text files (*_summary.txt) and adds entries to the summary legend.
    ''' </summary>
    ''' <param name="mainFilePath">Path to the active PWC 3 output file.</param>
    Private Sub LoadPwc3SummaryFile(mainFilePath As String)
        Dim directoryPath As String = Path.GetDirectoryName(mainFilePath)
        Dim fileNameWithoutExt As String = Path.GetFileNameWithoutExtension(mainFilePath)

        ' Isolate prefix up to the first underscore
        Dim prefix As String = fileNameWithoutExt
        If prefix.Contains("_") Then
            prefix = prefix.Split("_"c)(0)
        End If

        ' Determine target summary file suffix
        Dim summarySuffix As String = "_summary.txt"
        Select Case currentSourceKey
            Case "Daughter"
                summarySuffix = "_summary_Deg1.txt"
            Case "Granddaughter"
                summarySuffix = "_summary_Deg2.txt"
        End Select

        ' Resolve summary text file destination path
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

            ' Strip pond suffix if present
            Const pondSuffix As String = "_pond"
            If targetRowStart.EndsWith(pondSuffix, StringComparison.OrdinalIgnoreCase) Then
                targetRowStart = targetRowStart.Substring(0, targetRowStart.Length - pondSuffix.Length)
            End If

            ' Adjust degrade compound identifiers in match prefix
            If currentSourceKey = "Daughter" Then
                targetRowStart = targetRowStart.Replace("daughter", "deg1")
            ElseIf currentSourceKey = "Granddaughter" Then
                targetRowStart = targetRowStart.Replace("granddaughter", "deg2")
            End If

            ' Locate matching header line and target data row
            For Each line As String In lines
                Dim trimmedLine As String = line.Trim()

                If trimmedLine.Contains("1-d avg") OrElse trimmedLine.Contains("Runoff") Then
                    headerLine = trimmedLine
                End If

                If trimmedLine.StartsWith(targetRowStart, StringComparison.OrdinalIgnoreCase) Then
                    dataLine = trimmedLine
                End If
            Next

            ' Extract metrics matching table header names
            If headerLine IsNot Nothing AndAlso dataLine IsNot Nothing Then
                AddSummaryLegendItem("--------", "-----------")
                AddSummaryLegendItem("SW       1-d avg:", GetValueFromHeader(headerLine, dataLine, "1-d avg", isPercentage:=False))
                AddSummaryLegendItem("         4-d avg:", GetValueFromHeader(headerLine, dataLine, "4-d avg", isPercentage:=False))
                AddSummaryLegendItem("        21-d avg:", GetValueFromHeader(headerLine, dataLine, "21-d avg", isPercentage:=False))
                AddSummaryLegendItem("Benthic  1-d avg:", GetValueFromHeader(headerLine, dataLine, "B 1-day", isPercentage:=False))
                AddSummaryLegendItem("        21-d avg:", GetValueFromHeader(headerLine, dataLine, "B 21-d avg", isPercentage:=False))
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

    ''' <summary>
    ''' Helper function matching header position to extract specific numerical field values from data row strings.
    ''' </summary>
    Private Function GetValueFromHeader(headerLine As String, dataLine As String, keyName As String, Optional isPercentage As Boolean = False) As String
        Try
            Dim headers As String()
            Dim values As String()

            ' Parse comma-separated or space-delimited text tables
            If headerLine.Contains(","c) Then
                headers = headerLine.Split(","c).Select(Function(s) s.Trim()).ToArray()
                values = dataLine.Split(","c).Select(Function(s) s.Trim()).ToArray()
            Else
                headers = Regex.Split(headerLine.Trim(), "\s{2,}|\t").Select(Function(s) s.Trim()).ToArray()
                values = Regex.Split(dataLine.Trim(), "\s{2,}|\t").Select(Function(s) s.Trim()).ToArray()
            End If

            ' Locate key column index matching target name
            For index As Integer = 0 To headers.Length - 1
                If headers(index).Equals(keyName, StringComparison.OrdinalIgnoreCase) OrElse
                   headers(index).Replace(" ", "").Equals(keyName.Replace(" ", ""), StringComparison.OrdinalIgnoreCase) Then

                    If index < values.Length Then
                        Dim rawVal As String = values(index)
                        Dim parsedVal As Double

                        ' Format extracted numeric values
                        If Double.TryParse(rawVal, NumberStyles.Float, CultureInfo.InvariantCulture, parsedVal) Then

                            If isPercentage Then
                                Dim pctVal As Double = parsedVal * 100.0
                                Return pctVal.ToString("F2", CultureInfo.InvariantCulture) & " %"
                            End If

                            ' Format using scientific notation if magnitude is below threshold
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

    ''' <summary>
    ''' Updates current active compound selection title display and triggers chart rendering routines.
    ''' </summary>
    Private Sub DisplayCurrentSource()
        ' Update chart title string
        If mainTitle IsNot Nothing Then
            mainTitle.Text = $"Yearly max concentrations of {currentSourceKey}"
        End If

        ' Render primary chart data
        PlotMainChart()

        ' Update summary file side panel display
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

    ''' <summary>
    ''' Controls visibility state of the box plot series, area, and title.
    ''' </summary>
    ''' <param name="visible">True to make box plot controls visible; False to hide.</param>
    Private Sub SetBoxPlotVisible(visible As Boolean)
        ' Enable or disable box plot series rendering
        If chartMain.Series.FindByName("BoxPlotSeries") IsNot Nothing Then
            chartMain.Series("BoxPlotSeries").Enabled = visible
        End If
        If chartMain.Series.FindByName("BoxPlotPoints") IsNot Nothing Then
            chartMain.Series("BoxPlotPoints").Enabled = visible
        End If

        ' Enable or disable box plot chart area
        If chartMain.ChartAreas.FindByName("BoxPlotArea") IsNot Nothing Then
            chartMain.ChartAreas("BoxPlotArea").Visible = visible
        End If

        ' Toggle box plot title visibility
        Dim title = chartMain.Titles.FindByName("BoxPlotTitle")
        If title IsNot Nothing Then
            title.Visible = visible
        End If
    End Sub

    ''' <summary>
    ''' Calculates annual peak concentrations, binds values to main chart series, and constructs box plot statistics.
    ''' </summary>
    Private Sub PlotMainChart()
        Dim records = allRecords(currentSourceKey)

        Dim seriesWater = chartMain.Series("Surface Water Max")
        Dim seriesBenthic = chartMain.Series("Benthic Max")

        ' Clear main series points
        seriesWater.Points.Clear()
        seriesBenthic.Points.Clear()

        ' Clear box plot series data points
        Dim seriesBoxPlot = chartMain.Series("BoxPlotSeries")
        Dim seriesBoxPoints = chartMain.Series("BoxPlotPoints")
        seriesBoxPlot.Points.Clear()
        seriesBoxPoints.Points.Clear()

        ' Exit if no records are available
        If records.Count = 0 Then Exit Sub

        ' Group records by calendar year ordered chronologically
        Dim yearlyGroups = records.GroupBy(Function(r) r.DateValue.Year).OrderBy(Function(g) g.Key).ToList()
        If yearlyGroups.Count = 0 Then Exit Sub

        ' Determine start year baseline
        Dim startYear As Integer = yearlyGroups.First().Key

        ' Calculate yearly max values for surface water
        Dim yearlyMaxWater = yearlyGroups.Select(Function(g) New With {
            .RelativeYear = g.Key - startYear,
            .ActualDate = g.OrderByDescending(Function(r) r.WaterCol).First().DateValue,
            .Value = g.OrderByDescending(Function(r) r.WaterCol).First().WaterCol
        }).ToList()

        ' Calculate yearly max values for benthic compartment
        Dim yearlyMaxBenthic = yearlyGroups.Select(Function(g) New With {
            .RelativeYear = g.Key - startYear,
            .ActualDate = g.OrderByDescending(Function(r) r.Benthic).First().DateValue,
            .Value = g.OrderByDescending(Function(r) r.Benthic).First().Benthic
        }).ToList()

        ' Populate surface water maximum series points
        For Each item In yearlyMaxWater
            Dim ptIndex As Integer = seriesWater.Points.AddXY(item.RelativeYear, item.Value)
            seriesWater.Points(ptIndex).ToolTip = $"SW Max ({currentSourceKey}){vbCrLf}Relative Year: {item.RelativeYear} (Actual: {item.ActualDate:dd.MM.yyyy}){vbCrLf}Value: {item.Value:F4} ppb{vbCrLf}(Click for detail view)"
            seriesWater.Points(ptIndex).Tag = item.ActualDate
        Next

        ' Populate benthic maximum series points
        For Each item In yearlyMaxBenthic
            Dim ptIndex As Integer = seriesBenthic.Points.AddXY(item.RelativeYear, item.Value)
            seriesBenthic.Points(ptIndex).ToolTip = $"Benthic Max ({currentSourceKey}){vbCrLf}Relative Year: {item.RelativeYear} (Actual: {item.ActualDate:dd.MM.yyyy}){vbCrLf}Value: {item.Value:F4} ppb{vbCrLf}(Click for detail view)"
            seriesBenthic.Points(ptIndex).Tag = item.ActualDate
        Next

        ' Configure and render BoxPlot overlay structure when details pane is closed
        If Not chkShowDetailChart.Checked Then
            SetBoxPlotVisible(True)

            Dim values As List(Of Double)
            Dim currentTitle As String
            Dim pointColor As Drawing.Color
            Dim valuePrefix As String

            ' Select target data values depending on active active mode (Surface Water vs. Benthic)
            If isSwMode Then
                values = yearlyMaxWater.Select(Function(x) x.Value).OrderBy(Function(v) v).ToList()
                currentTitle = "Surface Water Yearly max (ppb)"
                valuePrefix = "SW Max"

                ' Query active surface water series color setting
                Dim targetSeries = chartMain.Series.FindByName("SW")
                If targetSeries IsNot Nothing Then
                    pointColor = targetSeries.Color
                Else
                    pointColor = Drawing.Color.RoyalBlue
                End If
            Else
                values = yearlyMaxBenthic.Select(Function(x) x.Value).OrderBy(Function(v) v).ToList()
                currentTitle = "Benthic Yearly Max (ppb)"
                valuePrefix = "Benthic Max"

                ' Query active benthic series color setting
                Dim benthicSeries = chartMain.Series.FindByName("Benthic")
                If benthicSeries IsNot Nothing Then
                    pointColor = benthicSeries.Color
                Else
                    pointColor = Drawing.Color.SandyBrown
                End If
            End If

            ' Update BoxPlot title text
            Dim boxTitle = chartMain.Titles.FindByName("BoxPlotTitle")
            If boxTitle IsNot Nothing Then
                boxTitle.Text = currentTitle
            End If

            seriesBoxPoints.Points.Clear()

            ' Build box plot statistical values if records exist
            If values.Count > 0 Then
                ' Add individual data points to box plot scatter overlay
                For Each val As Double In values
                    Dim ptIdx As Integer = seriesBoxPoints.Points.AddXY(1.0, val)
                    seriesBoxPoints.Points(ptIdx).ToolTip = $"{valuePrefix}: {val:F4} ppb"
                Next

                ' Link box plot series to underlying point data
                seriesBoxPlot("BoxPlotSeries") = "BoxPlotPoints"
                seriesBoxPlot.Points.Clear()

                ' Apply styling colors
                seriesBoxPoints.Color = pointColor
                seriesBoxPlot.BorderColor = Drawing.Color.Black
                seriesBoxPlot.Color = Drawing.Color.FromArgb(0, 255, 255, 255)

                ' Calculate statistical parameters: Minimum, Maximum, Median
                Dim minVal As Double = values.First()
                Dim maxVal As Double = values.Last()
                Dim medianVal As Double
                Dim count As Integer = values.Count

                ' Calculate median value
                If count Mod 2 = 0 Then
                    medianVal = (values(count \ 2 - 1) + values(count \ 2)) / 2.0
                Else
                    medianVal = values(count \ 2)
                End If

                Dim labelFont As New Drawing.Font("Segoe UI", 12.0!, Drawing.FontStyle.Bold)

                ' Apply Minimum point label
                Dim minPoint = seriesBoxPoints.Points.FirstOrDefault(Function(p) p.YValues(0) = minVal)
                If minPoint IsNot Nothing Then
                    minPoint.Label = $"              Min: {minVal:F2}"
                    minPoint.Font = labelFont
                    minPoint.LabelForeColor = Drawing.Color.Black
                    minPoint("LabelStyle") = "Right"
                End If

                ' Apply Maximum point label
                Dim maxPoint = seriesBoxPoints.Points.FirstOrDefault(Function(p) p.YValues(0) = maxVal)
                If maxPoint IsNot Nothing Then
                    maxPoint.Label = $"              Max: {maxVal:F2}"
                    maxPoint.Font = labelFont
                    maxPoint.LabelForeColor = Drawing.Color.Black
                    maxPoint("LabelStyle") = "Right"
                End If

                ' Apply Median point label
                Dim medianPoint = seriesBoxPoints.Points.OrderBy(Function(p) Math.Abs(p.YValues(0) - medianVal)).FirstOrDefault()
                If medianPoint IsNot Nothing AndAlso medianPoint IsNot minPoint AndAlso medianPoint IsNot maxPoint Then
                    medianPoint.Label = $"              Med: {medianVal:F2}"
                    medianPoint.Font = labelFont
                    medianPoint.LabelForeColor = Drawing.Color.Black
                    medianPoint("LabelStyle") = "Right"
                End If

                ' Define static X-axis boundaries for box plot area
                Dim boxArea = chartMain.ChartAreas(seriesBoxPoints.ChartArea)
                boxArea.AxisX.Minimum = 0.0
                boxArea.AxisX.Maximum = 2.5
            End If
        End If

        ' Recalculate chart axes scale boundaries
        chartMain.ChartAreas("MainArea").RecalculateAxesScale()
        chartMain.ChartAreas("BoxPlotArea").RecalculateAxesScale()

        ' Automatically update detail view for first year in range
        If yearlyGroups.Any() Then
            UpdateDetailChart(yearlyGroups.First().Key)
        End If
    End Sub

    ''' <summary>
    ''' Updates daily detailed concentration chart series for a targeted calendar year.
    ''' </summary>
    ''' <param name="selectedYear">Calendar year integer to query and plot.</param>
    Private Sub UpdateDetailChart(selectedYear As Integer)
        Dim records = allRecords(currentSourceKey)

        Dim seriesWater = chartDetail.Series("Surface Water (Daily)")
        Dim seriesBenthic = chartDetail.Series("Benthic (Daily)")

        ' Clear existing daily chart series points
        seriesWater.Points.Clear()
        seriesBenthic.Points.Clear()

        ' Filter record dataset for selected calendar year
        Dim yearData = records.Where(Function(r) r.DateValue.Year = selectedYear) _
                              .OrderBy(Function(r) r.DateValue) _
                              .ToList()

        ' Populate daily surface water and benthic points
        For Each record In yearData
            Dim pWaterIdx As Integer = seriesWater.Points.AddXY(record.DateValue, record.WaterCol)
            seriesWater.Points(pWaterIdx).ToolTip = $"Water Col ({currentSourceKey}){vbCrLf}Date: {record.DateValue:dd.MM.yyyy}{vbCrLf}Value: {record.WaterCol:F4} ppb"

            Dim pBenthicIdx As Integer = seriesBenthic.Points.AddXY(record.DateValue, record.Benthic)
            seriesBenthic.Points(pBenthicIdx).ToolTip = $"Benthic ({currentSourceKey}){vbCrLf}Date: {record.DateValue:dd.MM.yyyy}{vbCrLf}Value: {record.Benthic:F4} ppb"
        Next

        ' Reset zoom view and update x-axis title
        Dim area = chartDetail.ChartAreas("DetailArea")
        area.AxisX.ScaleView.ZoomReset(0)
        area.AxisX.Title = $"Daily Values for Year {selectedYear} ({currentSourceKey})"
        area.AxisX.IntervalType = DateTimeIntervalType.Months
        area.AxisX.Interval = 1
        area.RecalculateAxesScale()
    End Sub

    ''' <summary>
    ''' Formats and appends a structured custom key-value item entry to the result summary legend.
    ''' </summary>
    ''' <param name="label">Metric description text label.</param>
    ''' <param name="value">Metric formatted value string.</param>
    Private Sub AddSummaryLegendItem(label As String, value As String)
        Dim item As New LegendItem()
        Dim combinedText As String

        ' Format header items without value offset
        If String.IsNullOrEmpty(value) OrElse value = " ----- " Then
            combinedText = label
        Else
            ' Pad label to fixed 7-character width and attach value
            Const labelWidth As Integer = 7
            combinedText = $"{label.PadRight(labelWidth)}    {value}"
        End If

        ' Create formatted legend cell using monospaced font
        Dim cellCombined As New LegendCell(LegendCellType.Text, combinedText) With {
            .Alignment = Drawing.ContentAlignment.MiddleLeft,
            .Font = New Drawing.Font("Courier New", 12.0!, Drawing.FontStyle.Bold)
        }

        item.Cells.Add(cellCombined)
        summaryLegend.CustomItems.Add(item)
    End Sub

#End Region

#Region " Event Handlers & User Interaction "

    ' --- Menu Event Handlers ---

    ''' <summary>
    ''' Handles user file selection dialog to open and parse a PWC output file.
    ''' </summary>
    Private Sub MenuOpen_Click(sender As Object, e As EventArgs)
        Using ofd As New OpenFileDialog()
            ofd.Filter = "PWC Output Files (*.out;*.csv)|*.out;*.csv|All Files (*.*)|*.*"
            ofd.Title = "Select PWC Output File"

            If ofd.ShowDialog() = DialogResult.OK Then
                ProcessSelectedFile(ofd.FileName)
            End If
        End Using
    End Sub

    ''' <summary>
    ''' Clears loaded application data and resets user interface views to initial state.
    ''' </summary>
    Private Sub MenuReset_Click(sender As Object, e As EventArgs)
        ResetAppData()
    End Sub

    ''' <summary>
    ''' Exports the current main chart display area to a GIF image file.
    ''' </summary>
    Private Sub MenuSaveAsGif_Click(sender As Object, e As EventArgs)
        Dim activeFile As String = filePaths(currentSourceKey)
        ' Check if data file is loaded
        If String.IsNullOrEmpty(activeFile) Then
            MessageBox.Show("Please open a valid data file first.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Prompt user for output location
        Using sfd As New SaveFileDialog()
            sfd.Filter = "GIF Image|*.gif"
            sfd.Title = "Save Charts as GIF"
            sfd.FileName = Path.GetFileNameWithoutExtension(activeFile) & ".gif"

            If sfd.ShowDialog() = DialogResult.OK Then
                Try
                    Dim totalWidth As Integer = chartMain.Width
                    Dim totalHeight As Integer = chartMain.Height

                    ' Render main chart into bitmap and write GIF file
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

    ''' <summary>
    ''' Prompts user to choose a directory and batch-processes all matching files within it.
    ''' </summary>
    Private Sub MenuProcessFolder_Click(sender As Object, e As EventArgs)
        Using fbd As New FolderBrowserDialog()
            fbd.Description = "Select Folder Containing PWC Files (*.out / *.csv)"
            fbd.ShowNewFolderButton = False
            fbd.SelectedPath = "C:\"

            If fbd.ShowDialog() = DialogResult.OK Then
                ' Execute batch process across target directory
                ProcessDirectoryBatch(fbd.SelectedPath)

                ' Ensure box plot title remains visible
                If chartMain.Titles.FindByName("BoxPlotTitle") IsNot Nothing Then
                    chartMain.Titles("BoxPlotTitle").Visible = True
                End If

                MessageBox.Show("Batch processing completed!", "Process Folder", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End Using
    End Sub

    ''' <summary>
    ''' Displays help dialog window providing instructions on generating output files in PWC 3.
    ''' </summary>
    Private Sub MenuHelp_Click(sender As Object, e As EventArgs)
        Using helpForm As New Form()
            helpForm.Text = "PWC 3.X Output Generation"
            helpForm.Size = New Drawing.Size(750, 900)
            helpForm.StartPosition = FormStartPosition.CenterParent
            helpForm.FormBorderStyle = FormBorderStyle.FixedDialog
            helpForm.MaximizeBox = False
            helpForm.MinimizeBox = False

            ' Setup close button control
            Dim btnClose As New Button() With {
                .Text = "OK",
                .DialogResult = DialogResult.OK,
                .Dock = DockStyle.Bottom,
                .Height = 40,
                .Font = New Drawing.Font("Segoe UI", 11.0!, Drawing.FontStyle.Bold)
            }
            helpForm.Controls.Add(btnClose)

            ' Setup scrollable container panel
            Dim scrollPanel As New Panel() With {
                .Dock = DockStyle.Fill,
                .AutoScroll = True,
                .Padding = New Padding(15)
            }
            helpForm.Controls.Add(scrollPanel)

            ' Section 1 text description and screenshot
            Dim lblToggleOutput As New Label() With {
                .Text = "To generate the required *.out files in PWC 3, first enable 'Optional Output' by clicking" & vbCrLf & " 'More Tabs -> Toggle More Outputs'" & vbCrLf,
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .MaximumSize = New Drawing.Size(680, 0),
                .Padding = New Padding(0, 5, 0, 10),
                .Font = New Drawing.Font("Segoe UI", 11.0!, Drawing.FontStyle.Regular)
            }

            Dim picToggleOutput As New PictureBox() With {
                .Dock = DockStyle.Top,
                .Height = 300,
                .SizeMode = PictureBoxSizeMode.AutoSize,
                .Padding = New Padding(0, 0, 0, 15)
            }
            If File.Exists("ToggleOutput.png") Then
                picToggleOutput.Load("ToggleOutput.png")
            End If

            ' Section 2 text description and screenshot
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
            If File.Exists("OptionalOutput.png") Then
                picOptionalOutput.Load("OptionalOutput.png")
            End If

            ' Add controls in reverse dock order
            scrollPanel.Controls.Add(picOptionalOutput)
            scrollPanel.Controls.Add(lblOptionalOutput)
            scrollPanel.Controls.Add(picToggleOutput)
            scrollPanel.Controls.Add(lblToggleOutput)

            scrollPanel.BringToFront()
            helpForm.ShowDialog(Me)
        End Using
    End Sub

    ' --- Control State Event Handlers ---

    ''' <summary>
    ''' Toggles the visibility state of the lower detail chart panel and box plot display overlay.
    ''' </summary>
    Private Sub OnShowDetailChart_CheckedChanged(sender As Object, e As EventArgs)
        If chartSplitContainer IsNot Nothing Then
            chartSplitContainer.Panel2Collapsed = Not chkShowDetailChart.Checked
            SetBoxPlotVisible(Not chkShowDetailChart.Checked)
        End If
    End Sub

    ''' <summary>
    ''' Handles active compound data source selection changes via radio button toggle.
    ''' </summary>
    Private Sub OnSourceRadioButton_CheckedChanged(sender As Object, e As EventArgs)
        Dim rb As RadioButton = CType(sender, RadioButton)
        If rb.Checked Then
            If rb Is rbParent Then currentSourceKey = "Parent"
            If rb Is rbDaughter Then currentSourceKey = "Daughter"
            If rb Is rbGranddaughter Then currentSourceKey = "Granddaughter"

            ' Update views to display new source data
            DisplayCurrentSource()
        End If
    End Sub

    ' --- Drag & Drop Event Handlers ---

    ''' <summary>
    ''' Handles file drag-enter events to allow drag-and-drop file imports.
    ''' </summary>
    Private Sub Form_DragEnter(sender As Object, e As DragEventArgs) Handles MyBase.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    ''' <summary>
    ''' Handles file drop events to trigger immediate parsing of dropped file paths.
    ''' </summary>
    Private Sub Form_DragDrop(sender As Object, e As DragEventArgs) Handles MyBase.DragDrop
        Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())

        If files IsNot Nothing AndAlso files.Length > 0 Then
            ProcessSelectedFile(files(0))
        End If
    End Sub

    ' --- Chart Mouse Interaction Event Handlers ---

    ''' <summary>
    ''' Handles mouse hover interactions over the main chart area, adjusting cursor styles and tooltips dynamically.
    ''' </summary>
    Private Sub chartMain_MouseMove(sender As Object, e As MouseEventArgs) Handles chartMain.MouseMove
        Dim result As HitTestResult = chartMain.HitTest(e.X, e.Y)

        If result IsNot Nothing Then
            Dim isOverBoxPlot As Boolean = False

            ' Check if mouse cursor hovers above BoxPlot components
            If (result.ChartArea IsNot Nothing AndAlso result.ChartArea.Name = "BoxPlotArea") OrElse
               (result.Series IsNot Nothing AndAlso (result.Series.Name = "BoxPlotSeries" OrElse result.Series.Name = "BoxPlotPoints")) OrElse
               (result.Object IsNot Nothing AndAlso TypeOf result.Object Is Title AndAlso CType(result.Object, Title).Name = "BoxPlotTitle") Then

                isOverBoxPlot = True
            End If

            ' Adjust cursor shape and tooltip context
            If isOverBoxPlot Then
                chartMain.Cursor = Cursors.Hand

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

            ' Change cursor style when hovering over interactive data points
            If result.ChartElementType = ChartElementType.DataPoint Then
                chartMain.Cursor = Cursors.Hand
            Else
                chartMain.Cursor = Cursors.Default
            End If
        End If
    End Sub

    ''' <summary>
    ''' Adjusts mouse cursor icon when hovering over data points in the detail chart view.
    ''' </summary>
    Private Sub chartDetail_MouseMove(sender As Object, e As MouseEventArgs) Handles chartDetail.MouseMove
        Dim result As HitTestResult = chartDetail.HitTest(e.X, e.Y)
        If result.ChartElementType = ChartElementType.DataPoint Then
            chartDetail.Cursor = Cursors.Hand
        Else
            chartDetail.Cursor = Cursors.Default
        End If
    End Sub

    ''' <summary>
    ''' Handles mouse click events on the main chart area for box plot mode toggling and detail view opening.
    ''' </summary>
    Private Sub chartMain_MouseClick(sender As Object, e As MouseEventArgs) Handles chartMain.MouseClick
        Dim result As HitTestResult = chartMain.HitTest(e.X, e.Y)

        ' Process left mouse button clicks
        If e.Button = MouseButtons.Left Then
            ' Check if click target lies within box plot area or series
            If result IsNot Nothing AndAlso
              (result.ChartArea?.Name = "BoxPlotArea" OrElse
               (result.Series IsNot Nothing AndAlso (result.Series.Name = "BoxPlotSeries" OrElse result.Series.Name = "BoxPlotPoints"))) Then

                ' Toggle box plot active display mode (Surface Water <-> Benthic)
                isSwMode = Not isSwMode

                ' Refresh chart rendering
                PlotMainChart()
                Exit Sub
            End If
        End If

        ' Handle clicks on individual chart data points
        If result.ChartElementType = ChartElementType.DataPoint Then
            Dim point As DataPoint = result.Series.Points(result.PointIndex)
            If point.Tag IsNot Nothing AndAlso TypeOf point.Tag Is DateTime Then
                Dim realDate As DateTime = CType(point.Tag, DateTime)
                chkShowDetailChart.Checked = True
                UpdateDetailChart(realDate.Year)
            End If
        End If
    End Sub

    ''' <summary>
    ''' Handles mouse wheel scrolling events over the detail chart to execute dynamic zooming.
    ''' </summary>
    Private Sub chartDetail_MouseWheel(sender As Object, e As MouseEventArgs) Handles chartDetail.MouseWheel
        Dim area As ChartArea = chartDetail.ChartAreas("DetailArea")
        Try
            If e.Delta < 0 Then
                ' Reset zoom on scroll down
                area.AxisX.ScaleView.ZoomReset(0)
            ElseIf e.Delta > 0 Then
                ' Execute zoom centered around mouse position on scroll up
                Dim xMin As Double = area.AxisX.ScaleView.ViewMinimum
                Dim xMax As Double = area.AxisX.ScaleView.ViewMaximum
                Dim xMouse As Double = area.AxisX.PixelPositionToValue(e.X)

                Dim newSpan As Double = (xMax - xMin) / 2
                area.AxisX.ScaleView.Zoom(xMouse - newSpan / 2, xMouse + newSpan / 2)
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Resets zoom levels on the detail chart upon mouse double-click.
    ''' </summary>
    Private Sub chartDetail_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles chartDetail.MouseDoubleClick
        chartDetail.ChartAreas("DetailArea").AxisX.ScaleView.ZoomReset(0)
    End Sub

#End Region

End Class