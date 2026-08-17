Imports System.IO
Imports System.Windows.Forms

Public Module main

    <STAThread>
    Sub Main(args As String())
        ' Form-Styles aktivieren (wichtig für korrektes Rendering)
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)

        ' FALL 1: Keine Argumente übergeben -> Form ganz normal starten
        If args.Length = 0 Then
            Application.Run(New FrmPWCResultViewer())
            Return
        End If

        ' FALL 2: Argumente wurden übergeben
        Dim filePath As String = args(0)

        ' Validierung: Existiert die Datei?
        If Not File.Exists(filePath) Then
            Console.WriteLine($"Fehler: Die Datei '{filePath}' wurde nicht gefunden.")
            Return
        End If

        ' Parameter 2 parsen (UI anzeigen oder nicht - Standard ist False)
        Dim showUI As Boolean = False
        If args.Length > 1 Then
            Boolean.TryParse(args(1), showUI)
        End If

        ' Form instanziieren
        Dim viewerForm As New FrmPWCResultViewer()

        ' Form laden (initiiert Controls & Charts)
        viewerForm.Show()

        ' Datei verarbeiten
        viewerForm.ProcessSelectedFile(filePath)

        If showUI Then
            ' Interaktiver Modus: Form anzeigen und offen lassen
            viewerForm.Show()
            viewerForm.ProcessSelectedFile(filePath)
            Application.Run(viewerForm)
        Else
            ' Headless / CLI Modus: 100% Unsichtbar im Hintergrund
            Try
                ' 1. WinForms-Handle erzwingen, ohne die Form sichtbar zu machen (kein .Show()!)
                Dim dummyHandle As IntPtr = viewerForm.Handle

                ' 2. Daten laden & Plot verarbeiten
                viewerForm.ProcessSelectedFile(filePath)

                ' 3. Chart Layout & Rendering im Speicher erzwingen
                Dim chart = viewerForm.MainChartControl
                If chart IsNot Nothing Then
                    chart.Update()
                End If

                ' 4. GIF direkt im Hintergrund speichern
                Dim gifPath As String = Path.ChangeExtension(filePath, ".gif")
                SaveChartDirectlyToGif_Main(viewerForm, gifPath)

                Console.WriteLine($"GIF erfolgreich gespeichert unter: {gifPath}")
            Catch ex As Exception
                Console.WriteLine($"Fehler beim automatischen Speichern des GIFs: {ex.Message}")
            Finally
                ' Form ordnungsgemäß aus dem Speicher entfernen
                viewerForm.Close()
                viewerForm.Dispose()
            End Try
        End If
    End Sub

    Private Sub SaveChartDirectlyToGif_Main(form As FrmPWCResultViewer, outputPath As String)
        ' Direkt über die Eigenschaft auf das Chart-Control zugreifen
        Dim chart = form.MainChartControl

        If chart IsNot Nothing Then
            ' Sicherstellen, dass das Chart im Hintergrund komplett gerendert wird
            chart.Update()

            Using bmp As New System.Drawing.Bitmap(chart.Width, chart.Height)
                chart.DrawToBitmap(bmp, New System.Drawing.Rectangle(0, 0, chart.Width, chart.Height))
                bmp.Save(outputPath, System.Drawing.Imaging.ImageFormat.Gif)
            End Using
        Else
            Throw New InvalidOperationException("Das Haupt-Chart ist nicht initialisiert.")
        End If
    End Sub

End Module