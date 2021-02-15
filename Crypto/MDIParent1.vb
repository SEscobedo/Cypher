Imports System.Windows.Forms

Public Class MDIParent1

    Private Sub ShowNewForm(ByVal sender As Object, ByVal e As EventArgs) Handles NewToolStripMenuItem.Click, NewToolStripButton.Click, NewWindowToolStripMenuItem.Click
        ' Cree una nueva instancia del formulario secundario.
        Dim ChildForm As New Form1
        ' Conviértalo en un elemento secundario de este formulario MDI antes de mostrarlo.
        ChildForm.MdiParent = Me

        m_ChildFormNumber += 1
        ChildForm.Text = "Ventana " & m_ChildFormNumber
        ChildForm.Show()

    End Sub

    Private Sub OpenFile(ByVal sender As Object, ByVal e As EventArgs) Handles OpenToolStripMenuItem.Click, OpenToolStripButton.Click
        Dim OpenFileDialog As New OpenFileDialog
        OpenFileDialog.Multiselect = True
        OpenFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        OpenFileDialog.Filter = "Archivos de texto *.txt|*.txt|Todos los archivos *.*|*.*"

        If (OpenFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
            Dim FileName() As String = OpenFileDialog.FileNames
            Dim i As Long

            For i = 0 To UBound(FileName)
                '++++++++++++++++++++++++++++++++++++++ En esta versión evitar que se abran archivos de MSOffice o PDF
                If FileName(i) Like "*.docx" Then
                    MsgBox("Los archivos MSWord no son compatibles con esta versión de Cypher", MsgBoxStyle.Exclamation)
                    Exit For
                ElseIf FileName(i) Like "*.pdf" Then
                    MsgBox("Los archivos PDF no son compatibles con esta versión de Cypher", MsgBoxStyle.Exclamation)
                    Exit For
                End If
                '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                ' Cree una nueva instancia del formulario secundario.
                Dim ChildForm As New Form1
                ' Conviértalo en un elemento secundario de este formulario MDI antes de mostrarlo.
                ChildForm.MdiParent = Me

                m_ChildFormNumber += 1
                ChildForm.Text = FileName(i)
                Try
                    ChildForm.RichTextBox1.LoadFile(FileName(i), RichTextBoxStreamType.PlainText)
                    ChildForm.Show()
                    ChildForm.Tag = FileName(i)
                Catch ex As Exception
                    MsgBox(Err.Description)
                End Try
            Next
        End If
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        GuardarComo()
    End Sub


    Private Sub ExitToolsStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub CutToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CutToolStripMenuItem.Click
        ' Utilice My.Computer.Clipboard para insertar el texto o las imágenes seleccionadas en el Portapapeles
        For Each ChildForm As Form1 In Me.MdiChildren
            If ChildForm.Equals(Me.ActiveMdiChild) Then
                My.Computer.Clipboard.SetText(ChildForm.RichTextBox1.SelectedText)
                ChildForm.RichTextBox1.SelectedText = ""
            End If
        Next
    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CopyToolStripMenuItem.Click
        ' Utilice My.Computer.Clipboard para insertar el texto o las imágenes seleccionadas en el Portapapeles
        For Each ChildForm As Form1 In Me.MdiChildren
            If ChildForm.Equals(Me.ActiveMdiChild) Then
                My.Computer.Clipboard.SetText(ChildForm.RichTextBox1.SelectedText)
            End If
        Next
    End Sub

    Private Sub PasteToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles PasteToolStripMenuItem.Click
        'Utilice My.Computer.Clipboard.GetText() o My.Computer.Clipboard.GetData para recuperar la información del Portapapeles.
        For Each ChildForm As Form1 In Me.MdiChildren
            If ChildForm.Equals(Me.ActiveMdiChild) Then
                ChildForm.RichTextBox1.SelectedText = My.Computer.Clipboard.GetText
            End If
        Next
    End Sub

    Private Sub ToolBarToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ToolBarToolStripMenuItem.Click
        Me.ToolStrip.Visible = Me.ToolBarToolStripMenuItem.Checked
    End Sub

    Private Sub StatusBarToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles StatusBarToolStripMenuItem.Click
        Me.StatusStrip.Visible = Me.StatusBarToolStripMenuItem.Checked
    End Sub

    Private Sub CascadeToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CascadeToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub TileVerticalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles TileVerticalToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.TileVertical)
    End Sub

    Private Sub TileHorizontalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles TileHorizontalToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub ArrangeIconsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ArrangeIconsToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.ArrangeIcons)
    End Sub

    Private Sub CloseAllToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CloseAllToolStripMenuItem.Click
        ' Cierre todos los formularios secundarios del principal.
        For Each ChildForm As Form In Me.MdiChildren
            ChildForm.Close()
        Next
    End Sub

    Private m_ChildFormNumber As Integer

    Private Sub LlavesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LlavesToolStripMenuItem.Click
        Dim filename As String
        OpenFileDialog1.Filter = "Archivos de texto *.txt|*.txt|Todos los archivos *.*|*.*"
        OpenFileDialog1.FileName = ""
        Dim r = OpenFileDialog1.ShowDialog()
        If r = DialogResult.OK Then
            filename = OpenFileDialog1.FileName
            Form2.RichTextBox1.LoadFile(filename, RichTextBoxStreamType.PlainText)
            ToolStripTextBox1.Text = Mid$(Form2.RichTextBox1.Text, 1, 20) & "..."
            RANDOM_KEY = Form2.RichTextBox1.Text
            RANDOM_KEY_FILE = filename
        End If
    End Sub

    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click, SaveToolStripButton.Click
        Guardar()
    End Sub
    Sub GuardarComo()
        Dim SaveFileDialog As New SaveFileDialog
        SaveFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        SaveFileDialog.Filter = "Archivos de texto *.txt|*.txt|Todos los archivos *.*|*.*"

        For Each ChildForm As Form1 In Me.MdiChildren
            If ChildForm.Equals(Me.ActiveMdiChild) Then
                SaveFileDialog.FileName = "Encriptado " & ChildForm.Text

                If (SaveFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
                    Dim FileName As String = SaveFileDialog.FileName
                    ChildForm.RichTextBox1.SaveFile(FileName, RichTextBoxStreamType.PlainText)
                    ChildForm.Tag = FileName
                    ChildForm.Modificado = False
                End If
            End If
        Next

    End Sub


   
    Private Sub MostrarLlaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MostrarLlaveToolStripMenuItem.Click
        Form2.MdiParent = Me
        Form2.Show()
    End Sub

   
    Private Sub OptionsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptionsToolStripMenuItem.Click
        Dialog1.ShowDialog()
    End Sub

    Private Sub MDIParent1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        My.Computer.Registry.CurrentUser.CreateSubKey("InicioLlave")
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\InicioLlave", "InicioLlave", config.InicioLlave)
    End Sub

    Private Sub MDIParent1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        AboutToolStripMenuItem.Text = "Acerca de " & Application.ProductName & " " & Application.ProductVersion

        Dim readValue = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ArchivoInicio", "ArchivoInicio", "")
        If readValue Is Nothing Then readValue = ""
        If Not readValue = "" Then
            Try
                Form2.RichTextBox1.LoadFile(readValue, RichTextBoxStreamType.PlainText)
                Me.ToolStripTextBox1.Text = Mid$(Form2.RichTextBox1.Text, 1, 20) & "..."
                config.ArchivoInicio = readValue
                RANDOM_KEY = Form2.RichTextBox1.Text
                RANDOM_KEY_FILE = readValue
            Catch ex As Exception
                MsgBox(Err.Description)
            End Try
        Else
            config.ArchivoInicio = ""
        End If

        config.Alfabeto = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Alfabeto", "Alfabeto", ALFABETO_STD)
        If config.Alfabeto Is Nothing Then config.Alfabeto = ALFABETO_STD

        ALFABETO = config.Alfabeto

        readValue = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\InicioLlave", "InicioLlave", "1")
        If readValue Is Nothing Then readValue = "1"
        config.InicioLlave = Val(readValue)

    End Sub
   

    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim tx As New Form1
        For Each tx In MdiChildren
            If tx Is Me.ActiveMdiChild Then
                e.Graphics.DrawString(tx.RichTextBox1.Text, tx.Font, Brushes.Black, 100, 100)
            End If
        Next
    End Sub

    Private Sub PrintPreviewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintPreviewToolStripMenuItem.Click, PrintPreviewToolStripButton.Click
        Dim tx As New Form1

        For Each tx In MdiChildren
            If tx Is Me.ActiveMdiChild Then
                PrintPreviewDialog1.Document = PrintDocument1
                PrintPreviewDialog1.ShowIcon = False
                PrintPreviewDialog1.ShowDialog()
            End If
        Next
    End Sub

    Private Sub PrintToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem.Click, PrintToolStripButton.Click
        Dim tx As New Form1

        For Each tx In MdiChildren
            If tx Is Me.ActiveMdiChild Then

                PrintDialog1.Document = PrintDocument1
                PrintDialog1.ShowDialog()
            End If
        Next
    End Sub

    Private Sub PrintSetupToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintSetupToolStripMenuItem.Click
        PrintDialog1.ShowDialog()
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click

        AboutBox1.ShowDialog()
    End Sub

    Private Sub SelectAllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectAllToolStripMenuItem.Click
        Dim tx As New Form1

        For Each tx In MdiChildren
            If tx Is Me.ActiveMdiChild Then
                tx.RichTextBox1.SelectAll()
            End If
        Next
    End Sub

    Public Sub Guardar()
        For Each ChildForm As Form1 In Me.MdiChildren
            If ChildForm.Equals(Me.ActiveMdiChild) Then
                If ChildForm.Modificado Then
                    If ChildForm.Tag = "" Then
                        GuardarComo()
                    Else
                        Try
                            ChildForm.RichTextBox1.SaveFile(ChildForm.Tag, RichTextBoxStreamType.PlainText)
                            ChildForm.Modificado = False
                        Catch ex As Exception
                            MsgBox(Err.Description)
                        End Try
                    End If

                End If
            End If
        Next
    End Sub

    Private Sub HelpToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HelpToolStripButton.Click

        System.Diagnostics.Process.Start("https://sescobedo.github.io/CypherX/")

    End Sub

    Private Sub IndexToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IndexToolStripMenuItem.Click

        System.Diagnostics.Process.Start("https://sescobedo.github.io/CypherX/")

    End Sub

    Private Sub TutorialesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TutorialesToolStripMenuItem.Click
        System.Diagnostics.Process.Start("https://www.youtube.com/channel/UC8d_dR5orxGDJRuDogmNugA/videos")
    End Sub
End Class
