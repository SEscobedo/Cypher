Public Class Form2
    Dim fileName As String = False
    Dim flag As Boolean = False

    Private Sub RichTextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    End Sub

    Private Sub CargarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CargarToolStripMenuItem.Click

        OpenFileDialog1.Filter = "Archivos de texto *.txt|*.txt|Todos los archivos *.*|*.*"
        OpenFileDialog1.FileName = ""
        Dim r = OpenFileDialog1.ShowDialog()
        If r = DialogResult.OK Then
            fileName = OpenFileDialog1.FileName
            
            Try
                RichTextBox1.LoadFile(fileName, RichTextBoxStreamType.PlainText)
                flag = False
                Me.Tag = fileName
                RANDOM_KEY_FILE = fileName
            Catch ex As Exception
                MsgBox(Err.Description)
            End Try
            

        End If
    End Sub

    Private Sub Form2_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
       
        If flag = True Then
            Dim r = MsgBox("¿Desea guardar los cambios en la llave?", MsgBoxStyle.YesNoCancel)
            If r = MsgBoxResult.Yes Then
                If RANDOM_KEY_FILE = "" Then
                    GuardarComo()
                    flag = False
                Else
                    Try
                        RichTextBox1.SaveFile(RANDOM_KEY_FILE, RichTextBoxStreamType.PlainText)
                        flag = False
                        RANDOM_KEY = RichTextBox1.Text
                        MDIParent1.ToolStripTextBox1.Text = Mid$(RANDOM_KEY, 1, 20) & "..."
                    Catch ex As Exception
                        MsgBox(Err.Description)
                    End Try
                End If
            ElseIf r = MsgBoxResult.No Then
                'cerrar sin guardar
            Else
                'cancelar el cerrado de la ventana
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If RichTextBox1.Text = "" Then
            RichTextBox1.Text = RANDOM_KEY
        End If
    End Sub

    Private Sub RichTextBox1_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RichTextBox1.TextChanged
        flag = True
    End Sub

    Sub GuardarComo()
        Dim SaveFileDialog As New SaveFileDialog
        SaveFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        SaveFileDialog.Filter = "Archivos de texto *.txt|*.txt|Todos los archivos *.*|*.*"

        SaveFileDialog.FileName = "Llave " & Me.Text
        If SaveFileDialog.ShowDialog = DialogResult.OK Then
            Dim FileName As String = SaveFileDialog.FileName
            Me.RichTextBox1.Text = AdaptarLlave(Me.RichTextBox1.Text)
            Try
                Me.RichTextBox1.SaveFile(FileName, RichTextBoxStreamType.PlainText)
                Me.Tag = FileName
            Catch ex As Exception
                MsgBox(Err.Description)
            End Try

        End If

    End Sub

    Private Sub RevisarClaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RevisarClaveToolStripMenuItem.Click
        Dim t = AdaptarLlave(Me.RichTextBox1.Text)
        If t = "" Then
            MsgBox("La llave es válida")
        Else
            MsgBox("Los caracteres siguentes no son válidos: " & t)
        End If

    End Sub

    
End Class