Public Class Form1
    Property Modificado As Boolean = False
    Property ActiveKey As String = ""

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Me.RichTextBox1.Text = DesEncriptar(Me.RichTextBox1.Text, ToolStripTextBox1.Text, ToolStripComboBox1.Text)
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Dim text As String
        text = Encriptar(Me.RichTextBox1.Text, ToolStripTextBox1.Text, ToolStripComboBox1.Text)

        Me.RichTextBox1.Clear()
        Me.RichTextBox1.AppendText(text)
    End Sub

    Private Sub RichTextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RichTextBox1.KeyDown
        If e.KeyValue = Keys.F5 Then
            RichTextBox1.SelectedText = Clipboard.GetText()
        End If
    End Sub
    
  
    Private Sub RichTextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RichTextBox1.TextChanged
        Me.Modificado = True
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Me.modificado = True Then
            Dim r = MsgBox("¿Desea guardar los cambios en el documento activo?", MsgBoxStyle.YesNoCancel)
            If r = MsgBoxResult.Yes Then
                If Me.Tag = "" Then
                    GuardarComo()
                    Me.Modificado = False
                Else
                    MDIParent1.Guardar()
                End If
            ElseIf r = MsgBoxResult.No Then
                'cerrar sin guardar
            Else
                'cancelar el cerrado de la ventana
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If ToolStripComboBox1.Text = "Perfect security" Or ToolStripComboBox1.Text = "Perfect security II" Then
            ToolStripTextBox1.Enabled = False
            ToolStripLabel1.Enabled = False
        Else
            ToolStripTextBox1.Enabled = True
            ToolStripLabel1.Enabled = True
        End If
    End Sub
    Sub GuardarComo()
        Dim SaveFileDialog As New SaveFileDialog
        SaveFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        SaveFileDialog.Filter = "Archivos de texto (*.txt)|*.txt|Todos los archivos (*.*)|*.*"
       
        SaveFileDialog.FileName = "Encriptado " & Me.Text
        If SaveFileDialog.ShowDialog = DialogResult.OK Then
            Dim FileName As String = SaveFileDialog.FileName
            Try
                Me.RichTextBox1.SaveFile(FileName, RichTextBoxStreamType.PlainText)
                Me.Tag = FileName
            Catch ex As Exception
                MsgBox(Err.Description)
            End Try
            
        End If

        ToolStripButton2.ToolTipText = ""
        ToolStripButton1.ToolTipText = ""

    End Sub

   
    Private Sub ToolStripComboBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.TextChanged
        If ToolStripComboBox1.Text = "Perfect security" Or ToolStripComboBox1.Text = "Perfect security II" Then
            ToolStripTextBox1.Enabled = False
            ToolStripLabel1.Enabled = False
        Else
            ToolStripTextBox1.Enabled = True
            ToolStripLabel1.Enabled = True
        End If
    End Sub
End Class