Imports System.Windows.Forms

Public Class Dialog1

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim d As New OpenFileDialog
        d.Filter = "Archivos de texto *.txt|*.txt|Todos los archivos *.*|*.*"
        d.FileName = ""
        d.Title = "Archivo de llave de inicio"
        Dim r = d.ShowDialog()
        If r = DialogResult.OK Then
            Me.TextBox1.Text = d.FileName
        End If

    End Sub

    Private Sub Dialog1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Me.DialogResult = DialogResult.OK Then
            My.Computer.Registry.CurrentUser.CreateSubKey("ArchivoInicio")
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ArchivoInicio", "ArchivoInicio", Me.TextBox1.Text)
            My.Computer.Registry.CurrentUser.CreateSubKey("Alfabeto")
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Alfabeto", "Alfabeto", Me.TextBox2.Text)
            config.Alfabeto = TextBox2.Text
            ALFABETO = config.Alfabeto
            config.ArchivoInicio = TextBox1.Text
            config.InicioLlave = Val(TextBox3.Text)
            My.Computer.Registry.CurrentUser.CreateSubKey("InicioLlave")
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\InicioLlave", "InicioLlave", Me.TextBox3.Text)
        Else
            TextBox1.Text = config.ArchivoInicio
            TextBox2.Text = config.Alfabeto
            TextBox3.Text = config.InicioLlave
        End If

    End Sub

    Private Sub Dialog1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox2.Text = config.Alfabeto
        TextBox1.Text = config.ArchivoInicio
        TextBox3.Text = config.InicioLlave
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        TextBox2.Text = ALFABETO_STD
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged
       
    End Sub
End Class
