<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SplashScreen1
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub
    Friend WithEvents ApplicationTitle As System.Windows.Forms.Label
    Friend WithEvents Version As System.Windows.Forms.Label
    Friend WithEvents Copyright As System.Windows.Forms.Label
    Friend WithEvents MainLayoutPanel As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents DetailsLayoutPanel As System.Windows.Forms.TableLayoutPanel

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
		Me.MainLayoutPanel = New System.Windows.Forms.TableLayoutPanel()
		Me.DetailsLayoutPanel = New System.Windows.Forms.TableLayoutPanel()
		Me.Version = New System.Windows.Forms.Label()
		Me.Copyright = New System.Windows.Forms.Label()
		Me.ApplicationTitle = New System.Windows.Forms.Label()
		Me.PictureBox1 = New System.Windows.Forms.PictureBox()
		Me.MainLayoutPanel.SuspendLayout()
		Me.DetailsLayoutPanel.SuspendLayout()
		CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.SuspendLayout()
		'
		'MainLayoutPanel
		'
		Me.MainLayoutPanel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
		Me.MainLayoutPanel.ColumnCount = 2
		Me.MainLayoutPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 243.0!))
		Me.MainLayoutPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 100.0!))
		Me.MainLayoutPanel.Controls.Add(Me.DetailsLayoutPanel, 1, 1)
		Me.MainLayoutPanel.Controls.Add(Me.ApplicationTitle, 1, 0)
		Me.MainLayoutPanel.Controls.Add(Me.PictureBox1, 0, 0)
		Me.MainLayoutPanel.Dock = System.Windows.Forms.DockStyle.Left
		Me.MainLayoutPanel.Location = New System.Drawing.Point(0, 0)
		Me.MainLayoutPanel.Name = "MainLayoutPanel"
		Me.MainLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 218.0!))
		Me.MainLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 38.0!))
		Me.MainLayoutPanel.Size = New System.Drawing.Size(496, 303)
		Me.MainLayoutPanel.TabIndex = 0
		'
		'DetailsLayoutPanel
		'
		Me.DetailsLayoutPanel.Anchor = System.Windows.Forms.AnchorStyles.None
		Me.DetailsLayoutPanel.BackColor = System.Drawing.Color.Transparent
		Me.DetailsLayoutPanel.ColumnCount = 1
		Me.DetailsLayoutPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 247.0!))
		Me.DetailsLayoutPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 142.0!))
		Me.DetailsLayoutPanel.Controls.Add(Me.Version, 0, 0)
		Me.DetailsLayoutPanel.Controls.Add(Me.Copyright, 0, 1)
		Me.DetailsLayoutPanel.Location = New System.Drawing.Point(246, 221)
		Me.DetailsLayoutPanel.Name = "DetailsLayoutPanel"
		Me.DetailsLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.0!))
		Me.DetailsLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.0!))
		Me.DetailsLayoutPanel.Size = New System.Drawing.Size(247, 79)
		Me.DetailsLayoutPanel.TabIndex = 1
		'
		'Version
		'
		Me.Version.Anchor = System.Windows.Forms.AnchorStyles.None
		Me.Version.BackColor = System.Drawing.Color.Transparent
		Me.Version.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Version.ForeColor = System.Drawing.SystemColors.ButtonFace
		Me.Version.Location = New System.Drawing.Point(3, 9)
		Me.Version.Name = "Version"
		Me.Version.Size = New System.Drawing.Size(241, 20)
		Me.Version.TabIndex = 1
		Me.Version.Text = "Versión {0}.{1:00}"
		'
		'Copyright
		'
		Me.Copyright.Anchor = System.Windows.Forms.AnchorStyles.None
		Me.Copyright.BackColor = System.Drawing.Color.Transparent
		Me.Copyright.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Copyright.ForeColor = System.Drawing.Color.MidnightBlue
		Me.Copyright.Location = New System.Drawing.Point(3, 39)
		Me.Copyright.Name = "Copyright"
		Me.Copyright.Size = New System.Drawing.Size(241, 40)
		Me.Copyright.TabIndex = 2
		Me.Copyright.Text = "Copyright"
		'
		'ApplicationTitle
		'
		Me.ApplicationTitle.Anchor = System.Windows.Forms.AnchorStyles.None
		Me.ApplicationTitle.BackColor = System.Drawing.Color.Transparent
		Me.ApplicationTitle.Font = New System.Drawing.Font("MS UI Gothic", 36.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ApplicationTitle.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
		Me.ApplicationTitle.Location = New System.Drawing.Point(246, 3)
		Me.ApplicationTitle.Name = "ApplicationTitle"
		Me.ApplicationTitle.Size = New System.Drawing.Size(247, 212)
		Me.ApplicationTitle.TabIndex = 0
		Me.ApplicationTitle.Text = "Título de la aplicación"
		Me.ApplicationTitle.TextAlign = System.Drawing.ContentAlignment.BottomLeft
		'
		'PictureBox1
		'
		Me.PictureBox1.Dock = System.Windows.Forms.DockStyle.Fill
		Me.PictureBox1.Image = Global.CypherX.My.Resources.Resources.CypherX1
		Me.PictureBox1.Location = New System.Drawing.Point(3, 3)
		Me.PictureBox1.Name = "PictureBox1"
		Me.PictureBox1.Size = New System.Drawing.Size(237, 212)
		Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
		Me.PictureBox1.TabIndex = 2
		Me.PictureBox1.TabStop = False
		'
		'SplashScreen1
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(496, 303)
		Me.ControlBox = False
		Me.Controls.Add(Me.MainLayoutPanel)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.Name = "SplashScreen1"
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.MainLayoutPanel.ResumeLayout(False)
		Me.DetailsLayoutPanel.ResumeLayout(False)
		CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)

	End Sub

	Friend WithEvents PictureBox1 As PictureBox
End Class
