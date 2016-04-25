<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmImageModifier
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnOpen = New System.Windows.Forms.Button()
        Me.btnConvert = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.cboFilter = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.picConverted = New System.Windows.Forms.PictureBox()
        Me.picOriginal = New System.Windows.Forms.PictureBox()
        Me.txtQuit = New System.Windows.Forms.Button()
        CType(Me.picConverted, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picOriginal, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnOpen
        '
        Me.btnOpen.BackColor = System.Drawing.Color.Maroon
        Me.btnOpen.FlatAppearance.BorderSize = 0
        Me.btnOpen.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnOpen.ForeColor = System.Drawing.Color.White
        Me.btnOpen.Location = New System.Drawing.Point(32, 38)
        Me.btnOpen.Name = "btnOpen"
        Me.btnOpen.Size = New System.Drawing.Size(124, 32)
        Me.btnOpen.TabIndex = 0
        Me.btnOpen.Text = " Open Image"
        Me.btnOpen.UseVisualStyleBackColor = False
        '
        'btnConvert
        '
        Me.btnConvert.BackColor = System.Drawing.Color.DarkSeaGreen
        Me.btnConvert.FlatAppearance.BorderSize = 0
        Me.btnConvert.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnConvert.ForeColor = System.Drawing.Color.White
        Me.btnConvert.Location = New System.Drawing.Point(440, 38)
        Me.btnConvert.Name = "btnConvert"
        Me.btnConvert.Size = New System.Drawing.Size(75, 32)
        Me.btnConvert.TabIndex = 1
        Me.btnConvert.Text = "Convert"
        Me.btnConvert.UseVisualStyleBackColor = False
        '
        'btnSave
        '
        Me.btnSave.BackColor = System.Drawing.Color.Green
        Me.btnSave.FlatAppearance.BorderSize = 0
        Me.btnSave.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSave.ForeColor = System.Drawing.Color.White
        Me.btnSave.Location = New System.Drawing.Point(550, 38)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 32)
        Me.btnSave.TabIndex = 2
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = False
        '
        'cboFilter
        '
        Me.cboFilter.BackColor = System.Drawing.Color.Indigo
        Me.cboFilter.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cboFilter.ForeColor = System.Drawing.Color.White
        Me.cboFilter.FormattingEnabled = True
        Me.cboFilter.Items.AddRange(New Object() {"   Monochrome", "   Gray Scale", "   Conversion Factor", "   Sepia"})
        Me.cboFilter.Location = New System.Drawing.Point(187, 37)
        Me.cboFilter.Name = "cboFilter"
        Me.cboFilter.Size = New System.Drawing.Size(172, 24)
        Me.cboFilter.TabIndex = 3
        Me.cboFilter.Text = "   Please select a filter"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.Color.LightSteelBlue
        Me.Label1.Location = New System.Drawing.Point(437, 99)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(57, 17)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Preview"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.ForeColor = System.Drawing.Color.LightSteelBlue
        Me.Label2.Location = New System.Drawing.Point(29, 99)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(57, 17)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Original"
        '
        'picConverted
        '
        Me.picConverted.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picConverted.Location = New System.Drawing.Point(440, 128)
        Me.picConverted.Name = "picConverted"
        Me.picConverted.Size = New System.Drawing.Size(387, 291)
        Me.picConverted.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picConverted.TabIndex = 6
        Me.picConverted.TabStop = False
        '
        'picOriginal
        '
        Me.picOriginal.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picOriginal.Location = New System.Drawing.Point(32, 128)
        Me.picOriginal.Name = "picOriginal"
        Me.picOriginal.Size = New System.Drawing.Size(387, 291)
        Me.picOriginal.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picOriginal.TabIndex = 7
        Me.picOriginal.TabStop = False
        '
        'txtQuit
        '
        Me.txtQuit.BackColor = System.Drawing.Color.CadetBlue
        Me.txtQuit.FlatAppearance.BorderSize = 0
        Me.txtQuit.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.txtQuit.ForeColor = System.Drawing.Color.White
        Me.txtQuit.Location = New System.Drawing.Point(657, 38)
        Me.txtQuit.Name = "txtQuit"
        Me.txtQuit.Size = New System.Drawing.Size(75, 32)
        Me.txtQuit.TabIndex = 8
        Me.txtQuit.Text = "Quit"
        Me.txtQuit.UseVisualStyleBackColor = False
        '
        'frmImageModifier
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Black
        Me.ClientSize = New System.Drawing.Size(859, 447)
        Me.Controls.Add(Me.txtQuit)
        Me.Controls.Add(Me.picOriginal)
        Me.Controls.Add(Me.picConverted)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cboFilter)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnConvert)
        Me.Controls.Add(Me.btnOpen)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmImageModifier"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Image Modifier"
        CType(Me.picConverted, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picOriginal, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnOpen As Button
    Friend WithEvents btnConvert As Button
    Friend WithEvents btnSave As Button
    Friend WithEvents cboFilter As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents picConverted As PictureBox
    Friend WithEvents picOriginal As PictureBox
    Friend WithEvents txtQuit As Button
End Class
