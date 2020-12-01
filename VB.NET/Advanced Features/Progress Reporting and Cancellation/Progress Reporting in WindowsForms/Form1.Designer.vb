<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MainForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.ProgressBar = New System.Windows.Forms.ProgressBar()
        Me.LoadButton = New System.Windows.Forms.Button()
        Me.PercentageLabel = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'ProgressBar
        '
        Me.ProgressBar.Location = New System.Drawing.Point(86, 110)
        Me.ProgressBar.Name = "ProgressBar"
        Me.ProgressBar.Size = New System.Drawing.Size(291, 40)
        Me.ProgressBar.TabIndex = 0
        '
        'LoadButton
        '
        Me.LoadButton.Location = New System.Drawing.Point(86, 45)
        Me.LoadButton.Name = "LoadButton"
        Me.LoadButton.Size = New System.Drawing.Size(343, 39)
        Me.LoadButton.TabIndex = 1
        Me.LoadButton.Text = "Load"
        Me.LoadButton.UseVisualStyleBackColor = True
        '
        'Label
        '
        Me.PercentageLabel.AutoSize = True
        Me.PercentageLabel.Location = New System.Drawing.Point(383, 121)
        Me.PercentageLabel.Name = "Label"
        Me.PercentageLabel.Size = New System.Drawing.Size(0, 20)
        Me.PercentageLabel.TabIndex = 2
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(528, 192)
        Me.Controls.Add(Me.PercentageLabel)
        Me.Controls.Add(Me.LoadButton)
        Me.Controls.Add(Me.ProgressBar)
        Me.Name = "MainForm"
        Me.Text = "Progress reporting in Windows Forms"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ProgressBar As ProgressBar
    Friend WithEvents LoadButton As Button
    Friend WithEvents PercentageLabel As Label
End Class
