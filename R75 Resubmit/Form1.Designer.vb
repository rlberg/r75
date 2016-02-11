<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainForm
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
        Me.btnStart = New System.Windows.Forms.Button()
        Me.cmbEnv = New System.Windows.Forms.ComboBox()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(215, 40)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(109, 23)
        Me.btnStart.TabIndex = 0
        Me.btnStart.Text = "Start Reprocessing"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'cmbEnv
        '
        Me.cmbEnv.FormattingEnabled = True
        Me.cmbEnv.Items.AddRange(New Object() {"DEV01", "DEV02", "PROD01", "PROD03"})
        Me.cmbEnv.Location = New System.Drawing.Point(32, 40)
        Me.cmbEnv.Name = "cmbEnv"
        Me.cmbEnv.Size = New System.Drawing.Size(161, 21)
        Me.cmbEnv.TabIndex = 1
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(47, 92)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(0, 13)
        Me.lblStatus.TabIndex = 2
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(340, 134)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.cmbEnv)
        Me.Controls.Add(Me.btnStart)
        Me.Name = "MainForm"
        Me.Text = "R75 Resubmit"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents cmbEnv As System.Windows.Forms.ComboBox
    Friend WithEvents lblStatus As System.Windows.Forms.Label

End Class
