<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPgBar
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
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.lblProgressText = New System.Windows.Forms.Label()
        Me.startAsyncButton = New System.Windows.Forms.Button()
        Me.cancelAsyncButton = New System.Windows.Forms.Button()
        Me.resultLabel = New System.Windows.Forms.Label()
        Me.backgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.SuspendLayout()
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(31, 66)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(424, 23)
        Me.ProgressBar1.Style = System.Windows.Forms.ProgressBarStyle.Marquee
        Me.ProgressBar1.TabIndex = 0
        '
        'lblProgressText
        '
        Me.lblProgressText.AutoSize = True
        Me.lblProgressText.Location = New System.Drawing.Point(37, 29)
        Me.lblProgressText.Name = "lblProgressText"
        Me.lblProgressText.Size = New System.Drawing.Size(71, 13)
        Me.lblProgressText.TabIndex = 1
        Me.lblProgressText.Text = "Processing...."
        Me.lblProgressText.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'startAsyncButton
        '
        Me.startAsyncButton.Location = New System.Drawing.Point(63, 119)
        Me.startAsyncButton.Name = "startAsyncButton"
        Me.startAsyncButton.Size = New System.Drawing.Size(116, 30)
        Me.startAsyncButton.TabIndex = 2
        Me.startAsyncButton.Text = "Start Async"
        Me.startAsyncButton.UseVisualStyleBackColor = True
        '
        'cancelAsyncButton
        '
        Me.cancelAsyncButton.Location = New System.Drawing.Point(294, 119)
        Me.cancelAsyncButton.Name = "cancelAsyncButton"
        Me.cancelAsyncButton.Size = New System.Drawing.Size(116, 30)
        Me.cancelAsyncButton.TabIndex = 3
        Me.cancelAsyncButton.Text = "Cancel Async"
        Me.cancelAsyncButton.UseVisualStyleBackColor = True
        '
        'resultLabel
        '
        Me.resultLabel.AutoSize = True
        Me.resultLabel.Location = New System.Drawing.Point(205, 29)
        Me.resultLabel.Name = "resultLabel"
        Me.resultLabel.Size = New System.Drawing.Size(71, 13)
        Me.resultLabel.TabIndex = 4
        Me.resultLabel.Text = "Processing...."
        Me.resultLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'frmPgBar
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(498, 177)
        Me.Controls.Add(Me.resultLabel)
        Me.Controls.Add(Me.cancelAsyncButton)
        Me.Controls.Add(Me.startAsyncButton)
        Me.Controls.Add(Me.lblProgressText)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Name = "frmPgBar"
        Me.Text = "Form2"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents lblProgressText As Label
    Friend WithEvents startAsyncButton As Button
    Friend WithEvents cancelAsyncButton As Button
    Friend WithEvents resultLabel As Label
    Friend WithEvents backgroundWorker1 As System.ComponentModel.BackgroundWorker
End Class
