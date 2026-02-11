<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(disposing As Boolean)
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
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.btnFixVault = New System.Windows.Forms.Button()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.btnTest1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.AllowDrop = True
        Me.Button1.Location = New System.Drawing.Point(18, 21)
        Me.Button1.Margin = New System.Windows.Forms.Padding(1)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(246, 38)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "DXF Fix"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(18, 61)
        Me.Button2.Margin = New System.Windows.Forms.Padding(1)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(246, 36)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Get ACAD Test"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(18, 101)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(246, 37)
        Me.Button3.TabIndex = 2
        Me.Button3.Text = "Get Big DXF's"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'btnFixVault
        '
        Me.btnFixVault.Location = New System.Drawing.Point(18, 144)
        Me.btnFixVault.Name = "btnFixVault"
        Me.btnFixVault.Size = New System.Drawing.Size(246, 37)
        Me.btnFixVault.TabIndex = 3
        Me.btnFixVault.Text = "Fix Vault"
        Me.btnFixVault.UseVisualStyleBackColor = True
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(18, 201)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(246, 23)
        Me.ProgressBar1.Style = System.Windows.Forms.ProgressBarStyle.Marquee
        Me.ProgressBar1.TabIndex = 4
        Me.ProgressBar1.Visible = False
        '
        'btnTest1
        '
        Me.btnTest1.Location = New System.Drawing.Point(270, 144)
        Me.btnTest1.Name = "btnTest1"
        Me.btnTest1.Size = New System.Drawing.Size(53, 37)
        Me.btnTest1.TabIndex = 5
        Me.btnTest1.Text = "TEST"
        Me.btnTest1.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(386, 251)
        Me.Controls.Add(Me.btnTest1)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.btnFixVault)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Margin = New System.Windows.Forms.Padding(1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents btnFixVault As Button
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents btnTest1 As Button
End Class
