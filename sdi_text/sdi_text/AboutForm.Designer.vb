<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class aboutForm
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
        Me.btnOk = New System.Windows.Forms.Button()
        Me.txtText = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(139, 198)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(75, 23)
        Me.btnOk.TabIndex = 1
        Me.btnOk.Text = "&OK"
        Me.btnOk.UseVisualStyleBackColor = True
        '
        'txtText
        '
        Me.txtText.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.txtText.Cursor = System.Windows.Forms.Cursors.No
        Me.txtText.Location = New System.Drawing.Point(0, 0)
        Me.txtText.Multiline = True
        Me.txtText.Name = "txtText"
        Me.txtText.ReadOnly = True
        Me.txtText.Size = New System.Drawing.Size(227, 192)
        Me.txtText.TabIndex = 2
        '
        'aboutForm
        '
        Me.AcceptButton = Me.btnOk
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(226, 233)
        Me.Controls.Add(Me.txtText)
        Me.Controls.Add(Me.btnOk)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "aboutForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "AboutForm"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnOk As Button
    Friend WithEvents txtText As TextBox
End Class
