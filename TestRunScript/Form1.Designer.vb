<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.txtTransLog = New System.Windows.Forms.TextBox()
        Me.btnRun = New System.Windows.Forms.Button()
        Me.txtResult = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'txtTransLog
        '
        Me.txtTransLog.BackColor = System.Drawing.SystemColors.HighlightText
        Me.txtTransLog.Cursor = System.Windows.Forms.Cursors.Arrow
        Me.txtTransLog.Location = New System.Drawing.Point(28, 79)
        Me.txtTransLog.Multiline = True
        Me.txtTransLog.Name = "txtTransLog"
        Me.txtTransLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtTransLog.Size = New System.Drawing.Size(605, 368)
        Me.txtTransLog.TabIndex = 8
        '
        'btnRun
        '
        Me.btnRun.Location = New System.Drawing.Point(520, 29)
        Me.btnRun.Name = "btnRun"
        Me.btnRun.Size = New System.Drawing.Size(113, 44)
        Me.btnRun.TabIndex = 9
        Me.btnRun.Text = "Run"
        Me.btnRun.UseVisualStyleBackColor = True
        '
        'txtResult
        '
        Me.txtResult.BackColor = System.Drawing.SystemColors.HighlightText
        Me.txtResult.Cursor = System.Windows.Forms.Cursors.Arrow
        Me.txtResult.Location = New System.Drawing.Point(669, 79)
        Me.txtResult.Multiline = True
        Me.txtResult.Name = "txtResult"
        Me.txtResult.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtResult.Size = New System.Drawing.Size(605, 368)
        Me.txtResult.TabIndex = 10
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1513, 480)
        Me.Controls.Add(Me.txtResult)
        Me.Controls.Add(Me.btnRun)
        Me.Controls.Add(Me.txtTransLog)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtTransLog As TextBox
    Friend WithEvents btnRun As Button
    Friend WithEvents txtResult As TextBox
End Class
