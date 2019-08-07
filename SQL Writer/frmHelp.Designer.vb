<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmHelp
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmHelp))
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnSaveHelp = New System.Windows.Forms.Button()
        Me.txtHelp = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lstTopics = New System.Windows.Forms.ListBox()
        Me.SuspendLayout()
        '
        'btnClose
        '
        Me.btnClose.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 21.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Location = New System.Drawing.Point(22, 499)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(204, 41)
        Me.btnClose.TabIndex = 16
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = False
        '
        'btnSaveHelp
        '
        Me.btnSaveHelp.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.btnSaveHelp.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSaveHelp.Font = New System.Drawing.Font("Microsoft Sans Serif", 21.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSaveHelp.Location = New System.Drawing.Point(22, 391)
        Me.btnSaveHelp.Name = "btnSaveHelp"
        Me.btnSaveHelp.Size = New System.Drawing.Size(204, 41)
        Me.btnSaveHelp.TabIndex = 15
        Me.btnSaveHelp.Text = "Save to File"
        Me.btnSaveHelp.UseVisualStyleBackColor = False
        '
        'txtHelp
        '
        Me.txtHelp.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtHelp.Location = New System.Drawing.Point(304, 29)
        Me.txtHelp.Multiline = True
        Me.txtHelp.Name = "txtHelp"
        Me.txtHelp.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtHelp.Size = New System.Drawing.Size(834, 573)
        Me.txtHelp.TabIndex = 18
        Me.txtHelp.TabStop = False
        Me.txtHelp.Text = "========================================================================="
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(13, 91)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(213, 29)
        Me.Label1.TabIndex = 17
        Me.Label1.Text = "Double Click on a Topic"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lstTopics
        '
        Me.lstTopics.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!)
        Me.lstTopics.FormattingEnabled = True
        Me.lstTopics.ItemHeight = 24
        Me.lstTopics.Location = New System.Drawing.Point(14, 123)
        Me.lstTopics.Name = "lstTopics"
        Me.lstTopics.Size = New System.Drawing.Size(284, 172)
        Me.lstTopics.TabIndex = 14
        '
        'frmHelp
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1151, 631)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnSaveHelp)
        Me.Controls.Add(Me.txtHelp)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lstTopics)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmHelp"
        Me.Text = "Help"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnClose As Button
    Friend WithEvents btnSaveHelp As Button
    Friend WithEvents txtHelp As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents lstTopics As ListBox
End Class
