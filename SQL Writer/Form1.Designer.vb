<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.btnOpenForm = New System.Windows.Forms.Button()
        Me.btnImport = New System.Windows.Forms.Button()
        Me.btnTableFields = New System.Windows.Forms.Button()
        Me.lblAuthor = New System.Windows.Forms.Label()
        Me.btnCloseDB = New System.Windows.Forms.Button()
        Me.btnHelp = New System.Windows.Forms.Button()
        Me.btnSaveReport = New System.Windows.Forms.Button()
        Me.lblVersion = New System.Windows.Forms.Label()
        Me.btnLoadScript = New System.Windows.Forms.Button()
        Me.btnSaveScript = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lstTables = New System.Windows.Forms.ListBox()
        Me.cbxSelect = New System.Windows.Forms.CheckBox()
        Me.btnLoad = New System.Windows.Forms.Button()
        Me.lblFilePath = New System.Windows.Forms.Label()
        Me.rtxtSelect = New System.Windows.Forms.TextBox()
        Me.btnRun = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnCreateDB = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.tTip = New System.Windows.Forms.ToolTip(Me.components)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnOpenForm
        '
        Me.btnOpenForm.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnOpenForm.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnOpenForm.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOpenForm.Location = New System.Drawing.Point(9, 26)
        Me.btnOpenForm.Name = "btnOpenForm"
        Me.btnOpenForm.Size = New System.Drawing.Size(381, 41)
        Me.btnOpenForm.TabIndex = 38
        Me.btnOpenForm.Text = "Open SQL Textbox In New Window"
        Me.btnOpenForm.UseVisualStyleBackColor = False
        '
        'btnImport
        '
        Me.btnImport.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.btnImport.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnImport.Font = New System.Drawing.Font("Microsoft Sans Serif", 21.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnImport.Location = New System.Drawing.Point(787, 412)
        Me.btnImport.Name = "btnImport"
        Me.btnImport.Size = New System.Drawing.Size(225, 46)
        Me.btnImport.TabIndex = 37
        Me.btnImport.Text = "Import Table(s)"
        Me.btnImport.UseVisualStyleBackColor = False
        '
        'btnTableFields
        '
        Me.btnTableFields.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.btnTableFields.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnTableFields.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnTableFields.Location = New System.Drawing.Point(1029, 393)
        Me.btnTableFields.Name = "btnTableFields"
        Me.btnTableFields.Size = New System.Drawing.Size(166, 30)
        Me.btnTableFields.TabIndex = 36
        Me.btnTableFields.Text = "Display Table Fields"
        Me.btnTableFields.UseVisualStyleBackColor = False
        '
        'lblAuthor
        '
        Me.lblAuthor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblAuthor.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAuthor.Location = New System.Drawing.Point(971, 476)
        Me.lblAuthor.Name = "lblAuthor"
        Me.lblAuthor.Size = New System.Drawing.Size(241, 27)
        Me.lblAuthor.TabIndex = 35
        Me.lblAuthor.Text = "<author>"
        Me.lblAuthor.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnCloseDB
        '
        Me.btnCloseDB.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.btnCloseDB.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnCloseDB.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCloseDB.Location = New System.Drawing.Point(787, 341)
        Me.btnCloseDB.Name = "btnCloseDB"
        Me.btnCloseDB.Size = New System.Drawing.Size(225, 46)
        Me.btnCloseDB.TabIndex = 34
        Me.btnCloseDB.Text = "Close Database File"
        Me.btnCloseDB.UseVisualStyleBackColor = False
        '
        'btnHelp
        '
        Me.btnHelp.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.btnHelp.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnHelp.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnHelp.Location = New System.Drawing.Point(1134, 12)
        Me.btnHelp.Name = "btnHelp"
        Me.btnHelp.Size = New System.Drawing.Size(72, 41)
        Me.btnHelp.TabIndex = 32
        Me.btnHelp.Text = "Help"
        Me.btnHelp.UseVisualStyleBackColor = False
        '
        'btnSaveReport
        '
        Me.btnSaveReport.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.btnSaveReport.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSaveReport.Font = New System.Drawing.Font("Microsoft Sans Serif", 21.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSaveReport.Location = New System.Drawing.Point(787, 263)
        Me.btnSaveReport.Name = "btnSaveReport"
        Me.btnSaveReport.Size = New System.Drawing.Size(225, 46)
        Me.btnSaveReport.TabIndex = 20
        Me.btnSaveReport.Text = "Save Report"
        Me.btnSaveReport.UseVisualStyleBackColor = False
        '
        'lblVersion
        '
        Me.lblVersion.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVersion.Location = New System.Drawing.Point(983, 595)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.Size = New System.Drawing.Size(194, 29)
        Me.lblVersion.TabIndex = 33
        Me.lblVersion.Text = "<version>"
        Me.lblVersion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnLoadScript
        '
        Me.btnLoadScript.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.btnLoadScript.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnLoadScript.Font = New System.Drawing.Font("Microsoft Sans Serif", 21.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnLoadScript.Location = New System.Drawing.Point(791, 26)
        Me.btnLoadScript.Name = "btnLoadScript"
        Me.btnLoadScript.Size = New System.Drawing.Size(223, 41)
        Me.btnLoadScript.TabIndex = 22
        Me.btnLoadScript.Text = "Load Script"
        Me.btnLoadScript.UseVisualStyleBackColor = False
        '
        'btnSaveScript
        '
        Me.btnSaveScript.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.btnSaveScript.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSaveScript.Font = New System.Drawing.Font("Microsoft Sans Serif", 21.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSaveScript.Location = New System.Drawing.Point(522, 26)
        Me.btnSaveScript.Name = "btnSaveScript"
        Me.btnSaveScript.Size = New System.Drawing.Size(223, 41)
        Me.btnSaveScript.TabIndex = 21
        Me.btnSaveScript.Text = "Save Script"
        Me.btnSaveScript.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(1017, 183)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(195, 29)
        Me.Label1.TabIndex = 31
        Me.Label1.Text = "Tables"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lstTables
        '
        Me.lstTables.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!)
        Me.lstTables.FormattingEnabled = True
        Me.lstTables.ItemHeight = 24
        Me.lstTables.Location = New System.Drawing.Point(1018, 215)
        Me.lstTables.Name = "lstTables"
        Me.lstTables.Size = New System.Drawing.Size(194, 172)
        Me.lstTables.TabIndex = 27
        '
        'cbxSelect
        '
        Me.cbxSelect.AutoSize = True
        Me.cbxSelect.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxSelect.Location = New System.Drawing.Point(1018, 71)
        Me.cbxSelect.Name = "cbxSelect"
        Me.cbxSelect.Size = New System.Drawing.Size(169, 24)
        Me.cbxSelect.TabIndex = 23
        Me.cbxSelect.Text = "SELECT Statement"
        Me.cbxSelect.UseVisualStyleBackColor = True
        '
        'btnLoad
        '
        Me.btnLoad.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.btnLoad.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnLoad.Font = New System.Drawing.Font("Microsoft Sans Serif", 21.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnLoad.Location = New System.Drawing.Point(348, 552)
        Me.btnLoad.Name = "btnLoad"
        Me.btnLoad.Size = New System.Drawing.Size(307, 72)
        Me.btnLoad.TabIndex = 29
        Me.btnLoad.Text = "Load Database"
        Me.btnLoad.UseVisualStyleBackColor = False
        '
        'lblFilePath
        '
        Me.lblFilePath.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFilePath.Location = New System.Drawing.Point(9, 510)
        Me.lblFilePath.Name = "lblFilePath"
        Me.lblFilePath.Size = New System.Drawing.Size(945, 29)
        Me.lblFilePath.TabIndex = 26
        Me.lblFilePath.Text = "<Database File>"
        Me.lblFilePath.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'rtxtSelect
        '
        Me.rtxtSelect.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rtxtSelect.Location = New System.Drawing.Point(13, 73)
        Me.rtxtSelect.Multiline = True
        Me.rtxtSelect.Name = "rtxtSelect"
        Me.rtxtSelect.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.rtxtSelect.Size = New System.Drawing.Size(999, 173)
        Me.rtxtSelect.TabIndex = 24
        '
        'btnRun
        '
        Me.btnRun.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.btnRun.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnRun.Font = New System.Drawing.Font("Microsoft Sans Serif", 21.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRun.Location = New System.Drawing.Point(1018, 101)
        Me.btnRun.Name = "btnRun"
        Me.btnRun.Size = New System.Drawing.Size(148, 72)
        Me.btnRun.TabIndex = 25
        Me.btnRun.Text = "Run (F5)"
        Me.btnRun.UseVisualStyleBackColor = False
        '
        'btnClose
        '
        Me.btnClose.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 21.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Location = New System.Drawing.Point(686, 552)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(268, 72)
        Me.btnClose.TabIndex = 30
        Me.btnClose.Text = "Exit"
        Me.btnClose.UseVisualStyleBackColor = False
        '
        'btnCreateDB
        '
        Me.btnCreateDB.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.btnCreateDB.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnCreateDB.Font = New System.Drawing.Font("Microsoft Sans Serif", 21.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCreateDB.Location = New System.Drawing.Point(9, 552)
        Me.btnCreateDB.Name = "btnCreateDB"
        Me.btnCreateDB.Size = New System.Drawing.Size(307, 72)
        Me.btnCreateDB.TabIndex = 28
        Me.btnCreateDB.Text = "Create Database"
        Me.btnCreateDB.UseVisualStyleBackColor = False
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridView1.DefaultCellStyle = DataGridViewCellStyle2
        Me.DataGridView1.Location = New System.Drawing.Point(9, 263)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.DataGridView1.Size = New System.Drawing.Size(772, 229)
        Me.DataGridView1.TabIndex = 19
        Me.DataGridView1.TabStop = False
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1221, 636)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnOpenForm)
        Me.Controls.Add(Me.btnImport)
        Me.Controls.Add(Me.btnTableFields)
        Me.Controls.Add(Me.lblAuthor)
        Me.Controls.Add(Me.btnCloseDB)
        Me.Controls.Add(Me.btnHelp)
        Me.Controls.Add(Me.btnSaveReport)
        Me.Controls.Add(Me.lblVersion)
        Me.Controls.Add(Me.btnLoadScript)
        Me.Controls.Add(Me.btnSaveScript)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lstTables)
        Me.Controls.Add(Me.cbxSelect)
        Me.Controls.Add(Me.btnLoad)
        Me.Controls.Add(Me.lblFilePath)
        Me.Controls.Add(Me.rtxtSelect)
        Me.Controls.Add(Me.btnRun)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnCreateDB)
        Me.Controls.Add(Me.DataGridView1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmMain"
        Me.Text = "SQL Writer"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnOpenForm As Button
    Friend WithEvents btnImport As Button
    Friend WithEvents btnTableFields As Button
    Friend WithEvents lblAuthor As Label
    Friend WithEvents btnCloseDB As Button
    Friend WithEvents btnHelp As Button
    Friend WithEvents btnSaveReport As Button
    Friend WithEvents lblVersion As Label
    Friend WithEvents btnLoadScript As Button
    Friend WithEvents btnSaveScript As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents lstTables As ListBox
    Friend WithEvents cbxSelect As CheckBox
    Friend WithEvents btnLoad As Button
    Friend WithEvents lblFilePath As Label
    Friend WithEvents rtxtSelect As TextBox
    Friend WithEvents btnRun As Button
    Friend WithEvents btnClose As Button
    Friend WithEvents btnCreateDB As Button
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents tTip As ToolTip
End Class
