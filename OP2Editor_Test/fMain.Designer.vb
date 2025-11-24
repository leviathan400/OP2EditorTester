<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class fMain
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
        Me.txtConsole = New System.Windows.Forms.TextBox()
        Me.btnLoadMap = New System.Windows.Forms.Button()
        Me.TimerStartup = New System.Windows.Forms.Timer(Me.components)
        Me.btnExtractCLM = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'txtConsole
        '
        Me.txtConsole.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtConsole.BackColor = System.Drawing.Color.White
        Me.txtConsole.Font = New System.Drawing.Font("Consolas", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtConsole.Location = New System.Drawing.Point(12, 63)
        Me.txtConsole.Multiline = True
        Me.txtConsole.Name = "txtConsole"
        Me.txtConsole.ReadOnly = True
        Me.txtConsole.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtConsole.Size = New System.Drawing.Size(714, 363)
        Me.txtConsole.TabIndex = 1
        Me.txtConsole.WordWrap = False
        '
        'btnLoadMap
        '
        Me.btnLoadMap.Enabled = False
        Me.btnLoadMap.Location = New System.Drawing.Point(12, 12)
        Me.btnLoadMap.Name = "btnLoadMap"
        Me.btnLoadMap.Size = New System.Drawing.Size(122, 23)
        Me.btnLoadMap.TabIndex = 2
        Me.btnLoadMap.Text = "Load Map"
        Me.btnLoadMap.UseVisualStyleBackColor = True
        '
        'TimerStartup
        '
        '
        'btnExtractCLM
        '
        Me.btnExtractCLM.Enabled = False
        Me.btnExtractCLM.Location = New System.Drawing.Point(169, 12)
        Me.btnExtractCLM.Name = "btnExtractCLM"
        Me.btnExtractCLM.Size = New System.Drawing.Size(122, 23)
        Me.btnExtractCLM.TabIndex = 3
        Me.btnExtractCLM.Text = "Extract CLM"
        Me.btnExtractCLM.UseVisualStyleBackColor = True
        '
        'fMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(738, 438)
        Me.Controls.Add(Me.btnExtractCLM)
        Me.Controls.Add(Me.btnLoadMap)
        Me.Controls.Add(Me.txtConsole)
        Me.MaximizeBox = False
        Me.Name = "fMain"
        Me.Text = "OP2Editor Tester "
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtConsole As TextBox
    Friend WithEvents btnLoadMap As Button
    Friend WithEvents TimerStartup As Timer
    Friend WithEvents btnExtractCLM As Button
End Class
