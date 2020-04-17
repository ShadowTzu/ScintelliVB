<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormTest
    Inherits System.Windows.Forms.Form

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
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

    'Requise par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.  
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.TextBox_4 = New System.Windows.Forms.TextBox()
        Me.TextBox_3 = New System.Windows.Forms.TextBox()
        Me.TextBox_2 = New System.Windows.Forms.TextBox()
        Me.TextBox_1 = New System.Windows.Forms.TextBox()
        Me.TextBox_0 = New System.Windows.Forms.TextBox()
        Me.TimerText = New System.Windows.Forms.Timer(Me.components)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TextBox_4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TextBox_3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TextBox_2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TextBox_1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TextBox_0)
        Me.SplitContainer1.Size = New System.Drawing.Size(800, 450)
        Me.SplitContainer1.SplitterDistance = 151
        Me.SplitContainer1.TabIndex = 0
        '
        'TextBox_4
        '
        Me.TextBox_4.Dock = System.Windows.Forms.DockStyle.Top
        Me.TextBox_4.Location = New System.Drawing.Point(0, 80)
        Me.TextBox_4.Name = "TextBox_4"
        Me.TextBox_4.Size = New System.Drawing.Size(151, 20)
        Me.TextBox_4.TabIndex = 4
        '
        'TextBox_3
        '
        Me.TextBox_3.Dock = System.Windows.Forms.DockStyle.Top
        Me.TextBox_3.Location = New System.Drawing.Point(0, 60)
        Me.TextBox_3.Name = "TextBox_3"
        Me.TextBox_3.Size = New System.Drawing.Size(151, 20)
        Me.TextBox_3.TabIndex = 3
        '
        'TextBox_2
        '
        Me.TextBox_2.Dock = System.Windows.Forms.DockStyle.Top
        Me.TextBox_2.Location = New System.Drawing.Point(0, 40)
        Me.TextBox_2.Name = "TextBox_2"
        Me.TextBox_2.Size = New System.Drawing.Size(151, 20)
        Me.TextBox_2.TabIndex = 2
        '
        'TextBox_1
        '
        Me.TextBox_1.Dock = System.Windows.Forms.DockStyle.Top
        Me.TextBox_1.Location = New System.Drawing.Point(0, 20)
        Me.TextBox_1.Name = "TextBox_1"
        Me.TextBox_1.Size = New System.Drawing.Size(151, 20)
        Me.TextBox_1.TabIndex = 1
        '
        'TextBox_0
        '
        Me.TextBox_0.Dock = System.Windows.Forms.DockStyle.Top
        Me.TextBox_0.Location = New System.Drawing.Point(0, 0)
        Me.TextBox_0.Name = "TextBox_0"
        Me.TextBox_0.Size = New System.Drawing.Size(151, 20)
        Me.TextBox_0.TabIndex = 0
        '
        'TimerText
        '
        Me.TimerText.Enabled = True
        Me.TimerText.Interval = 150
        '
        'FormTest
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Name = "FormTest"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FormTest"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents TextBox_0 As TextBox
    Friend WithEvents TextBox_1 As TextBox
    Friend WithEvents TextBox_4 As TextBox
    Friend WithEvents TextBox_3 As TextBox
    Friend WithEvents TextBox_2 As TextBox
    Friend WithEvents TimerText As Timer
End Class
