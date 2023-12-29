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
        Me.components = New System.ComponentModel.Container()
        Me.BGWorker_CheckPrice = New System.ComponentModel.BackgroundWorker()
        Me.Timer_CheckPrice = New System.Windows.Forms.Timer(Me.components)
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.DoCalculation = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'BGWorker_CheckPrice
        '
        '
        'Timer_CheckPrice
        '
        Me.Timer_CheckPrice.Enabled = True
        Me.Timer_CheckPrice.Interval = 6000
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(115, 162)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(567, 49)
        Me.ProgressBar1.TabIndex = 0
        '
        'DoCalculation
        '
        Me.DoCalculation.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DoCalculation.Location = New System.Drawing.Point(115, 249)
        Me.DoCalculation.Name = "DoCalculation"
        Me.DoCalculation.Size = New System.Drawing.Size(116, 43)
        Me.DoCalculation.TabIndex = 1
        Me.DoCalculation.Text = "Do Calculation"
        Me.DoCalculation.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.DoCalculation)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents BGWorker_CheckPrice As System.ComponentModel.BackgroundWorker
    Friend WithEvents Timer_CheckPrice As Timer
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents DoCalculation As Button
End Class
