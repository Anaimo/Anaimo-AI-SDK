<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
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
    Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
    Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
    Me.picInputs = New System.Windows.Forms.PictureBox()
    Me.picOutputs = New System.Windows.Forms.PictureBox()
    Me.picTmp = New System.Windows.Forms.PictureBox()
    Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
    CType(Me.picInputs, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.picOutputs, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.picTmp, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SuspendLayout()
    '
    'OpenFileDialog1
    '
    Me.OpenFileDialog1.FileName = "OpenFileDialog1"
    '
    'picInputs
    '
    Me.picInputs.Location = New System.Drawing.Point(0, 0)
    Me.picInputs.Name = "picInputs"
    Me.picInputs.Size = New System.Drawing.Size(642, 290)
    Me.picInputs.TabIndex = 0
    Me.picInputs.TabStop = False
    '
    'picOutputs
    '
    Me.picOutputs.Location = New System.Drawing.Point(648, 0)
    Me.picOutputs.Name = "picOutputs"
    Me.picOutputs.Size = New System.Drawing.Size(153, 290)
    Me.picOutputs.TabIndex = 1
    Me.picOutputs.TabStop = False
    '
    'picTmp
    '
    Me.picTmp.Location = New System.Drawing.Point(134, 336)
    Me.picTmp.Name = "picTmp"
    Me.picTmp.Size = New System.Drawing.Size(119, 89)
    Me.picTmp.TabIndex = 2
    Me.picTmp.TabStop = False
    Me.picTmp.Visible = False
    '
    'Form1
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(800, 332)
    Me.Controls.Add(Me.picTmp)
    Me.Controls.Add(Me.picOutputs)
    Me.Controls.Add(Me.picInputs)
    Me.Name = "Form1"
    Me.Text = "Form1"
    CType(Me.picInputs, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.picOutputs, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.picTmp, System.ComponentModel.ISupportInitialize).EndInit()
    Me.ResumeLayout(False)

  End Sub

  Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
    Friend WithEvents picInputs As PictureBox
    Friend WithEvents picOutputs As PictureBox
    Friend WithEvents picTmp As PictureBox
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
End Class
