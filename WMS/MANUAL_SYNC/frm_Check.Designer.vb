<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_Check
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
        Me.btnImportFromExcel = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnImportFromExcel
        '
        Me.btnImportFromExcel.Location = New System.Drawing.Point(52, 55)
        Me.btnImportFromExcel.Name = "btnImportFromExcel"
        Me.btnImportFromExcel.Size = New System.Drawing.Size(182, 26)
        Me.btnImportFromExcel.TabIndex = 0
        Me.btnImportFromExcel.Text = "Sync"
        Me.btnImportFromExcel.UseVisualStyleBackColor = True
        '
        'frm_Check
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(313, 124)
        Me.Controls.Add(Me.btnImportFromExcel)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frm_Check"
        Me.Text = "Sync Intermediate DB to SAP"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnImportFromExcel As System.Windows.Forms.Button

End Class
