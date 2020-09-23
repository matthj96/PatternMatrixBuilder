<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmVariables
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
        Me.lstVariables = New System.Windows.Forms.ListBox()
        Me.btnOk = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnToggleAll = New System.Windows.Forms.Button()
        Me.btnToggleNone = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lstVariables
        '
        Me.lstVariables.FormattingEnabled = True
        Me.lstVariables.ItemHeight = 25
        Me.lstVariables.Location = New System.Drawing.Point(23, 18)
        Me.lstVariables.Name = "lstVariables"
        Me.lstVariables.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple
        Me.lstVariables.Size = New System.Drawing.Size(768, 779)
        Me.lstVariables.TabIndex = 0
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(562, 816)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(229, 75)
        Me.btnOk.TabIndex = 1
        Me.btnOk.Text = "Ok"
        Me.btnOk.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(327, 816)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(229, 75)
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnToggleAll
        '
        Me.btnToggleAll.Location = New System.Drawing.Point(23, 816)
        Me.btnToggleAll.Name = "btnToggleAll"
        Me.btnToggleAll.Size = New System.Drawing.Size(298, 37)
        Me.btnToggleAll.TabIndex = 3
        Me.btnToggleAll.Text = "Toggle All"
        Me.btnToggleAll.UseVisualStyleBackColor = True
        '
        'btnToggleNone
        '
        Me.btnToggleNone.Location = New System.Drawing.Point(23, 859)
        Me.btnToggleNone.Name = "btnToggleNone"
        Me.btnToggleNone.Size = New System.Drawing.Size(298, 32)
        Me.btnToggleNone.TabIndex = 4
        Me.btnToggleNone.Text = "Toggle None"
        Me.btnToggleNone.UseVisualStyleBackColor = True
        '
        'frmVariables
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(12.0!, 25.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(820, 913)
        Me.Controls.Add(Me.btnToggleNone)
        Me.Controls.Add(Me.btnToggleAll)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.lstVariables)
        Me.Name = "frmVariables"
        Me.Text = "Select Variables To Model"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lstVariables As ListBox
    Friend WithEvents btnOk As Button
    Friend WithEvents btnCancel As Button
    Friend WithEvents btnToggleAll As Button
    Friend WithEvents btnToggleNone As Button
End Class
