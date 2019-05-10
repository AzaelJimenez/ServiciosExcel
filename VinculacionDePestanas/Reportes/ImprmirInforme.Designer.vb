<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ImprmirInforme
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
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

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.txtMagnitud = New System.Windows.Forms.TextBox()
        Me.txtInforme = New System.Windows.Forms.TextBox()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'txtMagnitud
        '
        Me.txtMagnitud.Location = New System.Drawing.Point(91, 14)
        Me.txtMagnitud.Name = "txtMagnitud"
        Me.txtMagnitud.Size = New System.Drawing.Size(100, 20)
        Me.txtMagnitud.TabIndex = 0
        '
        'txtInforme
        '
        Me.txtInforme.Location = New System.Drawing.Point(283, 14)
        Me.txtInforme.Name = "txtInforme"
        Me.txtInforme.Size = New System.Drawing.Size(100, 20)
        Me.txtInforme.TabIndex = 1
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Enabled = False
        Me.Label26.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label26.Location = New System.Drawing.Point(5, 14)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(80, 18)
        Me.Label26.TabIndex = 166
        Me.Label26.Text = "MAGNITUD"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Enabled = False
        Me.Label1.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(210, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(67, 18)
        Me.Label1.TabIndex = 167
        Me.Label1.Text = "INFORME"
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.SteelBlue
        Me.Button2.FlatAppearance.BorderSize = 0
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.White
        Me.Button2.Location = New System.Drawing.Point(244, 40)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(139, 33)
        Me.Button2.TabIndex = 170
        Me.Button2.Text = "VISTA PREVIA"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'ImprmirInforme
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(395, 85)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label26)
        Me.Controls.Add(Me.txtInforme)
        Me.Controls.Add(Me.txtMagnitud)
        Me.Name = "ImprmirInforme"
        Me.Text = "ImprmirInforme"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtMagnitud As TextBox
    Friend WithEvents txtInforme As TextBox
    Friend WithEvents Label26 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents Button2 As Button
End Class
