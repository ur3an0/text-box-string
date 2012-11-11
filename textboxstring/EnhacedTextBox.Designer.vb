<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EnhacedTextBox
    Inherits System.Windows.Forms.UserControl

    'UserControl1 reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.label = New System.Windows.Forms.Label()
        Me.punto = New System.Windows.Forms.Label()
        Me.Informacion = New System.Windows.Forms.ToolTip(Me.components)
        Me.foco = New System.Windows.Forms.Button()
        Me.foco1 = New System.Windows.Forms.Button()
        Me.dato = New System.Windows.Forms.Label()
        Me.InfoDato = New System.Windows.Forms.ToolTip(Me.components)
        Me.txt = New System.Windows.Forms.TextBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'label
        '
        Me.label.AutoSize = True
        Me.label.BackColor = System.Drawing.Color.Transparent
        Me.label.Cursor = System.Windows.Forms.Cursors.Default
        Me.label.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.label.ForeColor = System.Drawing.Color.Olive
        Me.label.Location = New System.Drawing.Point(0, 5)
        Me.label.Name = "label"
        Me.label.Size = New System.Drawing.Size(14, 13)
        Me.label.TabIndex = 1
        Me.label.Text = "T"
        '
        'punto
        '
        Me.punto.AutoSize = True
        Me.punto.BackColor = System.Drawing.Color.Transparent
        Me.punto.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.punto.ForeColor = System.Drawing.Color.Olive
        Me.punto.Location = New System.Drawing.Point(17, 5)
        Me.punto.Name = "punto"
        Me.punto.Size = New System.Drawing.Size(10, 13)
        Me.punto.TabIndex = 1
        Me.punto.Text = ":"
        '
        'Informacion
        '
        Me.Informacion.IsBalloon = True
        Me.Informacion.ShowAlways = True
        Me.Informacion.ToolTipIcon = System.Windows.Forms.ToolTipIcon.Info
        '
        'foco
        '
        Me.foco.Location = New System.Drawing.Point(51, 1)
        Me.foco.Name = "foco"
        Me.foco.Size = New System.Drawing.Size(1, 1)
        Me.foco.TabIndex = 2
        Me.foco.Text = "Button1"
        Me.foco.UseVisualStyleBackColor = True
        '
        'foco1
        '
        Me.foco1.Location = New System.Drawing.Point(3, 1)
        Me.foco1.Name = "foco1"
        Me.foco1.Size = New System.Drawing.Size(1, 1)
        Me.foco1.TabIndex = 0
        Me.foco1.Text = "Button1"
        Me.foco1.UseVisualStyleBackColor = True
        '
        'dato
        '
        Me.dato.BackColor = System.Drawing.Color.Transparent
        Me.dato.Cursor = System.Windows.Forms.Cursors.Hand
        Me.dato.Font = New System.Drawing.Font("Tahoma", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dato.Location = New System.Drawing.Point(56, 5)
        Me.dato.Name = "dato"
        Me.dato.Size = New System.Drawing.Size(0, 13)
        Me.dato.TabIndex = 3
        '
        'InfoDato
        '
        Me.InfoDato.ShowAlways = True
        '
        'txt
        '
        Me.txt.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt.Location = New System.Drawing.Point(28, 1)
        Me.txt.Name = "txt"
        Me.txt.Size = New System.Drawing.Size(119, 21)
        Me.txt.TabIndex = 1
        Me.txt.Visible = False
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(28, 28)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(119, 21)
        Me.TextBox1.TabIndex = 1
        Me.TextBox1.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Olive
        Me.Label1.Location = New System.Drawing.Point(0, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(14, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "T"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Olive
        Me.Label2.Location = New System.Drawing.Point(17, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(10, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = ":"
        '
        'EnhacedTextBox
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.Controls.Add(Me.dato)
        Me.Controls.Add(Me.foco1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.punto)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.label)
        Me.Controls.Add(Me.foco)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.txt)
        Me.Name = "EnhacedTextBox"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Informacion As System.Windows.Forms.ToolTip
    Friend WithEvents foco As System.Windows.Forms.Button
    Private WithEvents label As System.Windows.Forms.Label
    Private WithEvents punto As System.Windows.Forms.Label
    Friend WithEvents foco1 As System.Windows.Forms.Button
    Friend WithEvents dato As System.Windows.Forms.Label
    Friend WithEvents InfoDato As System.Windows.Forms.ToolTip
    Private WithEvents txt As System.Windows.Forms.TextBox
    Private WithEvents TextBox1 As System.Windows.Forms.TextBox
    Private WithEvents Label1 As System.Windows.Forms.Label
    Private WithEvents Label2 As System.Windows.Forms.Label

End Class
