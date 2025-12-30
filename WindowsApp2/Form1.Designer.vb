<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Descartar substituições de formulário para limpar a lista de componentes.
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

    'Exigido pelo Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'OBSERVAÇÃO: o procedimento a seguir é exigido pelo Windows Form Designer
    'Pode ser modificado usando o Windows Form Designer.  
    'Não o modifique usando o editor de códigos.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.ListView1 = New System.Windows.Forms.ListView()
        Me.BTNbackup = New System.Windows.Forms.Button()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.LBLstatus = New System.Windows.Forms.Label()
        Me.CHKSelecionarTodos = New System.Windows.Forms.CheckBox()
        Me.btncancelar = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'ListView1
        '
        Me.ListView1.HideSelection = False
        Me.ListView1.Location = New System.Drawing.Point(14, 25)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(731, 373)
        Me.ListView1.TabIndex = 0
        Me.ListView1.UseCompatibleStateImageBehavior = False
        '
        'BTNbackup
        '
        Me.BTNbackup.Location = New System.Drawing.Point(125, 405)
        Me.BTNbackup.Name = "BTNbackup"
        Me.BTNbackup.Size = New System.Drawing.Size(109, 27)
        Me.BTNbackup.TabIndex = 1
        Me.BTNbackup.Text = "Fazer Backup"
        Me.BTNbackup.UseVisualStyleBackColor = True
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(240, 409)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(236, 22)
        Me.ProgressBar1.TabIndex = 2
        '
        'LBLstatus
        '
        Me.LBLstatus.AutoSize = True
        Me.LBLstatus.Location = New System.Drawing.Point(491, 415)
        Me.LBLstatus.Name = "LBLstatus"
        Me.LBLstatus.Size = New System.Drawing.Size(10, 13)
        Me.LBLstatus.TabIndex = 3
        Me.LBLstatus.Text = "."
        '
        'CHKSelecionarTodos
        '
        Me.CHKSelecionarTodos.AutoSize = True
        Me.CHKSelecionarTodos.Location = New System.Drawing.Point(14, 411)
        Me.CHKSelecionarTodos.Name = "CHKSelecionarTodos"
        Me.CHKSelecionarTodos.Size = New System.Drawing.Size(105, 17)
        Me.CHKSelecionarTodos.TabIndex = 4
        Me.CHKSelecionarTodos.Text = "Selecionar todos"
        Me.CHKSelecionarTodos.UseVisualStyleBackColor = True
        '
        'btncancelar
        '
        Me.btncancelar.Location = New System.Drawing.Point(125, 404)
        Me.btncancelar.Name = "btncancelar"
        Me.btncancelar.Size = New System.Drawing.Size(109, 27)
        Me.btncancelar.TabIndex = 5
        Me.btncancelar.Text = "Cancelar"
        Me.btncancelar.UseVisualStyleBackColor = True
        Me.btncancelar.Visible = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(760, 450)
        Me.Controls.Add(Me.btncancelar)
        Me.Controls.Add(Me.CHKSelecionarTodos)
        Me.Controls.Add(Me.LBLstatus)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.BTNbackup)
        Me.Controls.Add(Me.ListView1)
        Me.Name = "Form1"
        Me.Text = "Backup de drivers"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ListView1 As ListView
    Friend WithEvents BTNbackup As Button
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents LBLstatus As Label
    Friend WithEvents CHKSelecionarTodos As CheckBox
    Friend WithEvents btncancelar As Button
End Class
