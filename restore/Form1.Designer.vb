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
        Me.btncancelar = New System.Windows.Forms.Button()
        Me.CHKSelecionarTodos = New System.Windows.Forms.CheckBox()
        Me.LBLstatus = New System.Windows.Forms.Label()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.BTNrestaurarDrivers = New System.Windows.Forms.Button()
        Me.ListView1 = New System.Windows.Forms.ListView()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.BTNabrirbackup = New System.Windows.Forms.Button()
        Me.CHKdryrun = New System.Windows.Forms.CheckBox()
        Me.chkdebug = New System.Windows.Forms.CheckBox()
        Me.txtdebug = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'btncancelar
        '
        Me.btncancelar.Location = New System.Drawing.Point(123, 396)
        Me.btncancelar.Name = "btncancelar"
        Me.btncancelar.Size = New System.Drawing.Size(109, 27)
        Me.btncancelar.TabIndex = 11
        Me.btncancelar.Text = "Cancelar"
        Me.btncancelar.UseVisualStyleBackColor = True
        Me.btncancelar.Visible = False
        '
        'CHKSelecionarTodos
        '
        Me.CHKSelecionarTodos.AutoSize = True
        Me.CHKSelecionarTodos.Location = New System.Drawing.Point(12, 398)
        Me.CHKSelecionarTodos.Name = "CHKSelecionarTodos"
        Me.CHKSelecionarTodos.Size = New System.Drawing.Size(105, 17)
        Me.CHKSelecionarTodos.TabIndex = 10
        Me.CHKSelecionarTodos.Text = "Selecionar todos"
        Me.CHKSelecionarTodos.UseVisualStyleBackColor = True
        '
        'LBLstatus
        '
        Me.LBLstatus.AutoSize = True
        Me.LBLstatus.Location = New System.Drawing.Point(489, 402)
        Me.LBLstatus.Name = "LBLstatus"
        Me.LBLstatus.Size = New System.Drawing.Size(10, 13)
        Me.LBLstatus.TabIndex = 9
        Me.LBLstatus.Text = "."
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(238, 396)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(236, 22)
        Me.ProgressBar1.TabIndex = 8
        '
        'BTNrestaurarDrivers
        '
        Me.BTNrestaurarDrivers.Location = New System.Drawing.Point(123, 392)
        Me.BTNrestaurarDrivers.Name = "BTNrestaurarDrivers"
        Me.BTNrestaurarDrivers.Size = New System.Drawing.Size(109, 27)
        Me.BTNrestaurarDrivers.TabIndex = 7
        Me.BTNrestaurarDrivers.Text = "Restaurar"
        Me.BTNrestaurarDrivers.UseVisualStyleBackColor = True
        '
        'ListView1
        '
        Me.ListView1.HideSelection = False
        Me.ListView1.Location = New System.Drawing.Point(12, 12)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(731, 373)
        Me.ListView1.TabIndex = 6
        Me.ListView1.UseCompatibleStateImageBehavior = False
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'BTNabrirbackup
        '
        Me.BTNabrirbackup.Location = New System.Drawing.Point(606, 391)
        Me.BTNabrirbackup.Name = "BTNabrirbackup"
        Me.BTNabrirbackup.Size = New System.Drawing.Size(136, 44)
        Me.BTNabrirbackup.TabIndex = 12
        Me.BTNabrirbackup.Text = "Abrir um backup"
        Me.BTNabrirbackup.UseVisualStyleBackColor = True
        '
        'CHKdryrun
        '
        Me.CHKdryrun.AutoSize = True
        Me.CHKdryrun.Location = New System.Drawing.Point(12, 425)
        Me.CHKdryrun.Name = "CHKdryrun"
        Me.CHKdryrun.Size = New System.Drawing.Size(119, 17)
        Me.CHKdryrun.TabIndex = 13
        Me.CHKdryrun.Text = "Simular restauração"
        Me.CHKdryrun.UseVisualStyleBackColor = True
        '
        'chkdebug
        '
        Me.chkdebug.AutoSize = True
        Me.chkdebug.Location = New System.Drawing.Point(149, 430)
        Me.chkdebug.Name = "chkdebug"
        Me.chkdebug.Size = New System.Drawing.Size(64, 17)
        Me.chkdebug.TabIndex = 14
        Me.chkdebug.Text = "Console"
        Me.chkdebug.UseVisualStyleBackColor = True
        '
        'txtdebug
        '
        Me.txtdebug.BackColor = System.Drawing.Color.Black
        Me.txtdebug.Font = New System.Drawing.Font("Consolas", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtdebug.ForeColor = System.Drawing.Color.Lime
        Me.txtdebug.Location = New System.Drawing.Point(12, 461)
        Me.txtdebug.Multiline = True
        Me.txtdebug.Name = "txtdebug"
        Me.txtdebug.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtdebug.Size = New System.Drawing.Size(729, 62)
        Me.txtdebug.TabIndex = 15
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(754, 448)
        Me.Controls.Add(Me.txtdebug)
        Me.Controls.Add(Me.chkdebug)
        Me.Controls.Add(Me.CHKdryrun)
        Me.Controls.Add(Me.BTNabrirbackup)
        Me.Controls.Add(Me.btncancelar)
        Me.Controls.Add(Me.CHKSelecionarTodos)
        Me.Controls.Add(Me.LBLstatus)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.BTNrestaurarDrivers)
        Me.Controls.Add(Me.ListView1)
        Me.Name = "Form1"
        Me.Text = "Restaurar Drivers"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btncancelar As Button
    Friend WithEvents CHKSelecionarTodos As CheckBox
    Friend WithEvents LBLstatus As Label
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents BTNrestaurarDrivers As Button
    Friend WithEvents ListView1 As ListView
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents BTNabrirbackup As Button
    Friend WithEvents CHKdryrun As CheckBox
    Friend WithEvents chkdebug As CheckBox
    Friend WithEvents txtdebug As TextBox
End Class
