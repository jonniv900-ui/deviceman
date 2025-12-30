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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.statusbar1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.GerenciadorDeRecursosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GerenciadorDeProcessosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RestaurarDriversToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.btnExportarHtml = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripDropDownButton1 = New System.Windows.Forms.ToolStripDropDownButton()
        Me.RelatórioCompletoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SomenteTelaSelecionadaToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.TVdispositivos = New System.Windows.Forms.TreeView()
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.LVdetalhes = New System.Windows.Forms.ListView()
        Me.RestaurarDriversToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.StatusStrip1.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.statusbar1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 428)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(740, 22)
        Me.StatusStrip1.TabIndex = 0
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'statusbar1
        '
        Me.statusbar1.Name = "statusbar1"
        Me.statusbar1.Size = New System.Drawing.Size(0, 17)
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem1})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(740, 24)
        Me.MenuStrip1.TabIndex = 1
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.GerenciadorDeRecursosToolStripMenuItem, Me.GerenciadorDeProcessosToolStripMenuItem, Me.RestaurarDriversToolStripMenuItem, Me.RestaurarDriversToolStripMenuItem1})
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(84, 20)
        Me.ToolStripMenuItem1.Text = "Ferramentas"
        '
        'GerenciadorDeRecursosToolStripMenuItem
        '
        Me.GerenciadorDeRecursosToolStripMenuItem.Name = "GerenciadorDeRecursosToolStripMenuItem"
        Me.GerenciadorDeRecursosToolStripMenuItem.Size = New System.Drawing.Size(209, 22)
        Me.GerenciadorDeRecursosToolStripMenuItem.Text = "Monitor de Recursos"
        '
        'GerenciadorDeProcessosToolStripMenuItem
        '
        Me.GerenciadorDeProcessosToolStripMenuItem.Name = "GerenciadorDeProcessosToolStripMenuItem"
        Me.GerenciadorDeProcessosToolStripMenuItem.Size = New System.Drawing.Size(209, 22)
        Me.GerenciadorDeProcessosToolStripMenuItem.Text = "Gerenciador de processos"
        '
        'RestaurarDriversToolStripMenuItem
        '
        Me.RestaurarDriversToolStripMenuItem.Name = "RestaurarDriversToolStripMenuItem"
        Me.RestaurarDriversToolStripMenuItem.Size = New System.Drawing.Size(209, 22)
        Me.RestaurarDriversToolStripMenuItem.Text = "Backup de drivers"
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnExportarHtml, Me.ToolStripDropDownButton1})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 24)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(740, 25)
        Me.ToolStrip1.TabIndex = 2
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'btnExportarHtml
        '
        Me.btnExportarHtml.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnExportarHtml.Image = CType(resources.GetObject("btnExportarHtml.Image"), System.Drawing.Image)
        Me.btnExportarHtml.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnExportarHtml.Name = "btnExportarHtml"
        Me.btnExportarHtml.Size = New System.Drawing.Size(23, 22)
        Me.btnExportarHtml.Text = "ToolStripButton1"
        '
        'ToolStripDropDownButton1
        '
        Me.ToolStripDropDownButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripDropDownButton1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.RelatórioCompletoToolStripMenuItem, Me.SomenteTelaSelecionadaToolStripMenuItem})
        Me.ToolStripDropDownButton1.Image = CType(resources.GetObject("ToolStripDropDownButton1.Image"), System.Drawing.Image)
        Me.ToolStripDropDownButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripDropDownButton1.Name = "ToolStripDropDownButton1"
        Me.ToolStripDropDownButton1.Size = New System.Drawing.Size(29, 22)
        Me.ToolStripDropDownButton1.Text = "Exportar relatório"
        '
        'RelatórioCompletoToolStripMenuItem
        '
        Me.RelatórioCompletoToolStripMenuItem.Name = "RelatórioCompletoToolStripMenuItem"
        Me.RelatórioCompletoToolStripMenuItem.Size = New System.Drawing.Size(208, 22)
        Me.RelatórioCompletoToolStripMenuItem.Text = "Relatório completo"
        '
        'SomenteTelaSelecionadaToolStripMenuItem
        '
        Me.SomenteTelaSelecionadaToolStripMenuItem.Name = "SomenteTelaSelecionadaToolStripMenuItem"
        Me.SomenteTelaSelecionadaToolStripMenuItem.Size = New System.Drawing.Size(208, 22)
        Me.SomenteTelaSelecionadaToolStripMenuItem.Text = "Somente tela selecionada"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 49)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TVdispositivos)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.LVdetalhes)
        Me.SplitContainer1.Size = New System.Drawing.Size(740, 379)
        Me.SplitContainer1.SplitterDistance = 189
        Me.SplitContainer1.TabIndex = 3
        '
        'TVdispositivos
        '
        Me.TVdispositivos.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TVdispositivos.ImageIndex = 0
        Me.TVdispositivos.ImageList = Me.ImageList1
        Me.TVdispositivos.Location = New System.Drawing.Point(0, 0)
        Me.TVdispositivos.Name = "TVdispositivos"
        Me.TVdispositivos.SelectedImageIndex = 0
        Me.TVdispositivos.Size = New System.Drawing.Size(189, 379)
        Me.TVdispositivos.TabIndex = 0
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "win")
        Me.ImageList1.Images.SetKeyName(1, "folder")
        Me.ImageList1.Images.SetKeyName(2, "chip")
        Me.ImageList1.Images.SetKeyName(3, "cpu")
        Me.ImageList1.Images.SetKeyName(4, "device")
        Me.ImageList1.Images.SetKeyName(5, "driver")
        Me.ImageList1.Images.SetKeyName(6, "ram")
        Me.ImageList1.Images.SetKeyName(7, "info")
        Me.ImageList1.Images.SetKeyName(8, "report")
        Me.ImageList1.Images.SetKeyName(9, "mb")
        Me.ImageList1.Images.SetKeyName(10, "pnp")
        Me.ImageList1.Images.SetKeyName(11, "gpu")
        Me.ImageList1.Images.SetKeyName(12, "network")
        Me.ImageList1.Images.SetKeyName(13, "nic")
        '
        'LVdetalhes
        '
        Me.LVdetalhes.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LVdetalhes.HideSelection = False
        Me.LVdetalhes.Location = New System.Drawing.Point(0, 0)
        Me.LVdetalhes.Name = "LVdetalhes"
        Me.LVdetalhes.Size = New System.Drawing.Size(547, 379)
        Me.LVdetalhes.TabIndex = 0
        Me.LVdetalhes.UseCompatibleStateImageBehavior = False
        '
        'RestaurarDriversToolStripMenuItem1
        '
        Me.RestaurarDriversToolStripMenuItem1.Name = "RestaurarDriversToolStripMenuItem1"
        Me.RestaurarDriversToolStripMenuItem1.Size = New System.Drawing.Size(209, 22)
        Me.RestaurarDriversToolStripMenuItem1.Text = "Restaurar drivers"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(740, 450)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents StatusStrip1 As StatusStrip
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents btnExportarHtml As ToolStripButton
    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents TVdispositivos As TreeView
    Friend WithEvents LVdetalhes As ListView
    Friend WithEvents ImageList1 As ImageList
    Friend WithEvents statusbar1 As ToolStripStatusLabel
    Friend WithEvents ToolStripDropDownButton1 As ToolStripDropDownButton
    Friend WithEvents RelatórioCompletoToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SomenteTelaSelecionadaToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents GerenciadorDeRecursosToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents GerenciadorDeProcessosToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents RestaurarDriversToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents RestaurarDriversToolStripMenuItem1 As ToolStripMenuItem
End Class
