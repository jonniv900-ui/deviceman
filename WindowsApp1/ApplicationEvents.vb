Namespace My
    Partial Friend Class MyApplication
        Private Sub MyApplication_Startup(sender As Object, e As ApplicationServices.StartupEventArgs) Handles Me.Startup

            ' Verifica o parâmetro -report
            If e.CommandLine.Contains("-report") Then
                ' Instancia o formulário em memória, mas NÃO o exibe
                Dim frm As New Global.Systeminspector.Form1()

                ' Executa a geração
                frm.GerarRelatorioAutomatico()



                ' CANCELA a inicialização normal do Windows Forms (o app fecha aqui)
                e.Cancel = True

                ' Verifica o parâmetro -monitor
            ElseIf e.CommandLine.Contains("-monitor") Then
                Me.MainForm = Global.Systeminspector.SysRes

                ' Comportamento padrão
            Else
                Me.MainForm = Global.Systeminspector.Form1
            End If
        End Sub
    End Class
End Namespace