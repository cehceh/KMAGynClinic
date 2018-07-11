Imports System.Threading

Module Program
    Public Sub Main()
        'MessageBox.Show(" Before program ")
        AddHandler Application.ThreadException, AddressOf OnThreadException

        ''Added this
        AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf UnhandledExceptionEventRaised
        Application.Run(New Form1()) ''//Use your main form here
    End Sub
    Private Sub OnThreadException(ByVal sender As Object,
                      ByVal e As ThreadExceptionEventArgs)
        ' This is where you handle the exception
        MessageBox.Show(e.Exception.Message)
    End Sub

    'Added this
    Sub UnhandledExceptionEventRaised(ByVal sender As Object, ByVal e As UnhandledExceptionEventArgs)
        If e.IsTerminating Then
            Dim o As Object = e.ExceptionObject
            MessageBox.Show(o.ToString) ' use EventLog instead
        End If
    End Sub
End Module


