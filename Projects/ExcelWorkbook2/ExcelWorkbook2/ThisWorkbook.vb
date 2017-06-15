
Public Class ThisWorkbook

    Private Sub ThisWorkbook_Startup() Handles Me.Startup
        MsgBox("start")
    End Sub

    Private Sub ThisWorkbook_Shutdown() Handles Me.Shutdown
        MsgBox("end")
    End Sub


    Public Function pop(ByVal a As Integer, ByVal b As Integer)

        pop = a + b


    End Function


End Class

