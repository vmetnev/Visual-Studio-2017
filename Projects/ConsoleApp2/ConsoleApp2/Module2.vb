Module Module2



    Sub nmn()
        Dim D As Object = CreateObject("Excel.Application") 'Объявляем переменную D как объект Excel.

        Dim res As String

        D.Workbooks.open("C:\Users\vladislav.metnev\Desktop\1.xlsm") 'Открываем книгу по указанному адресу.
        D.Visible = True 'Это свойство таблицы, если True - то она будет видна, если False - то нет.В принципе это значение можно неписать, т.к. автоматически оно будет равным False
        D.Sheets(1).Activate() 'Активируем первый лист в книге (Если необходимо)

        res = D.Sheets(1).Range("B5").Value 'Считываем в  TextBox1.Text значение ячейки ("A1").
        MsgBox(res)
        Console.ReadKey()



        D.Sheets(1).Range("A2").Value = "lala"

        D.ActiveWindow.Close(SaveChanges:=True) 'Если значение рано False то при закрытии приложения изминения не сохранятся, если True изменения сохранятся.
        D.Quit()
        D = Nothing

    End Sub





End Module
