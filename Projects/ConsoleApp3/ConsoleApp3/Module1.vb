Imports System.IO


Module Module1

    Sub Main()
        Dim s As New System.IO.StreamWriter("C:\Index.html", False)
        Dim html_aux As String

        html_aux = "<html>
        <body>
        <h1>Heading</h1>
        <p>Hihihi1</p>
        </body>
        <script>
          
        </script>
    <style>
       p {color:blue;}

    </style>
        </html>"

        If System.IO.File.Exists("C:\Index.html") Then
            s.WriteLine(html_aux)
            s.Close()

        Else
            MkDir("c:\Index.html")
            s.WriteLine(html_aux)
            s.Close()
        End If
    End Sub

End Module
