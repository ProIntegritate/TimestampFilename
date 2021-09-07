Module Module1

    Public bUseUTCTime As Boolean = True

    Sub Main()

        Dim sArg() As String = Environment.GetCommandLineArgs
        Dim sFile As String = ""
        Dim bShort As Boolean = False
        Dim sExtension As String = ""
        Dim bDateOnly As Boolean = False

        If UBound(sArg) = 0 Then
            Console.ForegroundColor = Console.ForegroundColor.Green
            Console.WriteLine("Timestampfilename: Adds an ISO8601 timestamp to a given file." & vbCrLf & "Written in 2021 by Glenn Larsson.")
            Console.ForegroundColor = Console.ForegroundColor.White
            Console.WriteLine("-f, File (Required). Specify file to stamp i.e. -f'c:\temp\t.txt' (Quotation is optional)".Replace("'", Chr(34)))
            Console.WriteLine("-a, Append (Default). Adds timstamp after name i.e. 't.txt.2021-09-07T20_15_40.343Z'".Replace("'", Chr(34)))
            Console.WriteLine("-p, Prepend (Optional). Adds timestamp before name i.e. '2021-09-07T20_15_40.343Z.t.txt'".Replace("'", Chr(34)))
            Console.WriteLine("-s, Short timeformat (Optional), removes miliseconds i.e. '2021-09-07T20_15_40'".Replace("'", Chr(34)))
            Console.WriteLine("-e, Extension (Optional). Specify extension to append i.e. -e.csv adds '.csv' to the end of the file.".Replace("'", Chr(34)))
            Console.WriteLine("-d, Date (Optional). Cut off everything from timestamp except date i.e. '2021-09-07'".Replace("'", Chr(34)))
            Console.WriteLine("-l, Local time (Optional). Use local time instead of UTC time.".Replace("'", Chr(34)))
            Console.ForegroundColor = Console.ForegroundColor.Cyan
            Console.WriteLine("Example: Timestampfilename.exe -fc:\temp\t.txt -s -p -e.csv".Replace("'", Chr(34)))
            Console.WriteLine("This would rename : c:\temp\t.txt to c:\temp\2021-09-07T20_15_40.t.txt.csv".Replace("'", Chr(34)))
            Console.ForegroundColor = Console.ForegroundColor.White
            End
        End If

        Dim sPosition As String = "AFTER"

        Dim param As String = ""
        Try
            For n = 0 To UBound(sArg)
                param = Microsoft.VisualBasic.Left(sArg(n), 2)
                Select Case LCase(param)
                    Case "-p" ' Prepend
                        sPosition = "BEFORE"
                    Case "-a" ' Append
                        sPosition = "AFTER"
                    Case "-f" ' Filename
                        sFile = sArg(n).Replace(param, "")
                    Case "-s" ' Short timeformat (no miliseconds)
                        bShort = True
                    Case "-e" ' Extension
                        sExtension = sArg(n).Replace(param, "")
                    Case "-d" ' Date only
                        bDateOnly = True
                    Case "-l" ' Use Local time instead of UTC
                        bUseUTCTime = False
                End Select
            Next
        Catch ex As Exception
        End Try

        If System.IO.File.Exists(sFile) = False Then
            Console.WriteLine("File '" & sFile & "' does not exist.")
            Exit Sub
        End If

        Dim sTimestamp As String = fGetUTCTimestamp()
        If bShort = True Then sTimestamp = Microsoft.VisualBasic.Left(sTimestamp, sTimestamp.Length - 5) ' Removes ".000Z"
        If bDateOnly = True Then sTimestamp = Microsoft.VisualBasic.Left(sTimestamp, 10)

        sFile = sFile.Replace("/", "\")
        Dim sFN() As String = sFile.Split("\")

        Dim sFileName As String = sFN(UBound(sFN)).ToString
        Dim sFolderName As String = sFile.Replace(sFileName, "")

        If sPosition = "AFTER" Then
            System.IO.File.Move(sFile, sFolderName & sFileName & "." & sTimestamp & sExtension)
        End If

        If sPosition = "BEFORE" Then
            System.IO.File.Move(sFile, sFolderName & sTimestamp & "." & sFileName & sExtension)
        End If

    End Sub

    Public Function fGetUTCTimestamp() As String

        Dim d As Date
        If bUseUTCTime = True Then
            d = Now.ToUniversalTime
        Else
            d = Now.ToLocalTime
        End If

        Dim sCurrentTimestanp As String = ""

        sCurrentTimestanp = sCurrentTimestanp & d.Year.ToString & "-"
        sCurrentTimestanp = sCurrentTimestanp & Microsoft.VisualBasic.Right("0" & d.Month.ToString, 2) & "-"
        sCurrentTimestanp = sCurrentTimestanp & Microsoft.VisualBasic.Right("0" & d.Day.ToString, 2)
        sCurrentTimestanp = sCurrentTimestanp & "T"
        sCurrentTimestanp = sCurrentTimestanp & Microsoft.VisualBasic.Right("0" & d.Hour.ToString, 2) & "_"
        sCurrentTimestanp = sCurrentTimestanp & Microsoft.VisualBasic.Right("0" & d.Minute.ToString, 2) & "_"
        sCurrentTimestanp = sCurrentTimestanp & Microsoft.VisualBasic.Right("0" & d.Second.ToString, 2) & "."
        sCurrentTimestanp = sCurrentTimestanp & Microsoft.VisualBasic.Right("000" & d.Millisecond.ToString, 3) & "Z"

        Return sCurrentTimestanp

    End Function

End Module
