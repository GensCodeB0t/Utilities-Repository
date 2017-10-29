Imports System
Imports System.IO
Imports System.Text
Imports System.Configuration

Module Module1
    Private appSettings = ConfigurationSettings.AppSettings
    Private totDistilled As Double
    Private totOrig As Double
    Private totSavings As Double
    Private Const FILETYPE As String = ".txt"
    Private inputFile As String = appSettings("inputDirPath") & appSettings("inputFilePath") & FILETYPE

    Sub Main()
        If File.Exists(inputFile) Then
            Try
                ' Create the File Stream
                Dim inputStream As StreamReader = New StreamReader(inputFile)
                Dim line As String

                Do
                    ' Read the line and determine if there is something there
                    line = inputStream.ReadLine()
                    If Not IsNothing(line) Then
                        ' Parse the line for related values
                        totDistilled += parseLine(line, "distilled =")
                        totOrig += parseLine(line, "orig-size =")
                        totSavings += parseLine(line, "savings =")
                    End If
                Loop Until line Is Nothing
                ' Close the File Stream
                inputStream.Close()
            Catch ex As Exception
                Console.WriteLine("A problem occured while trying to read:" & vbCrLf & inputFile & vbCrLf & ex.ToString())
            End Try

            ' Write the totals to the console
            Console.WriteLine()
            Console.WriteLine("***************************************************************")
            Console.WriteLine("Total Distilled Value:" & ChrW(9) & totDistilled & " Kb")
            Console.WriteLine("Total Original Value:" & ChrW(9) & totOrig & " Kb")
            Console.WriteLine("Total Savings Value:" & ChrW(9) & totSavings & " Kb")
            Console.WriteLine("***************************************************************")
            Console.WriteLine()
        Else
            Console.WriteLine("The File that you are trying to read does not exist!")
        End If
        ' Wait for the user to end the program
        Console.WriteLine("Enter S to save the results...")
        Dim answer As String = Console.ReadLine().ToUpper()
        Select Case answer
            Case "S" : writeTotalsToFile()
        End Select

    End Sub

    Private Function parseLine(line As String, searchStr As String)
        If line.Contains(searchStr) Then
            ' Find the searchStr in the line
            Dim subString As String = line.Substring(line.IndexOf(searchStr) + searchStr.Length, 7)

            ' Split the substring into characters that can be parsed into a Double and
            ' remove and empty elements
            Dim separators() As String = {" ", "  "}
            subString = subString.Split(separators, StringSplitOptions.RemoveEmptyEntries)(0)

            Return Double.Parse(subString)
        End If

        Return 0
    End Function

    Private Sub writeTotalsToFile()
        Dim outputFile As String = appSettings("outputDirPath") & appSettings("outputFilePath") & FILETYPE
        
        Dim outputStream As StreamWriter = New StreamWriter(outputFile, True)
        outputStream.WriteLine("#############################################################################")
        outputStream.WriteLine("PDF Compression data for: " & inputFile)
        outputStream.WriteLine("Date / Time: " & DateAndTime.Now.ToString("MM-dd-yy hh:mm:ss"))
        outputStream.WriteLine()
        outputStream.WriteLine("*****************************************************************************")
        outputStream.WriteLine("Total Distilled Value:" & ChrW(9) & totDistilled & " Kb")
        outputStream.WriteLine("Total Original Value:" & ChrW(9) & totOrig & " Kb")
        outputStream.WriteLine("Total Savings Value:" & ChrW(9) & totSavings & " Kb")
        outputStream.WriteLine("*****************************************************************************")
        outputStream.WriteLine()
        outputStream.WriteLine("#############################################################################")
        outputStream.WriteLine()
        outputStream.WriteLine()
        outputStream.WriteLine()
        outputStream.Close()
    End Sub

End Module
