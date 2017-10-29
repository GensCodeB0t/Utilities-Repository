Imports System
Imports System.IO
Imports System.Printing
Imports System.Printing.IndexedProperties


Module Module1
    Private printSever As New LocalPrintServer()
    Private printQueue As PrintQueue
    Private fileName As String
    Private printJob As PrintSystemJobInfo

    Sub Main()
        Console.WriteLine("Enter File Name")
        fileName = Console.ReadLine()
        fileName = fileName.Trim().ToUpper()
        Console.WriteLine("File Name = " & fileName)
        printQueue = printSever.GetPrintQueue("PrintQueue")
        WaitForJob()
        'MonitorQueue()
        Console.WriteLine("END...")
        Dim q As String = Console.Read()

    End Sub

    Private Function MonitorQueue()
        Dim printSever As New LocalPrintServer()
        Dim printQueue As PrintQueue = printSever.GetPrintQueue("PrintQueue")
        Dim numOfJobs As Integer = printQueue.NumberOfJobs()

        Console.WriteLine(numOfJobs & " jobs are in the queue")

        Dim jobStatus = printJob.JobStatus
        Do While printJob.JobStatus <> PrintJobStatus.Completed
            If jobStatus <> printJob.JobStatus Then
                Console.WriteLine(fileName & " is currently in status " & printJob.JobStatus)
                jobStatus = printJob.JobStatus
            End If
        Loop
        Console.WriteLine("Print Job Complete!")

        Return True
    End Function

    Private Sub WaitForJob()
        Console.WriteLine("Waiting for job...")
        Do Until printQueue.NumberOfJobs > 0
            ' Waiting for a job to start...
        Loop
        For Each job As PrintSystemJobInfo In printQueue.GetPrintJobInfoCollection
            Console.WriteLine("JobName: " & job.JobName)
            Console.WriteLine("Name: " & job.Name)
            Console.WriteLine("JobIdentifier: " & job.JobIdentifier)

        Next
        Console.WriteLine(fileName & " Has Started...")
    End Sub



End Module
