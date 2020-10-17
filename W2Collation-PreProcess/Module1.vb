Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports Ionic.Zip
Imports ZipFile = Ionic.Zip.ZipFile

Module Module1
    Public conCollation As SqlConnection = New SqlConnection("Server=FL1CDULUDBSA01;Database=Collation;User Id=Svc_uluro;Password=Th3ultim@tew@y;MultipleActiveResultSets=true")
    Dim fs As FileStream = Nothing
    Dim fsLog As FileStream = Nothing
    Dim sw As StreamWriter
    Dim swLog As StreamWriter

    Sub Main()
        Dim fullLine As String
        Dim lineNumber As Integer = 0
        Dim Tax_Code As String = ""
        Dim Client_ID As String = ""
        Dim SplitRow() As String
        Dim rowW2Count As Integer = 0
        Dim lineSB As New StringBuilder
        Dim firstColumnIndex As Integer
        Dim returnData As String
        Dim processingFolder As String = "C:\W2 Test\"


        Dim FPath As String
        Dim DataFileName As String
        Dim submid As String
        Dim clientID_Col As Integer
        Dim taxcode_Col As Integer

        Dim clArgs() As String = Environment.GetCommandLineArgs()

        Try
            FPath = clArgs(1)
            DataFileName = clArgs(2)
            submid = clArgs(3)
            clientID_Col = clArgs(4)
            taxcode_Col = clArgs(5)
        Catch ex As Exception
            Console.WriteLine("Missing Command Line Args")
        End Try

        Dim zippedFile As String
        zippedFile = FPath + DataFileName

        'WriteLog("DataFileName:" + DataFileName)

        Try

            Dim extractPath = processingFolder + "Unzip\Complete\" + Guid.NewGuid.ToString + "\" + DataFileName

            System.IO.File.Copy(zippedFile, processingFolder + "Unzip\" + DataFileName)
            WriteLog("Completed Copy to Temp Folder")


            Console.WriteLine("Trying to Unzip: " + zippedFile)
            Dim archive As ZipFile = New ZipFile(processingFolder + "Unzip\" + DataFileName)

            archive.Password = "2D3crypt2017!"

            WriteLog("Start Unzip:" + DataFileName)
            archive.ExtractAll(extractPath)


            Dim FileLocation As DirectoryInfo = New DirectoryInfo(extractPath)

            Dim fi As FileInfo() = FileLocation.GetFiles("*.csv")

            For Each afile In fi
                Console.WriteLine("Found File in Temp")
                System.IO.File.Delete("E:\Submit\Data\" + DataFileName)
                System.IO.File.Move(afile.FullName, "E:\Submit\Data\" + DataFileName)
            Next

            Console.WriteLine("Deleted: " + extractPath)

            archive.Dispose()

            System.IO.File.Delete(processingFolder + "Unzip\" + DataFileName)
            'System.IO.File.Delete(extractPath)

        Catch
            Console.WriteLine("ERROR")
        End Try



        'UnzipReplace(FPath + DataFileName, "C:\W2 Test\Unzip", DataFileName)


        If conCollation.State = ConnectionState.Closed Then
            conCollation.Open()
        End If

        Dim strFile As String = processingFolder + "Output.txt"
        File.Delete(strFile)

        Dim commandRequest As SqlCommand

        Dim FILE_NAME As String = FPath + DataFileName

        If System.IO.File.Exists(FILE_NAME) = True Then

            'Console.WriteLine(FILE_NAME + " Found...")

            Dim objReader As New System.IO.StreamReader(FILE_NAME)
            'Console.WriteLine("Starting Employee Count: " + employeeCount.ToString)
            Console.WriteLine("Processing Started: " + FILE_NAME)
            Do While objReader.Peek() <> -1
                rowW2Count = 0
                lineSB.Clear()
                'aEmployeeRow = ""
                'aEmployeeStateRow = ""
                'aEmployeeBox12Row = ""
                'aEmployeeBox14Row = ""
                lineNumber += 1
                Dim reachedEnd As Boolean = False
                fullLine = objReader.ReadLine()

                'SKIP HEADER ROW'
                If lineNumber > 1 Then
                    SplitRow = fullLine.Split("~")

                    Tax_Code = SplitRow(taxcode_Col)
                    Client_ID = SplitRow(clientID_Col)



                    If (Not File.Exists(strFile)) Then
                        Try
                            fs = File.Create(strFile)
                            fs.Close()
                        Catch ex As Exception
                        End Try
                    Else
                        firstColumnIndex = fullLine.IndexOf("~")
                        sw = File.AppendText(strFile)
                        returnData = fullLine

                        sw.Close()
                    End If

                    sw = File.AppendText(strFile)

                    '************* CODE TO WRITE W2S to Datafile ******************'
                    If 0 = 0 Then
                        Try
                            commandRequest = New SqlCommand("SELECT * FROM DBO.W2 WHERE Tax_Code ='" + Tax_Code + "' AND Client_ID ='" + Client_ID + "' ORDER BY EE_LASTNAME, EE_FIRSTNAME, EE_EmployeeSSN", conCollation)
                            commandRequest.ExecuteNonQuery()


                            Dim commandRequest2 As SqlCommand

                            Dim objDataReaderRequest As SqlDataReader = commandRequest.ExecuteReader()

                            If objDataReaderRequest.HasRows Then

                                Do While objDataReaderRequest.Read()

                                    WriteLog("Running W2Pull")
                                    Try

                                        commandRequest2 = New SqlCommand("EXEC setW2Pulled '" + objDataReaderRequest(objDataReaderRequest.GetOrdinal("tax_code")).ToString + "','" + objDataReaderRequest(objDataReaderRequest.GetOrdinal("client_id")).ToString + "','" + objDataReaderRequest(objDataReaderRequest.GetOrdinal("EE_EmployeeSSN")).ToString + "','" + submid + "','" + objDataReaderRequest(objDataReaderRequest.GetOrdinal("id")).ToString + "'", conCollation)
                                        commandRequest2.ExecuteNonQuery()
                                    Catch ex As Exception
                                        WriteLog("Exception: " + ex.Message)
                                    End Try
                                    WriteLog("Ran W2Pull")

                                    lineSB.Clear()
                                    'Console.WriteLine("Has Rows")

                                    Dim colCount As Integer

                                    colCount = objDataReaderRequest.FieldCount
                                    '                                    Console.WriteLine("FieldCount: " + colCount.ToString)

                                    '                                    Console.WriteLine("Got Client_ID " + objDataReaderRequest(objDataReaderRequest.GetOrdinal("Client_ID")))


                                    '                                Do While objDataReaderRequest.Read()
                                    '  '                               If (Not File.Exists(strFile)) Then
                                    '                              Try
                                    '                             fs = File.Create(strFile)
                                    '                            fs.Close()
                                    '                   Catch ex As Exception
                                    '                  End Try
                                    '                     Else
                                    '                        sw = File.AppendText(strFile)


                                    'lineSB.Append(returnData)

                                    For index As Integer = 1 To 54

                                        If index <> 52 Then
                                            If index = 9 OrElse index = 20 OrElse index = 14 OrElse index = 15 OrElse index = 16 Then
                                                lineSB.Append("~" + objDataReaderRequest(index).ToString.TrimEnd(" "))
                                            Else
                                                lineSB.Append("~" + objDataReaderRequest(index).ToString)
                                            End If
                                            'Console.WriteLine(objDataReaderRequest(1).ToString)
                                        End If
                                    Next

                                    'If objDataReaderRequest(objDataReaderRequest.GetOrdinal("currentCopy")) = 1 Then
                                    'sw.WriteLine("~" + Client_ID.Replace("_", "") + Tax_Code + lineSB.ToString)
                                    'Else
                                    sw.WriteLine(Client_ID + Tax_Code + "~" + returnData + "~" + Client_ID + Tax_Code + lineSB.ToString)
                                    'End If


                                    rowW2Count = rowW2Count + 1
                                Loop
                            Else
                                WriteLog("No W2 Found for ClientID:" + Client_ID + " Taxcode: " + Tax_Code)
                                'sw.WriteLine(Client_ID.Replace("_", "") + Tax_Code + fullLine.Substring(firstColumnIndex))
                                sw.WriteLine(Client_ID + Tax_Code + "~" + fullLine)


                                rowW2Count = 0
                            End If

                            If rowW2Count = 0 Then
                                'Console.WriteLine("Client_ID:  " + Client_ID + " has no rows")
                            Else
                                'Console.WriteLine("Client_ID: " + Client_ID + " has " + rowW2Count.ToString, +" rows")
                            End If


                            objDataReaderRequest.Close()
                            commandRequest.Dispose()
                            commandRequest2.Dispose()


                        Catch ex As Exception
                            Console.WriteLine("Exception: " + ex.ToString)
                        End Try
                    End If

                    sw.Close()

                    'Console.WriteLine("Row " + lineNumber.ToString + " Taxcode: " + Tax_Code + " Client ID: " + Client_ID.Replace("_", ""))
                    '                    Console.WriteLine("EXEC pullW2 '" + Client_ID.Replace("_", "") + "' '" + Tax_Code + "'")
                Else
                    sw = File.AppendText(strFile)
                    sw.WriteLine("~" + fullLine + "~~Tax_Code~Client_ID~ARNumber~TaxYear~EIN~EmployerName~LocationAddress~DeliveryAddress~City~State~Zip~ZipCodeExtension~EE_EmployeeSSN~EE_FirstName~EE_LastName~EE_MiddleName~EE_Suffix~EE_LocationAddress~EE_DeliveryAddress~EE_City~EE_State~EE_Zip~EE_Zip_Code_Extension~Box1~Box2~Box3~Box4~Box5~Box6~Box7~Box8~Box9~Box10~Box11~Box12_1~Box12_2~Box12_3~Box12_4~Box13~Box14_1~Box14_2~Box14_3~Box14_4~Box14_5~EE_Retirement_Plan~EE_Third_Party_Sick_Pay~Box16~Box17~Box18~Box19~Box20~instance~currentcopy~hasmultiplecopies")

                    sw.Close()
                End If

                If reachedEnd = False Then
                    'MessageBox.Show("Test")
                    '   Console.WriteLine("Beginning to Read Line")
                End If
                If objReader.Peek() = -1 Then
                    reachedEnd = True
                End If
                'Console.WriteLine("End Peek")


            Loop
            'Console.WriteLine("Final Row:")
            'Console.WriteLine(W2DataArray(EmployeeIndex).EEData)
            'Console.WriteLine("Number of Employees:" + employeeCount.ToString)
            objReader.Close()
        Else
        End If
        Try
            System.IO.File.Delete("E:\Submit\DATA\" + DataFileName)
            System.IO.File.Copy(processingFolder + "Output.txt", "E:\SUBMIT\DATA\" + DataFileName)
            WriteLog("Wrote Collated File Back to E Submit")
        Catch ex As Exception
            WriteLog("Error Writting Back " + DataFileName + " :" + ex.Message)
        End Try


    End Sub

    Sub UnzipReplace(ByVal zippedFileName As String, ByVal extractPath As String, ByVal newFileName As String)
        'ZipFile.CreateFromDirectory(startPath, zipPath)

        'ZipFile.ExtractToDirectory(zippedFileName, extractPath)

        Dim FileLocation As DirectoryInfo = New DirectoryInfo(extractPath)

        Dim fi As FileInfo() = FileLocation.GetFiles("*.csv")

        For Each afile In fi
            System.IO.File.Move(afile.FullName, "E:\Submit\Data\" + newFileName)
        Next
    End Sub
    Private Sub WriteLog(message As String)
        Dim strFile As String = "C:\Ultimate Software Print Services - ADHOC Service\Logs\W2Collation Log " + DateTime.Now.ToString("MM-dd-yyyy") + ".txt"
        If (Not File.Exists(strFile)) Then
            Try
                fsLog = File.Create(strFile)
                fsLog.Close()
            Catch ex As Exception
            End Try
        Else
            swLog = File.AppendText(strFile)
            swLog.WriteLine(DateTime.Now + " - " + message)
            swLog.Close()
        End If
    End Sub
End Module
