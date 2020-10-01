Imports System.Data.SqlClient
Imports System.IO

Module Module1
    Public conCollation As SqlConnection = New SqlConnection("Server=FL1CDULUDBSA01;Database=Collation;User Id=Svc_uluro;Password=Th3ultim@tew@y;MultipleActiveResultSets=true")
    Dim fs As FileStream = Nothing
    Dim sw As StreamWriter

    Sub Main()
        If conCollation.State = ConnectionState.Closed Then
            conCollation.Open()
        End If
        Dim previousSSN As String

        Dim strFile As String = "C:\W2 Test\Output.txt"

        Dim commandRequestAccessLevel As SqlCommand
        Dim commandRequest2 As SqlCommand
        Dim maxCopies As Integer


        '************* CODE TO WRITE W2S to Datafile ******************'
        Try
            commandRequestAccessLevel = New SqlCommand("SELECT * FROM DBO.W2", conCollation)
            commandRequestAccessLevel.ExecuteNonQuery()

            Dim objDataReaderRequest As SqlDataReader = commandRequestAccessLevel.ExecuteReader()

            If objDataReaderRequest.HasRows Then
                Console.WriteLine("Has Rows")

                Dim colCount As Integer

                colCount = objDataReaderRequest.FieldCount
                Console.WriteLine("FieldCount: " + colCount.ToString)


                Do While objDataReaderRequest.Read()
                    If (Not File.Exists(strFile)) Then
                        Try
                            fs = File.Create(strFile)
                            fs.Close()
                        Catch ex As Exception
                        End Try
                    Else
                        sw = File.AppendText(strFile)

                        For index As Integer = 1 To 54
                            If index <> 52 Then
                                sw.Write(objDataReaderRequest(index).ToString + "~")
                            End If
                        Next

                        'Console.WriteLine("currentCopy: " + objDataReaderRequest(objDataReaderRequest.GetOrdinal("currentCopy")).ToString)
                        ' If objDataReaderRequest(objDataReaderRequest.GetOrdinal("hasMultipleCopies")) = "True" Then

                        ' Do Not Make New Row

                        ' GET Max Copies Count
                        ' Try
                        'commandRequest2 = New SqlCommand("SELECT MAX(CURRENTCOPY) FROM DBO.W2 WHERE EE_EmployeeSSN = '" + objDataReaderRequest(objDataReaderRequest.GetOrdinal("EE_EmployeeSSN")) + "' AND Client_ID = '" + objDataReaderRequest(objDataReaderRequest.GetOrdinal("Client_ID")) + "' AND TAX_CODE ='" + objDataReaderRequest(objDataReaderRequest.GetOrdinal("Tax_Code")) + "'", conCollation)
                        'commandRequest2.ExecuteNonQuery()
                        '
                        'Dim objDataReaderRequest2 As SqlDataReader = commandRequest2.ExecuteReader()
                        '
                        '                       If objDataReaderRequest2.HasRows Then
                        '                      Do While objDataReaderRequest2.Read()
                        '                     maxCopies = objDataReaderRequest2(0)
                        '                'Console.WriteLine("Max Copies:" + maxCopies.ToString)
                        '           Loop
                        '                  Else
                        '             End If
                        'Catch ex As Exception
                        '               Console.WriteLine("Exception: " + ex.ToString)
                        'End Try
                        '
                        '       If objDataReaderRequest(objDataReaderRequest.GetOrdinal("currentCopy")) = maxCopies Then
                        '      Console.WriteLine("Reached end of Combined Line")
                        '     sw.WriteLine()
                        'End If

                        'Else
                        '   maxCopies = 1
                        sw.WriteLine()
                            'End If

                            sw.Close()
                    End If
                Loop
            Else
            End If

            objDataReaderRequest.Close()
            commandRequestAccessLevel.Dispose()

        Catch ex As Exception
            Console.WriteLine("Exception: " + ex.ToString)
        End Try

    End Sub
End Module
