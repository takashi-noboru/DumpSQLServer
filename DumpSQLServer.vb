Module modMasterDB

    Public Sub DumpSQL(ByVal sFilename As String)
        Dim dtColumns As DataTable
        Dim drTables As DataRow
        Dim drColumns As DataRow
        Dim drRows As DataRow
        Dim sSQL As String
        Dim bIdentify As Boolean
        Dim cnt As Long
        Dim sFields As String
        Dim sValues As String
        Dim iTmp As Long

        Dim sw As New System.IO.StreamWriter(sFilename, False, System.Text.Encoding.GetEncoding("utf-8"))

        sSQL = "SELECT name,object_id FROM Sys.Tables ORDER BY name"
        For Each drTables In gDB.ExecuteSql(sSQL).Rows
            sw.WriteLine("DROP TABLE " & drTables.Item("name") & ";")
            sw.WriteLine("CREATE TABLE " & drTables.Item("name") & " (")

            bIdentify = False
            sFields = ""
            sSQL = "SELECT Sys.Columns.*," & vbCrLf
            sSQL &= "TYPE_NAME(Columns.system_type_id) AS column_data_type, " & vbCrLf
            sSQL &= "defaultConstraints.definition     AS column_default " & vbCrLf
            sSQL &= "FROM Sys.Tables AS Tables " & vbCrLf
            sSQL &= "INNER JOIN Sys.Columns ON " & vbCrLf
            sSQL &= "   Tables.object_id = Sys.Columns.object_id" & vbCrLf
            sSQL &= "LEFT OUTER JOIN Sys.Sysconstraints AS Constraints ON" & vbCrLf
            sSQL &= "    Columns.object_id = Constraints.id AND" & vbCrLf
            sSQL &= "    Columns.column_id = Constraints.colid AND" & vbCrLf
            sSQL &= "    (Constraints.status & 2069) = 2069" & vbCrLf
            sSQL &= "LEFT OUTER JOIN Sys.default_constraints defaultConstraints ON" & vbCrLf
            sSQL &= "    Constraints.constid = defaultConstraints.object_id AND" & vbCrLf
            sSQL &= "    Tables.schema_id    = defaultConstraints.schema_id" & vbCrLf
            sSQL &= "WHERE Sys.Columns.object_id = " & drTables.Item("object_id") & " ORDER BY column_id"
            dtColumns = gDB.ExecuteSql(sSQL)
            For Each drColumns In dtColumns.Rows
                sFields &= "[" & drColumns.Item("name") & "],"
                sw.Write("[" & drColumns.Item("name") & "] ")
                sw.Write("[" & drColumns.Item("column_data_type") & "] ")
                Select Case drColumns.Item("system_type_id")
                    Case 35 ' text
                    Case 40 ' Date
                    Case 42 ' Datetime2
                    Case 48 ' tinyint
                    Case 52 ' smallint
                    Case 56 ' int
                    Case 59 ' real
                    Case 61 ' datetime
                    Case 104 ' Bit
                    Case 127 ' Bigint
                    Case 106, 108 ' Decimal,Numeric
                        sw.Write("(" & drColumns.Item("precision") & "," & drColumns.Item("scale") & ") ")
                    Case Else
                        sw.Write("(" & drColumns.Item("max_length") & ") ")
                End Select


                If drColumns.Item("is_identity") Then
                    bIdentify = True
                    sw.Write("IDENTITY(1,1) ")
                End If
                If drColumns.Item("is_nullable") Then
                    sw.Write("NULL ")
                Else
                    sw.Write("NOT NULL ")
                End If
                If Not IsDBNull(drColumns.Item("column_default")) Then
                    sw.Write("DEFAULT " & drColumns.Item("column_default"))
                End If
                sw.WriteLine(",")
                sw.Write("")

            Next
            sw.WriteLine(") ON [PRIMARY];")

            sFields = sFields.Substring(0, sFields.Length - 1)


            sSQL = "SELECT * FROM " & drTables.Item("name") & vbCrLf
            cnt = 0
            For Each drRows In gDB.ExecuteSql(sSQL).Rows
                If (cnt Mod IIf(dtColumns.Rows.Count < 15, 50, 10)) = 0 Then
                    If cnt Then
                        sw.WriteLine(";")
                        If bIdentify Then
                            sw.WriteLine("SET IDENTITY_INSERT " & drTables.Item("name") & " OFF;")
                        End If
                    End If
                    If bIdentify Then
                        sw.WriteLine("SET IDENTITY_INSERT " & drTables.Item("name") & " ON;")
                    End If
                    sw.WriteLine("INSERT INTO " & drTables.Item("name") & " (" & sFields & ") VALUES")
                Else
                    sw.WriteLine(",")
                End If


                sw.Write("(")
                sValues = ""
                For Each drColumns In dtColumns.Rows
                    If IsDBNull(drRows.Item(drColumns.Item("name"))) Then
                        sValues &= "NULL"
                    Else
                        Select Case drColumns.Item("system_type_id")
                            Case 35 ' text
                                sValues &= "'" & drRows.Item(drColumns.Item("name")) & "'"
                            Case 40 ' Date
                                sValues &= "'" & drRows.Item(drColumns.Item("name")) & "'"
                            Case 42 ' Datetime2
                                sValues &= "'" & drRows.Item(drColumns.Item("name")) & "'"
                            Case 48 ' tinyint
                                sValues &= drRows.Item(drColumns.Item("name"))
                            Case 52 ' smallint
                                sValues &= drRows.Item(drColumns.Item("name"))
                            Case 56 ' int
                                sValues &= drRows.Item(drColumns.Item("name"))
                            Case 59 ' real
                                sValues &= drRows.Item(drColumns.Item("name"))
                            Case 61 ' Datetime
                                sValues &= "'" & drRows.Item(drColumns.Item("name")) & "'"
                            Case 104 ' Bit
                                sValues &= IIf(drRows.Item(drColumns.Item("name")), 1, 0)
                            Case 127 ' Bigint
                                sValues &= drRows.Item(drColumns.Item("name"))
                            Case 173 ' Binary
                                iTmp = 0
                                For i = 0 To CType(drRows.Item(drColumns.Item("name")), Array).Length - 1
                                    iTmp *= 256
                                    iTmp += (CType(drRows.Item(drColumns.Item("name")), Array)(i))
                                Next
                                sValues &= iTmp
                            Case 106 , 108 ' Decimal , Numeric
                                sValues &= drRows.Item(drColumns.Item("name"))
                            Case Else
                                sValues &= "'" & drRows.Item(drColumns.Item("name")) & "'"
                        End Select
                    End If
                    sValues &= ","
                Next
                sValues = sValues.Substring(0, sValues.Length - 1)

                sw.Write(sValues & ")")



                cnt += 1
            Next
            If cnt Then
                sw.WriteLine(";")
                If bIdentify Then
                    sw.WriteLine("SET IDENTITY_INSERT " & drTables.Item("name") & " OFF;")
                End If
            End If


        Next
        sw.Close()
        sw = Nothing


    End Sub
End Module
