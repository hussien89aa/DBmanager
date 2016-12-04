Imports System.Data
Imports System.Data.OracleClient

Public Enum ColoumnType As Integer

    Int16 = 1
    Int32 = 2
    UInt16 = 3
    UInt32 = 4
    Float = 5
    Double1 = 6
    BFile = 7
    Char1 = 8
    DateTime = 9
    varchar = 10
    NChar = 11
    NVarChar = 12
    Number = 13
    Byte1 = 14
    SByte1 = 15

End Enum
Public Class ColoumnParam

    Public Property ColoumnsName As String
    Public Property ColoumnsType As ColoumnType
    Public Property ColoumnsData As Object
    Public Sub New(ByVal ColoumnName As String, ByVal ColoumnType As ColoumnType, ByVal ColoumnData As Object)
        ColoumnsName = ColoumnName
        ColoumnsType = ColoumnType
        ColoumnsData = ColoumnData
    End Sub



End Class
Public Class ColTabelParam

    Public Property ColoumnsName As String
    Public Property ColoumnsType As ColoumnType
    Public Property ColumnsSize As Integer
    Public Property NotNull As Boolean
    Public Property AutoIncrement As Boolean
    Public Property PrimaryKey As Boolean
    Public Sub New(ByVal ColoumnName As String, ByVal ColoumnType As ColoumnType, Optional ByVal Columns_Size As Integer = 1, Optional ByVal NOT_NULL As Boolean = False, Optional ByVal PRIMARY_KEY As Boolean = False, Optional ByVal AUTO_INCREMENT As Boolean = False)
        ColoumnsName = ColoumnName
        ColoumnsType = ColoumnType
        ColumnsSize = Columns_Size
        NotNull = NOT_NULL
        PrimaryKey = PRIMARY_KEY
        AutoIncrement = AUTO_INCREMENT

    End Sub

End Class

Public Class DBOpeartion
    Public Property connectionString As String ' Private mean it declarted only on it page we use Public becaouse it may shange it connection string
    Private Property SQLString As String   ' the sql string we userd in database
    Private Property ErrorHappend As String  ' erro happned in data base 
    Public ReadOnly Property SQLComand() As String
        Get
            Return (SQLString)
        End Get
    End Property
    Public ReadOnly Property ErrorMessage() As String
        Get
            Return (ErrorHappend)
        End Get
    End Property

    Public Sub New(ByVal connectionStringDB As String)
        connectionString = connectionStringDB
    End Sub
    ''' <summary>
    ''' '''''''''''' tabel inside modfiction
    ''' </summary>
    ''' <param name="TabelName"></param>
    ''' <param name="coloumns"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function InsertRow(ByVal TabelName As String, ByVal coloumns() As ColoumnParam) As Boolean
        Try
            Dim query As String = "INSERT INTO " & TabelName & " ("
            Dim SecondQuery As String = "VALUES ("

            For i = 0 To coloumns.Length - 1
                query = query & coloumns(i).ColoumnsName & ","
                SecondQuery = SecondQuery & "@" & coloumns(i).ColoumnsName & ","
            Next
            query = Mid(query, 1, Len(query) - 1) & ") " & Mid(SecondQuery, 1, Len(SecondQuery) - 1) & ") "

            SQLString = query ' the sql string we userd in database
            Using conn As New OracleConnection(connectionString)
                Using comm As New OracleCommand()
                    With comm
                        .Connection = conn
                        .CommandType = CommandType.Text
                        .CommandText = query
                        For i = 0 To coloumns.Length - 1

                            If coloumns(i).ColoumnsType = ColoumnType.Int16 Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OracleType.Int16)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = Convert.ToInt16(coloumns(i).ColoumnsData)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Int32 Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OracleType.Int32)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = Convert.ToInt32(coloumns(i).ColoumnsData)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.UInt16 Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OracleType.UInt16)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = Convert.ToUInt16(coloumns(i).ColoumnsData)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.UInt32 Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OracleType.UInt32)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = Convert.ToUInt32(coloumns(i).ColoumnsData)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Float Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OracleType.Float)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Double)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Double1 Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OracleType.Double)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Double)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.BFile Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OracleType.BFile)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Byte())

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Char1 Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OracleType.Char)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Char)


                            ElseIf coloumns(i).ColoumnsType = ColoumnType.DateTime Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OracleType.DateTime)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = Convert.ToDateTime(coloumns(i).ColoumnsData)


                            ElseIf coloumns(i).ColoumnsType = ColoumnType.varchar Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OracleType.VarChar)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, String)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.NChar Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OracleType.NChar)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, String)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.NVarChar Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OracleType.NVarChar)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, String)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Number Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OracleType.Number)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = Convert.ToDecimal(coloumns(i).ColoumnsData)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Byte1 Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OracleType.Byte)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Byte)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.SByte1 Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OracleType.SByte)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = Convert.ToSByte(coloumns(i).ColoumnsData)


                            End If

                        Next


                    End With
                    Try

                        conn.Open()
                        comm.ExecuteNonQuery()
                        conn.Close()
                        Return (True)
                    Catch ex1 As Exception '   inserted in db
                        ErrorHappend = ex1.Message   ' properties of display messag error
                        Return (False)
                    End Try

                End Using
            End Using

        Catch ex As Exception ' if he error enter  
            ErrorHappend = ex.Message   ' properties of display messag error
            Return (False)
        End Try


    End Function

    Public Function UpdateRow(ByVal TabelName As String, ByVal coloumns() As ColoumnParam, ByVal Condition As String) As Boolean
        Try
            Dim query As String = "UPDATE " & TabelName & " set "

            For i = 0 To coloumns.Length - 1
                query = query & coloumns(i).ColoumnsName & "=@" & coloumns(i).ColoumnsName & ","

            Next
            query = Mid(query, 1, Len(query) - 1) & " Where " & Condition

            SQLString = query  ' the sql string we userd in database
            Using conn As New OracleConnection(connectionString)
                Using comm As New OracleCommand()
                    With comm
                        .Connection = conn
                        .CommandType = CommandType.Text
                        .CommandText = query

                        For i = 0 To coloumns.Length - 1

                            If coloumns(i).ColoumnsType = ColoumnType.Int16 Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OracleType.Int16)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = Convert.ToInt16(coloumns(i).ColoumnsData)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Int32 Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OracleType.Int32)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = Convert.ToInt32(coloumns(i).ColoumnsData)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.UInt16 Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OracleType.UInt16)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = Convert.ToUInt16(coloumns(i).ColoumnsData)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.UInt32 Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OracleType.UInt32)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = Convert.ToUInt32(coloumns(i).ColoumnsData)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Float Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OracleType.Float)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Double)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Double1 Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OracleType.Double)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Double)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.BFile Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OracleType.BFile)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Byte())

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Char1 Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OracleType.Char)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Char)


                            ElseIf coloumns(i).ColoumnsType = ColoumnType.DateTime Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OracleType.DateTime)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = Convert.ToDateTime(coloumns(i).ColoumnsData)


                            ElseIf coloumns(i).ColoumnsType = ColoumnType.varchar Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OracleType.VarChar)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, String)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.NChar Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OracleType.NChar)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, String)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.NVarChar Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OracleType.NVarChar)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, String)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Number Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OracleType.Number)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = Convert.ToDecimal(coloumns(i).ColoumnsData)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Byte1 Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OracleType.Byte)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Byte)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.SByte1 Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OracleType.SByte)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = Convert.ToSByte(coloumns(i).ColoumnsData)


                            End If

                        Next


                    End With
                    Try

                        conn.Open()
                        comm.ExecuteNonQuery()
                        conn.Close()
                        Return (True)
                    Catch ex1 As Exception '   inserted in db
                        ErrorHappend = ex1.Message   ' properties of display messag error
                        Return (False)
                    End Try

                End Using
            End Using

        Catch ex As Exception ' if he error enter  
            ErrorHappend = ex.Message   ' properties of display messag error
            Return (False)
        End Try


    End Function
    Public Function DeletedRow(ByVal TabelName As String, ByVal Condtion As String) As Boolean

        Try
            Dim query As String = "DELETE FROM " & Trim(TabelName) & " WHERE " & Trim(Condtion)
            SQLString = query  ' the sql string we userd in database
            Using conn As New OracleConnection(connectionString)
                Using comm As New OracleCommand()
                    With comm
                        .Connection = conn
                        .CommandType = CommandType.Text
                        .CommandText = query


                    End With
                    Try

                        conn.Open()
                        comm.ExecuteNonQuery()
                        conn.Close()
                        Return (True)
                    Catch ex1 As Exception '   inserted in db
                        ErrorHappend = ex1.Message   ' properties of display messag error
                        Return (False)
                    End Try

                End Using
            End Using

        Catch ex As Exception
            ErrorHappend = ex.Message   ' properties of display messag error
            Return (False)
        End Try
    End Function

    Public Function SelectDataSet(ByVal TabelName As String, Optional ByVal coloumns As String = "*", Optional ByVal Condition As String = Nothing, Optional ByVal ColoumnsOrder As String = Nothing, Optional ByVal ColumnsGROUPBY As String = Nothing, Optional ByVal ColumnsHAVING As String = Nothing) As DataSet
        Dim SDataset As New DataSet
        SDataset.Clear()
        Try

            Dim query As String = "SELECT " & Trim(coloumns) & " FROM " & Trim(TabelName)
            If IsNothing(Condition) = False Then  ' condtions
                query = query & " where " & Trim(Condition)
            End If
            If IsNothing(ColoumnsOrder) = False Then  ' order by
                query = query & " order by " & Trim(ColoumnsOrder)
            End If  '
            If IsNothing(ColumnsGROUPBY) = False Then   ' group by
                query = query & " GROUP BY " & Trim(ColumnsGROUPBY)
            End If

            If IsNothing(ColumnsHAVING) = False Then   ' '
                query = query & " HAVING " & Trim(ColumnsHAVING)
            End If
            SQLString = query  ' the sql string we userd in database
            Using conn As New OracleConnection(connectionString)
                Using comm As New OracleCommand()
                    With comm
                        .Connection = conn
                        .CommandType = CommandType.Text
                        .CommandText = query


                    End With

                    Try

                        conn.Open()
                        Dim DataAdapter1 As New OracleDataAdapter(comm)
                        DataAdapter1.Fill(SDataset)
                        conn.Close()
                    Catch ex1 As Exception  ' if he not not match return tabel but it clear
                        ErrorHappend = ex1.Message   ' properties of display messag error
                        Dim TBUse As New DataTable("Tabel1")
                        SDataset.Tables.Add(TBUse)


                    End Try


                End Using
            End Using
        Catch ex As Exception
            ErrorHappend = ex.Message   ' properties of display messag error
            Dim TBUse As New DataTable("Tabel1")
            SDataset.Tables.Add(TBUse)
        End Try
        Return (SDataset)
    End Function

    Public Function SelectCell(ByVal TabelName As String, ByVal SearchColoumnName As String, ByVal Condtion As String) As String
        Dim SDataset As New DataSet
        SDataset.Clear()
        Try
            Dim query As String = "SELECT " & Trim(SearchColoumnName) & " FROM " & Trim(TabelName) & " where " & Trim(Condtion)
            SQLString = query  ' the sql string we userd in database
            Using conn As New OracleConnection(connectionString)
                Using comm As New OracleCommand()
                    With comm
                        .Connection = conn
                        .CommandType = CommandType.Text
                        .CommandText = query


                    End With

                    Try

                        conn.Open()
                        Dim DataAdapter1 As New OracleDataAdapter(comm)
                        DataAdapter1.Fill(SDataset)
                        conn.Close()
                        Return (SDataset.Tables(0).Rows(0).Item(Trim(SearchColoumnName)).ToString)
                    Catch ex1 As Exception  ' if he not not match return tabel but it clear
                        ErrorHappend = ex1.Message   ' properties of display messag error
                        Return (Nothing)

                    End Try


                End Using
            End Using
        Catch ex As Exception
            ErrorHappend = ex.Message   ' properties of display messag error
            Return (Nothing)
        End Try
    End Function
    '--------------------------------------------------------------------------
    ''''''''''''''' data base creat deleted tabel add deleted
    Public Function CreateNewTable(ByVal TabelName As String, ByVal coloumns() As ColTabelParam) As Boolean
        Try
            Dim query As String = "CREATE TABLE " & TabelName & " ("

            For i = 0 To coloumns.Length - 1


                If coloumns(i).ColoumnsType = ColoumnType.Int16 Then

                    query = query & coloumns(i).ColoumnsName & " Int16"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.Int32 Then

                    query = query & coloumns(i).ColoumnsName & " Int32"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.UInt16 Then

                    query = query & coloumns(i).ColoumnsName & " UInt16"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.UInt32 Then

                    query = query & coloumns(i).ColoumnsName & " UInt32"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.Float Then

                    query = query & coloumns(i).ColoumnsName & " Float"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.Double1 Then

                    query = query & coloumns(i).ColoumnsName & " Double"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.BFile Then

                    query = query & coloumns(i).ColoumnsName & " BFile"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.Char1 Then

                    query = query & coloumns(i).ColoumnsName & " Char"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.DateTime Then

                    query = query & coloumns(i).ColoumnsName & " DateTime"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.varchar Then '*************

                    query = query & coloumns(i).ColoumnsName & " VarChar" & "(" & coloumns(i).ColumnsSize & ")"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.NChar Then '*************

                    query = query & coloumns(i).ColoumnsName & " NChar" & "(" & coloumns(i).ColumnsSize & ")"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.NVarChar Then '*************

                    query = query & coloumns(i).ColoumnsName & " NVarChar" & "(" & coloumns(i).ColumnsSize & ")"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.Number Then '*************

                    query = query & coloumns(i).ColoumnsName & " Number"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.Byte1 Then '*************

                    query = query & coloumns(i).ColoumnsName & " Byte"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.SByte1 Then '*************

                    query = query & coloumns(i).ColoumnsName & " SByte"
                End If
                '' addons 
                If coloumns(i).NotNull = True Then
                    query = query & " NOT NULL"
                End If

                If coloumns(i).PrimaryKey = True Then '
                    query = query & " PRIMARY KEY"
                End If
                If coloumns(i).AutoIncrement = True Then   ' for acees  (AUTOINCREMENT)  for oracle ( SEQUENCE seq_person MINVALUE 1 START WITH 1 INCREMENT BY 1)
                    query = query & " SEQUENCE seq_person MINVALUE 1 START WITH 1 INCREMENT BY 1"
                End If
                query = query & ","

            Next
            query = Mid(query, 1, Len(query) - 1) & ") "

            SQLString = query  ' the sql string we userd in database
            Using conn As New OracleConnection(connectionString)
                Using comm As New OracleCommand()
                    With comm
                        .Connection = conn
                        .CommandType = CommandType.Text
                        .CommandText = query

                    End With
                    Try

                        conn.Open()
                        comm.ExecuteNonQuery()
                        conn.Close()
                        Return (True)
                    Catch ex1 As Exception '   inserted in db
                        ErrorHappend = ex1.Message   ' properties of display messag error
                        Return (False)
                    End Try

                End Using
            End Using

        Catch ex As Exception ' if he error enter  
            ErrorHappend = ex.Message   ' properties of display messag error
            Return (False)
        End Try


    End Function
    Public Function DeleteTable(ByVal TabelName As String) As Boolean
        Try
            Dim query As String = "DROP TABLE " & TabelName


            SQLString = query  ' the sql string we userd in database
            Using conn As New OracleConnection(connectionString)
                Using comm As New OracleCommand()
                    With comm
                        .Connection = conn
                        .CommandType = CommandType.Text
                        .CommandText = query

                    End With
                    Try

                        conn.Open()
                        comm.ExecuteNonQuery()
                        conn.Close()
                        Return (True)
                    Catch ex1 As Exception '   inserted in db
                        ErrorHappend = ex1.Message   ' properties of display messag error
                        Return (False)
                    End Try

                End Using
            End Using

        Catch ex As Exception ' if he error enter  
            ErrorHappend = ex.Message   ' properties of display messag error
            Return (False)
        End Try


    End Function
    'Public Function RenameTable(ByVal OldTabelName As String, ByVal NewTabelName As String) As Boolean
    '    Try
    '        Dim query As String = "ALTER TABLE " & OldTabelName & " RENAME " & NewTabelName


    '        SQLString = query  ' the sql string we userd in database
    '        Using conn As New SqlConnection(connectionString)
    '            Using comm As New SqlCommand()
    '                With comm
    '                    .Connection = conn
    '                    .CommandType = CommandType.Text
    '                    .CommandText = query

    '                End With
    '                Try

    '                    conn.Open()
    '                    comm.ExecuteNonQuery()
    '                    conn.Close()
    '                    Return (True)
    '                Catch ex1 As Exception '   inserted in db
    '                    ErrorHappend = ex1.Message   ' properties of display messag error
    '                    Return (False)
    '                End Try

    '            End Using
    '        End Using

    '    Catch ex As Exception ' if he error enter  
    '        ErrorHappend = ex.Message   ' properties of display messag error
    '        Return (False)
    '    End Try


    'End Function

    Public Function insertColumn(ByVal TabelName As String, ByVal coloumns() As ColTabelParam) As Boolean
        Try
            Dim query As String = "ALTER TABLE " & TabelName & " ADD "

            For i = 0 To coloumns.Length - 1




                If coloumns(i).ColoumnsType = ColoumnType.Int16 Then

                    query = query & coloumns(i).ColoumnsName & " Int16"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.Int32 Then

                    query = query & coloumns(i).ColoumnsName & " Int32"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.UInt16 Then

                    query = query & coloumns(i).ColoumnsName & " UInt16"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.UInt32 Then

                    query = query & coloumns(i).ColoumnsName & " UInt32"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.Float Then

                    query = query & coloumns(i).ColoumnsName & " Float"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.Double1 Then

                    query = query & coloumns(i).ColoumnsName & " Double"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.BFile Then

                    query = query & coloumns(i).ColoumnsName & " BFile"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.Char1 Then

                    query = query & coloumns(i).ColoumnsName & " Char"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.DateTime Then

                    query = query & coloumns(i).ColoumnsName & " DateTime"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.varchar Then '*************

                    query = query & coloumns(i).ColoumnsName & " VarChar" & "(" & coloumns(i).ColumnsSize & ")"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.NChar Then '*************

                    query = query & coloumns(i).ColoumnsName & " NChar" & "(" & coloumns(i).ColumnsSize & ")"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.NVarChar Then '*************

                    query = query & coloumns(i).ColoumnsName & " NVarChar" & "(" & coloumns(i).ColumnsSize & ")"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.Number Then '*************

                    query = query & coloumns(i).ColoumnsName & " Number"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.Byte1 Then '*************

                    query = query & coloumns(i).ColoumnsName & " Byte"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.SByte1 Then '*************

                    query = query & coloumns(i).ColoumnsName & " SByte"
                End If
                '' addons 
                If coloumns(i).NotNull = True Then
                    query = query & " NOT NULL"
                End If

                If coloumns(i).PrimaryKey = True Then '
                    query = query & " PRIMARY KEY"
                End If
                If coloumns(i).AutoIncrement = True Then   ' for acees  (AUTOINCREMENT)  for oracle ( SEQUENCE seq_person MINVALUE 1 START WITH 1 INCREMENT BY 1)
                    query = query & " SEQUENCE seq_person MINVALUE 1 START WITH 1 INCREMENT BY 1"
                End If
                query = query & ","

            Next
            query = Mid(query, 1, Len(query) - 1) ' & ") "

            SQLString = query  ' the sql string we userd in database
            Using conn As New OracleConnection(connectionString)
                Using comm As New OracleCommand()
                    With comm
                        .Connection = conn
                        .CommandType = CommandType.Text
                        .CommandText = query

                    End With
                    Try

                        conn.Open()
                        comm.ExecuteNonQuery()
                        conn.Close()
                        Return (True)
                    Catch ex1 As Exception '   inserted in db
                        ErrorHappend = ex1.Message   ' properties of display messag error
                        Return (False)
                    End Try

                End Using
            End Using

        Catch ex As Exception ' if he error enter  
            ErrorHappend = ex.Message   ' properties of display messag error
            Return (False)
        End Try


    End Function

    Public Function DeleteColumn(ByVal TabelName As String, ByVal columns As String) As Boolean
        Try
            Dim query As String = "ALTER TABLE " & TabelName & " DROP COLUMN "
            ' Dim columnsSplit() As String = Split(columns, ",")
            ' For i = 0 To columnsSplit.Length - 1

            query = query & columns




            '     query = query & ","

            '   Next
            ' query = Mid(query, 1, Len(query) - 1) ' & ") "

            SQLString = query  ' the sql string we userd in database
            Using conn As New OracleConnection(connectionString)
                Using comm As New OracleCommand()
                    With comm
                        .Connection = conn
                        .CommandType = CommandType.Text
                        .CommandText = query

                    End With
                    Try

                        conn.Open()
                        comm.ExecuteNonQuery()
                        conn.Close()
                        Return (True)
                    Catch ex1 As Exception '   inserted in db
                        ErrorHappend = ex1.Message   ' properties of display messag error
                        Return (False)
                    End Try

                End Using
            End Using

        Catch ex As Exception ' if he error enter  
            ErrorHappend = ex.Message   ' properties of display messag error
            Return (False)
        End Try


    End Function

    Public Function ModifyColumn(ByVal TabelName As String, ByVal coloumns As ColTabelParam) As Boolean
        Try
            Dim query As String = "ALTER TABLE " & TabelName & " MODIFY "  ' for acces and SQL orcale use MODIFY



            If coloumns.ColoumnsType = ColoumnType.Int16 Then

                query = query & coloumns.ColoumnsName & " Int16"
            ElseIf coloumns.ColoumnsType = ColoumnType.Int32 Then

                query = query & coloumns.ColoumnsName & " Int32"
            ElseIf coloumns.ColoumnsType = ColoumnType.UInt16 Then

                query = query & coloumns.ColoumnsName & " UInt16"
            ElseIf coloumns.ColoumnsType = ColoumnType.UInt32 Then

                query = query & coloumns.ColoumnsName & " UInt32"
            ElseIf coloumns.ColoumnsType = ColoumnType.Float Then

                query = query & coloumns.ColoumnsName & " Float"
            ElseIf coloumns.ColoumnsType = ColoumnType.Double1 Then

                query = query & coloumns.ColoumnsName & " Double"
            ElseIf coloumns.ColoumnsType = ColoumnType.BFile Then

                query = query & coloumns.ColoumnsName & " BFile"
            ElseIf coloumns.ColoumnsType = ColoumnType.Char1 Then

                query = query & coloumns.ColoumnsName & " Char"

            ElseIf coloumns.ColoumnsType = ColoumnType.DateTime Then

                query = query & coloumns.ColoumnsName & " DateTime"

            ElseIf coloumns.ColoumnsType = ColoumnType.varchar Then '*************

                query = query & coloumns.ColoumnsName & " VarChar" & "(" & coloumns.ColumnsSize & ")"
            ElseIf coloumns.ColoumnsType = ColoumnType.NChar Then '*************

                query = query & coloumns.ColoumnsName & " NChar" & "(" & coloumns.ColumnsSize & ")"
            ElseIf coloumns.ColoumnsType = ColoumnType.NVarChar Then '*************

                query = query & coloumns.ColoumnsName & " NVarChar" & "(" & coloumns.ColumnsSize & ")"
            ElseIf coloumns.ColoumnsType = ColoumnType.Number Then '*************

                query = query & coloumns.ColoumnsName & " Number"
            ElseIf coloumns.ColoumnsType = ColoumnType.Byte1 Then '*************

                query = query & coloumns.ColoumnsName & " Byte"
            ElseIf coloumns.ColoumnsType = ColoumnType.SByte1 Then '*************

                query = query & coloumns.ColoumnsName & " SByte"
            End If
            '' addons 
            If coloumns.NotNull = True Then
                query = query & " NOT NULL"
            End If

            If coloumns.PrimaryKey = True Then '
                query = query & " PRIMARY KEY"
            End If
            If coloumns.AutoIncrement = True Then   ' for acees  (AUTOINCREMENT)  for oracle ( SEQUENCE seq_person MINVALUE 1 START WITH 1 INCREMENT BY 1)
                query = query & " SEQUENCE seq_person MINVALUE 1 START WITH 1 INCREMENT BY 1"
            End If

            ' query = query & ","   ' we here not use (,)


            '  query = Mid(query, 1, Len(query) - 1)  '& ") "

            SQLString = query  ' the sql string we userd in database
            Using conn As New OracleConnection(connectionString)
                Using comm As New OracleCommand()
                    With comm
                        .Connection = conn
                        .CommandType = CommandType.Text
                        .CommandText = query

                    End With
                    Try

                        conn.Open()
                        comm.ExecuteNonQuery()
                        conn.Close()
                        Return (True)
                    Catch ex1 As Exception '   inserted in db
                        ErrorHappend = ex1.Message   ' properties of display messag error
                        Return (False)
                    End Try

                End Using
            End Using

        Catch ex As Exception ' if he error enter  
            ErrorHappend = ex.Message   ' properties of display messag error
            Return (False)
        End Try


    End Function

End Class
