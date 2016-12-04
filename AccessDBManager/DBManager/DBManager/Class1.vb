Imports System.Data
Imports System.Data.OleDb

Public Enum ColoumnType As Integer

    Int = 1
    Double1 = 2
    LongVarBinary = 3
    Char1 = 4
    Date1 = 5
    Boolean1 = 6
    Decimal1 = 7
    Time = 8
    LongVarChar = 9
    varchar = 10
    Filetime = 11
    Binary = 12
    BigInt = 13
    LongVarWChar = 14
    VarBinary = 15
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
    Public Property SQLString As String   ' the sql string we userd in database
    Public Property ErrorHappend As String  ' erro happned in data base 
    Public Sub New(ByVal connectionStringDB As String)
        connectionString = connectionStringDB
    End Sub

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
            Using conn As New OleDbConnection(connectionString)
                Using comm As New OleDbCommand()
                    With comm
                        .Connection = conn
                        .CommandType = CommandType.Text
                        .CommandText = query
                        For i = 0 To coloumns.Length - 1

                            If coloumns(i).ColoumnsType = ColoumnType.Int Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OleDbType.Integer)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Integer)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Double1 Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OleDbType.Double)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Double)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.LongVarBinary Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OleDbType.LongVarBinary)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Byte())

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Char1 Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OleDbType.Char)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Char)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Date1 Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OleDbType.Date)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Date)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Boolean1 Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OleDbType.Boolean)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Boolean)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.varchar Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OleDbType.VarChar)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, String)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.LongVarChar Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OleDbType.LongVarChar)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, String)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Decimal1 Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OleDbType.Decimal)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = Convert.ToDecimal(coloumns(i).ColoumnsData)
                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Time Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OleDbType.DBTime)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, String)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Filetime Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OleDbType.Filetime)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = Convert.ToDateTime(coloumns(i).ColoumnsData)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Binary Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OleDbType.Binary)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Byte())

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.BigInt Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OleDbType.BigInt)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = Convert.ToInt64(coloumns(i).ColoumnsData)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.LongVarWChar Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OleDbType.LongVarWChar)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, String)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.VarBinary Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OleDbType.VarBinary)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Byte())

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
            Using conn As New OleDbConnection(connectionString)
                Using comm As New OleDbCommand()
                    With comm
                        .Connection = conn
                        .CommandType = CommandType.Text
                        .CommandText = query
                        For i = 0 To coloumns.Length - 1

                            If coloumns(i).ColoumnsType = ColoumnType.Int Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OleDbType.Integer)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Integer)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Double1 Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OleDbType.Double)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Double)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.LongVarBinary Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OleDbType.LongVarBinary)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Byte())

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Char1 Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OleDbType.Char)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Char)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Date1 Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OleDbType.Date)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Date)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Boolean1 Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OleDbType.Boolean)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Boolean)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.varchar Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OleDbType.VarChar)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, String)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.LongVarChar Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OleDbType.LongVarChar)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, String)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Decimal1 Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OleDbType.Decimal)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = Convert.ToDecimal(coloumns(i).ColoumnsData)
                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Time Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OleDbType.DBTime)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, String)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Filetime Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OleDbType.Filetime)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = Convert.ToDateTime(coloumns(i).ColoumnsData)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Binary Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OleDbType.Binary)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Byte())

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.BigInt Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OleDbType.BigInt)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = Convert.ToInt64(coloumns(i).ColoumnsData)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.LongVarWChar Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OleDbType.LongVarWChar)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, String)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.VarBinary Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, OleDbType.VarBinary)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Byte())

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
            Using conn As New OleDbConnection(connectionString)
                Using comm As New OleDbCommand()
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
            Using conn As New OleDbConnection(connectionString)
                Using comm As New OleDbCommand()
                    With comm
                        .Connection = conn
                        .CommandType = CommandType.Text
                        .CommandText = query


                    End With

                    Try

                        conn.Open()
                        Dim DataAdapter1 As New OleDbDataAdapter(comm)
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

    Public Function SelectCell(ByVal TabelName As String, ByVal SearchColoumnName As String, ByVal ColoumnCondtion As String, ByVal ColoumnCondtionValue As String) As String
        Dim SDataset As New DataSet
        SDataset.Clear()
        Try
            Dim query As String = "SELECT " & Trim(SearchColoumnName) & " FROM " & Trim(TabelName) & " where " & Trim(ColoumnCondtion) & " LIKE " & Trim(ColoumnCondtionValue)
            SQLString = query  ' the sql string we userd in database
            Using conn As New OleDbConnection(connectionString)
                Using comm As New OleDbCommand()
                    With comm
                        .Connection = conn
                        .CommandType = CommandType.Text
                        .CommandText = query


                    End With

                    Try

                        conn.Open()
                        Dim DataAdapter1 As New OleDbDataAdapter(comm)
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



                If coloumns(i).ColoumnsType = ColoumnType.Int Then

                    query = query & coloumns(i).ColoumnsName & " Integer"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.Double1 Then

                    query = query & coloumns(i).ColoumnsName & " Double"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.LongVarBinary Then

                    query = query & coloumns(i).ColoumnsName & " LongVarBinary"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.Char1 Then

                    query = query & coloumns(i).ColoumnsName & " Char"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.Date1 Then

                    query = query & coloumns(i).ColoumnsName & " Date"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.Boolean1 Then '*************

                    query = query & coloumns(i).ColoumnsName & " Boolean"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.varchar Then '*************

                    query = query & coloumns(i).ColoumnsName & " VarChar" & "(" & coloumns(i).ColumnsSize & ")"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.LongVarChar Then '*************

                    query = query & coloumns(i).ColoumnsName & " LongVarChar"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.Decimal1 Then '*************

                    query = query & coloumns(i).ColoumnsName & " Decimal"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.Time Then '*************

                    query = query & coloumns(i).ColoumnsName & " DBTime"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.Filetime Then '*************

                    query = query & coloumns(i).ColoumnsName & " Filetime"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.Binary Then '*************

                    query = query & coloumns(i).ColoumnsName & " Binary"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.BigInt Then '*************

                    query = query & coloumns(i).ColoumnsName & " BigInt"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.LongVarWChar Then '*************

                    query = query & coloumns(i).ColoumnsName & " LongVarWChar"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.VarBinary Then '*************

                    query = query & coloumns(i).ColoumnsName & " VarBinary" & "(" & coloumns(i).ColumnsSize & ")"

                End If


                '' addons 
                If coloumns(i).NotNull = True Then
                    query = query & " NOT NULL"
                End If

                If coloumns(i).PrimaryKey = True Then '
                    query = query & " PRIMARY KEY"
                End If
                If coloumns(i).AutoIncrement = True Then   ' for acees  (AUTOINCREMENT)  for oracle ( SEQUENCE seq_person MINVALUE 1 START WITH 1 INCREMENT BY 1)
                    query = query & " IDENTITY"
                End If
                query = query & ","

            Next
            query = Mid(query, 1, Len(query) - 1) & ") "

            SQLString = query  ' the sql string we userd in database
            Using conn As New OleDbConnection(connectionString)
                Using comm As New OleDbCommand()
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
            Using conn As New OleDbConnection(connectionString)
                Using comm As New OleDbCommand()
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




                If coloumns(i).ColoumnsType = ColoumnType.Int Then

                    query = query & coloumns(i).ColoumnsName & " Integer"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.Double1 Then

                    query = query & coloumns(i).ColoumnsName & " Double"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.LongVarBinary Then

                    query = query & coloumns(i).ColoumnsName & " LongVarBinary"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.Char1 Then

                    query = query & coloumns(i).ColoumnsName & " Char"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.Date1 Then

                    query = query & coloumns(i).ColoumnsName & " Date"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.Boolean1 Then '*************

                    query = query & coloumns(i).ColoumnsName & " Boolean"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.varchar Then '*************

                    query = query & coloumns(i).ColoumnsName & " VarChar" & "(" & coloumns(i).ColumnsSize & ")"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.LongVarChar Then '*************

                    query = query & coloumns(i).ColoumnsName & " LongVarChar"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.Decimal1 Then '*************

                    query = query & coloumns(i).ColoumnsName & " Decimal"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.Time Then '*************

                    query = query & coloumns(i).ColoumnsName & " DBTime"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.Filetime Then '*************

                    query = query & coloumns(i).ColoumnsName & " Filetime"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.Binary Then '*************

                    query = query & coloumns(i).ColoumnsName & " Binary"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.BigInt Then '*************

                    query = query & coloumns(i).ColoumnsName & " BigInt"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.LongVarWChar Then '*************

                    query = query & coloumns(i).ColoumnsName & " LongVarWChar"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.VarBinary Then '*************

                    query = query & coloumns(i).ColoumnsName & " VarBinary" & "(" & coloumns(i).ColumnsSize & ")"

                End If

                '' addons 
                If coloumns(i).NotNull = True Then
                    query = query & " NOT NULL"
                End If

                If coloumns(i).PrimaryKey = True Then '
                    query = query & " PRIMARY KEY"
                End If
                If coloumns(i).AutoIncrement = True Then   ' for acees  (AUTOINCREMENT)  for oracle ( SEQUENCE seq_person MINVALUE 1 START WITH 1 INCREMENT BY 1)
                    query = query & " IDENTITY"
                End If
                query = query & ","

            Next
            query = Mid(query, 1, Len(query) - 1) ' & ") "

            SQLString = query  ' the sql string we userd in database
            Using conn As New OleDbConnection(connectionString)
                Using comm As New OleDbCommand()
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
            Using conn As New OleDbConnection(connectionString)
                Using comm As New OleDbCommand()
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
            Dim query As String = "ALTER TABLE " & TabelName & " ALTER COLUMN "  ' for acces and SQL orcale use MODIFY



            If coloumns.ColoumnsType = ColoumnType.Int Then

                query = query & coloumns.ColoumnsName & " Integer"
            ElseIf coloumns.ColoumnsType = ColoumnType.Double1 Then

                query = query & coloumns.ColoumnsName & " Double"

            ElseIf coloumns.ColoumnsType = ColoumnType.LongVarBinary Then

                query = query & coloumns.ColoumnsName & " LongVarBinary"

            ElseIf coloumns.ColoumnsType = ColoumnType.Char1 Then

                query = query & coloumns.ColoumnsName & " Char"

            ElseIf coloumns.ColoumnsType = ColoumnType.Date1 Then

                query = query & coloumns.ColoumnsName & " Date"

            ElseIf coloumns.ColoumnsType = ColoumnType.Boolean1 Then '*************

                query = query & coloumns.ColoumnsName & " Boolean"

            ElseIf coloumns.ColoumnsType = ColoumnType.varchar Then '*************

                query = query & coloumns.ColoumnsName & " VarChar" & "(" & coloumns.ColumnsSize & ")"

            ElseIf coloumns.ColoumnsType = ColoumnType.LongVarChar Then '*************

                query = query & coloumns.ColoumnsName & " LongVarChar"

            ElseIf coloumns.ColoumnsType = ColoumnType.Decimal1 Then '*************

                query = query & coloumns.ColoumnsName & " Decimal"
            ElseIf coloumns.ColoumnsType = ColoumnType.Time Then '*************

                query = query & coloumns.ColoumnsName & " DBTime"

            ElseIf coloumns.ColoumnsType = ColoumnType.Filetime Then '*************

                query = query & coloumns.ColoumnsName & " Filetime"

            ElseIf coloumns.ColoumnsType = ColoumnType.Binary Then '*************

                query = query & coloumns.ColoumnsName & " Binary"

            ElseIf coloumns.ColoumnsType = ColoumnType.BigInt Then '*************

                query = query & coloumns.ColoumnsName & " BigInt"

            ElseIf coloumns.ColoumnsType = ColoumnType.LongVarWChar Then '*************

                query = query & coloumns.ColoumnsName & " LongVarWChar"

            ElseIf coloumns.ColoumnsType = ColoumnType.VarBinary Then '*************

                query = query & coloumns.ColoumnsName & " VarBinary" & "(" & coloumns.ColumnsSize & ")"

            End If

            '' addons 
            If coloumns.NotNull = True Then
                query = query & " NOT NULL"
            End If

            If coloumns.PrimaryKey = True Then '
                query = query & " PRIMARY KEY"
            End If
            If coloumns.AutoIncrement = True Then   ' for acees  (AUTOINCREMENT)  for oracle ( SEQUENCE seq_person MINVALUE 1 START WITH 1 INCREMENT BY 1)
                query = query & " IDENTITY"
            End If
            ' query = query & ","   ' we here not use (,)


            '  query = Mid(query, 1, Len(query) - 1)  '& ") "

            SQLString = query  ' the sql string we userd in database
            Using conn As New OleDbConnection(connectionString)
                Using comm As New OleDbCommand()
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

