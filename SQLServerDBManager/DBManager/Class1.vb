'Imports System.Data.Common
Imports System.Data
Imports System.Data.SqlClient

Public Enum ColoumnType As Integer
    Text = 1
    Int = 2
    Float = 3
    Char1 = 4
    image = 5
    Date1 = 6
    DateTime = 7
    Bit = 8
    varchar50 = 9
    NChar = 10
    NText = 11
    NVarChar = 12
    Decimal1 = 13
    Time = 14
    money = 15
    uniqueidenifer = 16
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

Public Class cellProcedureParam

    Public Property ParameterName As String
    Public Property ColoumnType As SqlDbType
    Public Property Length As Integer
    Public Sub New(ByVal ParameterName As String, ByVal ColoumnType As SqlDbType, Optional ByVal Length As Integer = 0)

        Me.ParameterName = ParameterName

        Me.ColoumnType = ColoumnType
        Me.Length = Length



    End Sub



End Class
Public Class ColProcedureParam

    Public Property ParameterName As String
    Public Property ParameterValue As String
    Public Sub New(ByVal ParameterName As String, ByVal ParameterValue As String)
        Me.ParameterName = ParameterName


        Me.ParameterValue = ParameterValue

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
            Using conn As New SqlConnection(connectionString)
                Using comm As New SqlCommand()
                    With comm
                        .Connection = conn
                        .CommandType = CommandType.Text
                        .CommandText = query
                        For i = 0 To coloumns.Length - 1

                            If coloumns(i).ColoumnsType = ColoumnType.Int Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.Int)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Integer)


                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Text Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.Text)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, String)


                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Float Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.Float)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Double)
                            ElseIf coloumns(i).ColoumnsType = ColoumnType.image Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.Image)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Byte())

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Char1 Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.Char)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, String)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Date1 Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.Date)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Date)
                            ElseIf coloumns(i).ColoumnsType = ColoumnType.DateTime Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.DateTime)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = Convert.ToDateTime(coloumns(i).ColoumnsData) 'CType(coloumns(i).ColoumnsData, String)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Bit Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.Bit)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Boolean)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.varchar50 Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.VarChar)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, String)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.NChar Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.NChar)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, String)
                            ElseIf coloumns(i).ColoumnsType = ColoumnType.NText Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.NText)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, String)
                            ElseIf coloumns(i).ColoumnsType = ColoumnType.NVarChar Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.NVarChar)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, String)
                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Decimal1 Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.Decimal)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = Convert.ToDecimal(coloumns(i).ColoumnsData)
                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Time Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.Time)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = coloumns(i).ColoumnsData.ToString("hh:mm:ss tt") ' CType(coloumns(i).ColoumnsData, String)
                            ElseIf coloumns(i).ColoumnsType = ColoumnType.money Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.Money)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Double)
                            ElseIf coloumns(i).ColoumnsType = ColoumnType.uniqueidenifer Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.UniqueIdentifier)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = Guid.NewGuid()

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
            Using conn As New SqlConnection(connectionString)
                Using comm As New SqlCommand()
                    With comm
                        .Connection = conn
                        .CommandType = CommandType.Text
                        .CommandText = query
                        For i = 0 To coloumns.Length - 1

                            If coloumns(i).ColoumnsType = ColoumnType.Int Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.Int)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Integer)


                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Text Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.Text)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, String)


                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Float Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.Float)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Double)
                            ElseIf coloumns(i).ColoumnsType = ColoumnType.image Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.Image)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Byte())

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Char1 Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.Char)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, String)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Date1 Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.Date)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Date)
                            ElseIf coloumns(i).ColoumnsType = ColoumnType.DateTime Then
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.DateTime)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = Convert.ToDateTime(coloumns(i).ColoumnsData)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Bit Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.Bit)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Boolean)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.varchar50 Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.VarChar)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, String)

                            ElseIf coloumns(i).ColoumnsType = ColoumnType.NChar Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.NChar)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, String)
                            ElseIf coloumns(i).ColoumnsType = ColoumnType.NText Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.NText)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, String)
                            ElseIf coloumns(i).ColoumnsType = ColoumnType.NVarChar Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.NVarChar)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, String)
                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Decimal1 Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.Decimal)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = Convert.ToDecimal(coloumns(i).ColoumnsData)
                            ElseIf coloumns(i).ColoumnsType = ColoumnType.Time Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.Time)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, String)
                            ElseIf coloumns(i).ColoumnsType = ColoumnType.money Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.Money)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = CType(coloumns(i).ColoumnsData, Double)
                            ElseIf coloumns(i).ColoumnsType = ColoumnType.uniqueidenifer Then '*************
                                .Parameters.Add("@" & coloumns(i).ColoumnsName.ToString, SqlDbType.UniqueIdentifier)
                                .Parameters("@" & coloumns(i).ColoumnsName.ToString).Value = Guid.NewGuid()

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
            Using conn As New SqlConnection(connectionString)
                Using comm As New SqlCommand()
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

    Public Function HasInjection(Query As String) As Boolean
        Dim Inject As Boolean = False
        If Len(Query) > 0 Then
            If Query.IndexOf(";") > -1 Or Query.IndexOf("--") > -1 Then
                Inject = True
            End If
        End If

        Return (Inject)


    End Function



    Public Function SelectDataSet(ByVal TabelName As String, Optional ByVal coloumns As String = "*", Optional ByVal Condition As String = Nothing, Optional ByVal ColoumnsOrder As String = Nothing, Optional ByVal ColumnsGROUPBY As String = Nothing, Optional ByVal ColumnsHAVING As String = Nothing) As DataSet
        Dim SDataset As New DataSet
        SDataset.Clear()
        Try
            If HasInjection(coloumns) Or HasInjection(Condition) Then
                ErrorHappend = "data dont acepted"   ' properties of display messag error
                Dim TBUse As New DataTable("Tabel1")
                SDataset.Tables.Add(TBUse)
                Return (SDataset)
            End If
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
            Using conn As New SqlConnection(connectionString)
                Using comm As New SqlCommand()
                    With comm
                        .Connection = conn
                        .CommandType = CommandType.Text
                        .CommandText = query


                    End With

                    Try

                        conn.Open()
                        Dim DataAdapter1 As New SqlDataAdapter(comm)
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
            Using conn As New SqlConnection(connectionString)
                Using comm As New SqlCommand()
                    With comm
                        .Connection = conn
                        .CommandType = CommandType.Text
                        .CommandText = query


                    End With

                    Try

                        conn.Open()
                        Dim DataAdapter1 As New SqlDataAdapter(comm)
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

                    query = query & coloumns(i).ColoumnsName & " Int"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.Text Then

                    query = query & coloumns(i).ColoumnsName & " Text"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.Float Then

                    query = query & coloumns(i).ColoumnsName & " Float"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.image Then

                    query = query & coloumns(i).ColoumnsName & " Image"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.Char1 Then

                    query = query & coloumns(i).ColoumnsName & " Char"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.Date1 Then

                    query = query & coloumns(i).ColoumnsName & " Date"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.DateTime Then

                    query = query & coloumns(i).ColoumnsName & " DateTime"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.Bit Then '*************

                    query = query & coloumns(i).ColoumnsName & " Bit"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.varchar50 Then '*************

                    query = query & coloumns(i).ColoumnsName & " varchar" & "(" & coloumns(i).ColumnsSize & ")"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.NChar Then '*************

                    query = query & coloumns(i).ColoumnsName & " NChar" & "(" & coloumns(i).ColumnsSize & ")"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.NText Then '*************

                    query = query & coloumns(i).ColoumnsName & " NText" & "(" & coloumns(i).ColumnsSize & ")"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.NVarChar Then '*************

                    query = query & coloumns(i).ColoumnsName & " NVarChar" & "(" & coloumns(i).ColumnsSize & ")"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.Decimal1 Then '*************

                    query = query & coloumns(i).ColoumnsName & " Decimal"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.Time Then '*************

                    query = query & coloumns(i).ColoumnsName & " Time"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.uniqueidenifer Then '*************
                    query = query & coloumns(i).ColoumnsName & " uniqueidenifer"
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
            Using conn As New SqlConnection(connectionString)
                Using comm As New SqlCommand()
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
            Using conn As New SqlConnection(connectionString)
                Using comm As New SqlCommand()
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

                    query = query & coloumns(i).ColoumnsName & " Int"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.Text Then

                    query = query & coloumns(i).ColoumnsName & " Text"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.Float Then

                    query = query & coloumns(i).ColoumnsName & " Float"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.image Then

                    query = query & coloumns(i).ColoumnsName & " Image"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.Char1 Then

                    query = query & coloumns(i).ColoumnsName & " Char"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.Date1 Then

                    query = query & coloumns(i).ColoumnsName & " Date"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.DateTime Then

                    query = query & coloumns(i).ColoumnsName & " DateTime"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.Bit Then '*************

                    query = query & coloumns(i).ColoumnsName & " Bit"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.varchar50 Then '*************

                    query = query & coloumns(i).ColoumnsName & " varchar" & "(" & coloumns(i).ColumnsSize & ")"

                ElseIf coloumns(i).ColoumnsType = ColoumnType.NChar Then '*************

                    query = query & coloumns(i).ColoumnsName & " NChar" & "(" & coloumns(i).ColumnsSize & ")"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.NText Then '*************

                    query = query & coloumns(i).ColoumnsName & " NText" & "(" & coloumns(i).ColumnsSize & ")"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.NVarChar Then '*************

                    query = query & coloumns(i).ColoumnsName & " NVarChar" & "(" & coloumns(i).ColumnsSize & ")"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.Decimal1 Then '*************

                    query = query & coloumns(i).ColoumnsName & " Decimal"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.Time Then '*************

                    query = query & coloumns(i).ColoumnsName & " Time"
                ElseIf coloumns(i).ColoumnsType = ColoumnType.uniqueidenifer Then '*************
                    query = query & coloumns(i).ColoumnsName & " uniqueidenifer"
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
            Using conn As New SqlConnection(connectionString)
                Using comm As New SqlCommand()
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
            Using conn As New SqlConnection(connectionString)
                Using comm As New SqlCommand()
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

                query = query & coloumns.ColoumnsName & " Int"

            ElseIf coloumns.ColoumnsType = ColoumnType.Text Then

                query = query & coloumns.ColoumnsName & " Text"

            ElseIf coloumns.ColoumnsType = ColoumnType.Float Then

                query = query & coloumns.ColoumnsName & " Float"
            ElseIf coloumns.ColoumnsType = ColoumnType.image Then

                query = query & coloumns.ColoumnsName & " Image"
            ElseIf coloumns.ColoumnsType = ColoumnType.Char1 Then

                query = query & coloumns.ColoumnsName & " Char"
            ElseIf coloumns.ColoumnsType = ColoumnType.Date1 Then

                query = query & coloumns.ColoumnsName & " Date"
            ElseIf coloumns.ColoumnsType = ColoumnType.DateTime Then

                query = query & coloumns.ColoumnsName & " DateTime"

            ElseIf coloumns.ColoumnsType = ColoumnType.Bit Then '*************

                query = query & coloumns.ColoumnsName & " Bit"

            ElseIf coloumns.ColoumnsType = ColoumnType.varchar50 Then '*************

                query = query & coloumns.ColoumnsName & " varchar" & "(" & coloumns.ColumnsSize & ")"

            ElseIf coloumns.ColoumnsType = ColoumnType.NChar Then '*************

                query = query & coloumns.ColoumnsName & " NChar" & "(" & coloumns.ColumnsSize & ")"
            ElseIf coloumns.ColoumnsType = ColoumnType.NText Then '*************

                query = query & coloumns.ColoumnsName & " NText" & "(" & coloumns.ColumnsSize & ")"
            ElseIf coloumns.ColoumnsType = ColoumnType.NVarChar Then '*************

                query = query & coloumns.ColoumnsName & " NVarChar" & "(" & coloumns.ColumnsSize & ")"
            ElseIf coloumns.ColoumnsType = ColoumnType.Decimal1 Then '*************

                query = query & coloumns.ColoumnsName & " Decimal"
            ElseIf coloumns.ColoumnsType = ColoumnType.Time Then '*************


            ElseIf coloumns.ColoumnsType = ColoumnType.uniqueidenifer Then '*************
                query = query & coloumns.ColoumnsName & " uniqueidenifer"

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
            Using conn As New SqlConnection(connectionString)
                Using comm As New SqlCommand()
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

    Public Function SelectDataSetProcedureTable(ByVal ProcedureName As String, Optional ByVal ProcedureParam() As ColProcedureParam = Nothing) As DataSet
        Dim SDataset As New DataSet
        SDataset.Clear()
        Try

            Using conn As New SqlConnection(connectionString)
                Using comm As New SqlCommand()
                    With comm
                        .Connection = conn
                        .CommandType = CommandType.StoredProcedure
                        .CommandText = ProcedureName
                        If Not (ProcedureParam Is Nothing) Then
                            For i = 0 To ProcedureParam.Length - 1
                                .Parameters.Add(
                                   New SqlParameter("@" & ProcedureParam(i).ParameterName, ProcedureParam(i).ParameterValue))

                            Next
                        End If
                    End With

                    Try

                        conn.Open()
                        Dim DataAdapter1 As New SqlDataAdapter(comm)
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

    Public Function SelectDataSetProcedureCell(ByVal ProcedureName As String, Optional ByVal OutProcedureParam() As cellProcedureParam = Nothing, Optional ByVal InProcedureParam() As ColProcedureParam = Nothing) As SqlParameter()
        Dim SDataset As New DataSet
        SDataset.Clear()
        Try

            Using conn As New SqlConnection(connectionString)
                Using comm As New SqlCommand()
                    With comm
                        .Connection = conn
                        .CommandType = CommandType.StoredProcedure
                        .CommandText = ProcedureName
                        If Not (InProcedureParam Is Nothing) Then
                            For i = 0 To InProcedureParam.Length - 1
                                .Parameters.Add(
                                   New SqlParameter("@" & InProcedureParam(i).ParameterName, InProcedureParam(i).ParameterValue))

                            Next
                        End If


                    End With

                    Try

                        If Not (OutProcedureParam Is Nothing) Then
                            Dim output(OutProcedureParam.Length) As SqlParameter
                            For i = 0 To OutProcedureParam.Length - 1
                                output(i) = New SqlParameter("@" & OutProcedureParam(i).ParameterName, OutProcedureParam(i).ColoumnType)
                                output(i).Direction = ParameterDirection.Output
                                output(i).Size = OutProcedureParam(i).Length
                                comm.Parameters.Add(output(i))

                            Next

                            conn.Open()
                            comm.ExecuteNonQuery()

                            Return (output)
                            conn.Close()
                        Else
                            Return (Nothing)
                        End If
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

End Class