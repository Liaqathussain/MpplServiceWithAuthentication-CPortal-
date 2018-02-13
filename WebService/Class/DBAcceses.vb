Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.Common
Imports System.Data.OleDb
Imports System.Data.SqlClient

Public Class DBAcceses
    Dim chkQry As Integer
    Private conStr As String = String.Empty
    Private Shared _sServerName As String
    Private Shared _iPort As Integer
    Private Shared _isMADS As Boolean
    Private _eDBType As DBEngineType
    Private m_sEngineType
    Private objDT As DataTable
    Private objDS As DataSet
    Private objCon As DbConnection
    Private objCmd As DbCommand
    Private objAdp As DbDataAdapter
    Private objRD As DbDataReader
    Private _connectString As String

    
    Public Sub New(ByVal sConnectionString As String, ByVal DBEngine As DBEngineType)
        DBType = DBEngine
        objCon = GetConnection(sConnectionString, DBEngine)
        conStr = sConnectionString
    End Sub

    Public ReadOnly Property Connection() As DbConnection
        Get
            Return objCon
        End Get
    End Property

    Public Shared Property ServerName() As String
        Get
            Return _sServerName
        End Get
        Set(ByVal value As String)
            _sServerName = value
        End Set
    End Property

    Public Shared Property Port() As Integer
        Get
            Return _iPort
        End Get
        Set(ByVal value As Integer)
            _iPort = value
        End Set
    End Property

    Public Shared Property MADSON() As Boolean
        Get
            Return _isMADS
        End Get
        Set(ByVal value As Boolean)
            _isMADS = value
        End Set
    End Property

    Public Property ConnectString() As String
        Get
            Return conStr
        End Get
        Set(ByVal value As String)
            conStr = value
        End Set
    End Property

    Public Property DBType() As DBEngineType
        Get
            Return _eDBType
        End Get
        Set(ByVal value As DBEngineType)
            _eDBType = value
        End Set
    End Property

    Public Function GetConnection(ByVal sConnectionString As String, ByVal DBEngine As DBEngineType) As DbConnection
        Try
            Select Case DBEngine
                Case DBEngineType.OLEDB
                    objCon = New OleDb.OleDbConnection(sConnectionString)
                Case DBEngineType.ODBC
                    objCon = New Odbc.OdbcConnection(sConnectionString)
                Case DBEngineType.SQL
                    objCon = New SqlClient.SqlConnection(sConnectionString)
            End Select

            ' objCon.Open()
        Catch ex As Exception
            'objLog.PersistException("DataAccess", "GetConnection()", ex)
            ' Logger.logException("DataAccess", "GetConnection", "CS:" & sConnectionString & ", Message:" & ex.Message)
        End Try
        Return objCon
    End Function

    Public Function getDataTable(ByVal qry As String) As DataTable
        Try

            Try
                'Logger.logActivity("DataAccess", "GetDT", "Query : " & qry)

                objDT = New DataTable
                objCon.Open()

                objCmd = getCommandObject(qry)
                objCmd.Connection = objCon
                ' Logger.logActivity("DataAccess", "getDataTable", "Connection : " & objCon.ConnectionString)
                objAdp = getAdapterObject()
                objAdp.SelectCommand = objCmd

                objAdp.Fill(objDT)
                ' Logger.logActivity("DataAccess", "GetDT", "Count : " & objDT.Rows.Count)
            Catch ex As Exception
                ' Logger.logException("DataAccess", "GetDT", "CS:" & objCon.ConnectionString & ", Message: " & ex.Message)
            Finally
                objCon.Close()
            End Try
            Return objDT
        Catch ex As Exception
            ' Logger.logException("DataAccess", "GetDT", ex.Message)
        Finally

        End Try
        Return objDT
    End Function

    Public Function getDataSet(ByVal qry As String) As DataSet
        Try
            objCmd = getCommandObject(qry)
            objDS = New DataSet
            objCon.Open()
            objCmd.Connection = objCon
            objAdp = getAdapterObject()
            objAdp.SelectCommand = objCmd
            objAdp.Fill(objDS)
        Catch ex As Exception
            ' Logger.logException("DataAccess", "GetDataSet", "CS:" & objCon.ConnectionString & ", Message:" & ex.Message)
        Finally
            objCon.Close()
        End Try
        Return objDS
    End Function

    Public Function getDataReader(ByVal qry As String) As DbDataReader
        Try
            objCon.Open()
            objCmd = getCommandObject(qry)
            objCmd.Connection = objCon
            objRD = objCmd.ExecuteReader()
        Catch ex As Exception
            ' Logger.logException("DataAccess", "GetDataReader", "CS:" & objCon.ConnectionString & ", Message:" & ex.Message)
        Finally
            objCon.Close()
        End Try
        Return objRD
    End Function

    Public Sub close()
        objCon.Close()
    End Sub

    Public Function executeProc(ByVal ProcName As String) As Boolean


        Try
            objCon.Open()
            objCmd = New SqlClient.SqlCommand
            objCmd.Connection = objCon
            objCmd.CommandType = CommandType.StoredProcedure
            objCmd.CommandText = ProcName
            chkQry = objCmd.ExecuteNonQuery()
        Catch ex As Exception
            chkQry = 0
        Finally
            objCon.Close()
        End Try
        If chkQry = 0 Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Function executeProcess(ByVal ProcName As String, ByVal Par(,) As String) As Boolean


        Try
            objCon.Open()

            Dim cmd As New SqlClient.SqlCommand(ProcName, objCon)

            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandTimeout = 45


            For i As Int16 = 0 To UBound(Par)
                cmd.Parameters.AddWithValue(Par(i, 0), Par(i, 1))
            Next


            chkQry = cmd.ExecuteNonQuery()
        Catch ex As Exception
            chkQry = 0
        Finally
            objCon.Close()
        End Try
        If chkQry = 0 Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Function InsertProc_GetLastIncidentID(ByVal ProcName As String, ByVal Par(,) As String)
        Try
            objCon.Open()

            Dim cmd As New SqlClient.SqlCommand(ProcName, objCon)

            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandTimeout = 45


            For i As Int16 = 0 To UBound(Par)
                If i = 0 Then
                    cmd.Parameters.Add(New SqlParameter("@IncidentID", SqlDbType.BigInt))
                    cmd.Parameters("@IncidentID").Direction = ParameterDirection.Output
                    cmd.Parameters.AddWithValue(Par(i, 0), Par(i, 1))

                Else
                    cmd.Parameters.AddWithValue(Par(i, 0), Par(i, 1))

                End If

            Next


            cmd.ExecuteNonQuery()

            Dim lastComplainID = cmd.Parameters("@IncidentID").Value
            Return lastComplainID

        Catch ex As Exception
            chkQry = 0
        Finally
            objCon.Close()
        End Try

    End Function

    Public Function InsertProc_GeComplainID(ByVal ProcName As String, ByVal Par(,) As String)
        Try
            objCon.Open()

            Dim cmd As New SqlClient.SqlCommand(ProcName, objCon)

            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandTimeout = 45


            For i As Int16 = 0 To UBound(Par)
                If i = 0 Then
                    cmd.Parameters.Add(New SqlParameter("@ComplainID", SqlDbType.BigInt))
                    cmd.Parameters("@ComplainID").Direction = ParameterDirection.Output
                    cmd.Parameters.AddWithValue(Par(i, 0), Par(i, 1))

                Else
                    cmd.Parameters.AddWithValue(Par(i, 0), Par(i, 1))

                End If

            Next


            cmd.ExecuteNonQuery()
            Dim InsertedID = cmd.Parameters("@ComplainID").Value.ToString()
            Return InsertedID

        Catch ex As Exception
            chkQry = 0
        Finally
            objCon.Close()
        End Try

    End Function

    Public Function SP_Datatable(ByVal SP_Name As String, ByVal Par(,) As String) As DataTable

        Try

            objCon.Open()
            Dim cmd As New SqlClient.SqlCommand(SP_Name, objCon)

            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandTimeout = 45

            For i As Int16 = 0 To UBound(Par)
                cmd.Parameters.AddWithValue(Par(i, 0), Par(i, 1))
            Next

            objDT = New DataTable
            objAdp = getAdapterObject()
            objAdp.SelectCommand = cmd
            objAdp.Fill(objDT)

            Return objDT

        Catch ex As Exception

            objDT = Nothing
            Return objDT

        Finally
            objCon.Close()
        End Try

    End Function

    Public Function SP_Datatable(ByVal SP_Name As String) As DataTable

        Try

            objCon.Open()
            Dim cmd As New SqlClient.SqlCommand(SP_Name, objCon)

            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandTimeout = 45



            objDT = New DataTable
            objAdp = getAdapterObject()
            objAdp.SelectCommand = cmd
            objAdp.Fill(objDT)

            Return objDT

        Catch ex As Exception

            objDT = Nothing
            Return objDT

        Finally
            objCon.Close()
        End Try

    End Function

    Public Function executeQry(ByVal qry As String) As Boolean
        Try

            Try
                objCon.Open()
                objCmd = getCommandObject(qry)
                objCmd.Connection = objCon
                chkQry = objCmd.ExecuteNonQuery()
            Catch ex As Exception
                chkQry = 0
            Finally
                objCon.Close()
            End Try


        Catch ex As Exception
            'objLog.PersistException("DataAccess", "executeQry()", ex)
        Finally

        End Try
        If chkQry = 0 Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Function executeProc(ByVal ProcName As String, ByVal Par(,) As String) As Boolean


        Try
            objCon.Open()

            Dim cmd As New SqlClient.SqlCommand(ProcName, objCon)

            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandTimeout = 45


            For i As Int16 = 0 To UBound(Par)
                cmd.Parameters.AddWithValue(Par(i, 0), Par(i, 1))
            Next


            chkQry = cmd.ExecuteNonQuery()
        Catch ex As Exception
            chkQry = 0
        Finally
            objCon.Close()
        End Try
        If chkQry = 0 Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Function getCommandObject(ByVal qry As String) As DbCommand
        Dim objTmpCmd As DbCommand = Nothing
        Select Case DBType
            Case DBEngineType.OLEDB
                objTmpCmd = New OleDb.OleDbCommand(qry)
            Case DBEngineType.ODBC
                objTmpCmd = New Odbc.OdbcCommand(qry)
            Case DBEngineType.SQL
                objTmpCmd = New SqlClient.SqlCommand(qry)
        End Select

        Return objTmpCmd
    End Function

    Public Function getAdapterObject() As DbDataAdapter
        Dim objTmpAdp As DbDataAdapter = Nothing
        Select Case DBType
            Case DBEngineType.OLEDB
                objTmpAdp = New OleDb.OleDbDataAdapter
            Case DBEngineType.SQL
                objTmpAdp = New SqlClient.SqlDataAdapter
            Case DBEngineType.ODBC
                objTmpAdp = New Odbc.OdbcDataAdapter
        End Select
        Return objTmpAdp
    End Function

    Public Property EngineType() As String

        Get
            Return m_sEngineType
        End Get

        Set(ByVal value As String)
            m_sEngineType = value
        End Set
    End Property

End Class

Public Enum DBEngineType
    ODBC
    OLEDB
    SQL
End Enum


