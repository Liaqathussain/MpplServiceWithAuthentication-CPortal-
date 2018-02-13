Public Class AuthUser
    Inherits System.Web.Services.Protocols.SoapHeader
    Private m_UserName As String
    Private m_Password As String
    Dim objMW As New DBAcceses(My.Settings.conMidd, DBEngineType.SQL)
    Dim objcls As New EmailTemplate

    Public Property UserName() As String

        Get
            Return m_UserName
        End Get

        Set(value As String)
            m_UserName = Value
        End Set

    End Property

    Public Property Password() As String

        Get
            Return m_Password
        End Get

        Set(value As String)
            m_Password = value
        End Set

    End Property

    Public Function IsValid(ByVal Method As String) As Boolean
        Try
            'UserName = "Portal_01"
            'Password = "Mppl@2017"


            'UserName = "APPDB"
            'Password = "Multi@2017"

            Dim dt As New DataTable
            Dim dt_2 As New DataTable
            Dim para As String(,) = {{"@Username", UserName}, _
                                    {"@Password", Password}}


            dt = objMW.SP_Datatable("sp_GetServiceCredentials", para)
            dt.TableName = "sp_GetServiceCredentials"

            If dt.Rows.Count <> 0 Then
                dt_2 = objcls.GetMethodViaUser(Convert.ToInt32(dt.Rows(0)("UserID")))
                If dt_2.Rows.Count > 0 Then
                    For i = 0 To dt_2.Rows.Count - 1
                        If Method = Convert.ToString(dt_2.Rows(i)("MethodName")) Then
                            objcls.WriteLoggedDB("MpplService", Method + " Proceed", Convert.ToString(dt.Rows(0)("Domain")))
                            Return True
                        End If
                    Next
                End If



            Else
                objcls.WriteLoggedDB("MpplService", Method + " Failed", "Unknown")
                Return False
            End If
        Catch ex As Exception
            Throw ex
        End Try


    End Function

End Class
