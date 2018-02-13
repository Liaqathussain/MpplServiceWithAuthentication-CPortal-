
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Web.Script.Services
Imports System.Web.Script.Serialization
Imports System.Net.Mail
Imports System.IO
Imports System.Net
Imports System
Imports System.Web
Imports System.Security.Principal
Imports System.Drawing
Imports System.IO.Compression


' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class WebService
    Inherits System.Web.Services.WebService
    Dim objMW As New DBAcceses(My.Settings.conMidd, DBEngineType.SQL)
    Dim objBSS As New DBAcceses(My.Settings.conBSS, DBEngineType.SQL)
    Dim objGP As New DBAcceses(My.Settings.conGP, DBEngineType.SQL)
    Dim objRM As New DBAcceses(My.Settings.conRM, DBEngineType.SQL)
    Dim ObjEmail As New EmailTemplate
    Dim Smtp_Server As SmtpClient
    Public User As New AuthUser


    '=======================================Call Center JSON Response==========================================

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function GetCustomerDetailsViaCaller(ByVal Phone As String) As String
        Dim dt As New DataTable
        Try

            If User IsNot Nothing Then
                If Not User.IsValid("GetCustomerDetailsViaCaller") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetCustomerDetailsViaCaller  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetCustomerDetailsViaCaller  Failed", "Unknown")
                Return "Please provide details"
            End If

            Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
            Dim rows As New List(Of Dictionary(Of String, Object))()
            Dim row As Dictionary(Of String, Object)

            Dim para As String(,) = {{"@Phone", Phone}}
            dt = objMW.SP_Datatable("sp_cc_GetMasterDetails", para)
            dt.TableName = "sp_cc_GetMasterDetails"
            For Each dr As DataRow In dt.Rows
                row = New Dictionary(Of String, Object)()
                For Each col As DataColumn In dt.Columns
                    row.Add(col.ColumnName, dr(col))
                Next
                rows.Add(row)
            Next

            Return serializer.Serialize(rows)

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    '=======================================Ldap Authentication Return True Or False===========================

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    Public Function Ldap_Authentication(ByVal Username As String, ByVal Password As String) As Boolean
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("Ldap_Authentication") Then
                    ObjEmail.WriteLoggedDB("MpplService", "Ldap_Authentication  Failed", "Unknown")
                    Return False
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "Ldap_Authentication  Failed", "Unknown")
                Return False
            End If


            If ObjEmail.ValidateActiveDirectoryLogin(Username, Password) Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    '=======================================Rain Maker GetNearestPlacemark in JSON=============================

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    Public Function RM_GetNearestPlacemark(ByVal cur_lat As String, ByVal cur_lng As String) As String
        Dim dt As New DataTable
        Try

            If User IsNot Nothing Then
                If Not User.IsValid("RM_GetNearestPlacemark") Then
                    ObjEmail.WriteLoggedDB("MpplService", "RM_GetNearestPlacemark  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "RM_GetNearestPlacemark  Failed", "Unknown")
                Return "Please provide details"
            End If

            Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
            Dim rows As New List(Of Dictionary(Of String, Object))()
            Dim row As Dictionary(Of String, Object)

            Dim para As String(,) = {{"@cur_lat", CDbl(cur_lat)}, _
                                     {"@cur_lng", CDbl(cur_lng)}}

            dt = objRM.SP_Datatable("sp_GetNearestPlacemark", para)
            dt.TableName = "sp_GetNearestPlacemark"

            For Each dr As DataRow In dt.Rows
                row = New Dictionary(Of String, Object)()
                For Each col As DataColumn In dt.Columns
                    row.Add(col.ColumnName, dr(col))
                Next
                rows.Add(row)
            Next

            Return serializer.Serialize(rows)

        Catch ex As Exception
            Throw ex
        End Try
    End Function


    '=======================================Methods for Tone C5 Response in XML================================

    <WebMethod> _
   <SoapHeader("User", Required:=True)> _
    Public Function TC5_GETCUSTOMERCIRCUIT(ByVal UserID As String) As DataTable
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GETCUSTOMERCIRCUIT") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GETCUSTOMERCIRCUIT  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GETCUSTOMERCIRCUIT  Failed", "Unknown")
                Return Nothing
            End If

            dt = objMW.SP_Datatable("SP_GETCUSTOMERCIRCUIT")
            dt.TableName = "SP_GETCUSTOMERCIRCUIT"
            Return dt

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
  <SoapHeader("User", Required:=True)> _
    Public Function TC5_GETCUSTOMERMAIN(ByVal UserID As String) As DataTable
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GETCUSTOMERMAIN") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GETCUSTOMERMAIN  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GETCUSTOMERMAIN  Failed", "Unknown")
                Return Nothing
            End If

            dt = objMW.SP_Datatable("SP_GETCUSTOMERMAIN")
            dt.TableName = "SP_GETCUSTOMERMAIN"
            Return dt

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    Public Function TC5_GETINVOICES(ByVal UserID As String) As DataTable
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GETINVOICES") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GETINVOICES  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GETINVOICES  Failed", "Unknown")
                Return Nothing
            End If

            dt = objMW.SP_Datatable("SP_GETINVOICES")
            dt.TableName = "SP_GETINVOICES"
            Return dt

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
   <SoapHeader("User", Required:=True)> _
    Public Function TC5_GETPAYMENTS(ByVal UserID As String) As DataTable
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GETPAYMENTS") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GETPAYMENTS  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GETPAYMENTS  Failed", "Unknown")
                Return Nothing
            End If

            dt = objMW.SP_Datatable("SP_GETPAYMENTS")
            dt.TableName = "SP_GETPAYMENTS"
            Return dt

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    '=======================================Methods for Web Portal Response in XML==============================

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    Public Function CustomerCredential(ByVal UserLogin As String, _
                                       ByVal Password As String) As DataTable
        Dim dt As New DataTable
        Try

            If User IsNot Nothing Then
                If Not User.IsValid("CustomerCredential") Then
                    ObjEmail.WriteLoggedDB("MpplService", "CustomerCredential  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "CustomerCredential  Failed", "Unknown")
                Return Nothing
            End If

            Dim para As String(,) = {{"@UserLogin", UserLogin}, {"@Password", Password}}
            dt = objMW.SP_Datatable("sp_GetUserCredentials", para)
            dt.TableName = "sp_GetUserCredentials"
            If dt.Rows.Count = 1 Then
                Return dt
            End If


        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    Public Function GetUserAccountAccess(ByVal UserID As String) As DataTable
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GetUserAccountAccess") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetUserAccountAccess  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetUserAccountAccess  Failed", "Unknown")
                Return Nothing
            End If

            Dim para As String(,) = {{"@UserID", CInt(UserID)}}
            dt = objMW.SP_Datatable("sp_GetUserAccountAccess", para)
            dt.TableName = "sp_GetUserAccountAccess"
            Return dt

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    Public Function GetRegion(ByVal BSSMasterCode As String) As DataTable
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GetRegion") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetRegion  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetRegion  Failed", "Unknown")
                Return Nothing
            End If


            Dim para As String(,) = {{"@BSSMasterCode", CInt(BSSMasterCode)}}
            dt = objMW.SP_Datatable("sp_GetCustomerRegion", para)
            dt.TableName = "sp_GetCustomerRegion"
            Return dt

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    Public Function GetCities(ByVal BSSMasterCode As String, _
                              ByVal Region As String) As DataTable
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GetCities") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetCities  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetCities  Failed", "Unknown")
                Return Nothing
            End If


            Dim para As String(,) = {{"@BSSMasterCode", CInt(BSSMasterCode)}, {"@Region", (Region)}}
            dt = objMW.SP_Datatable("sp_GetCustomerCities", para)
            dt.TableName = "sp_GetCustomerCities"
            Return dt

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    Public Function GetInfra(ByVal BSSMasterCode As String) As DataTable
        Dim dt As New DataTable
        Try

            If User IsNot Nothing Then
                If Not User.IsValid("GetInfra") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetInfra  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "CustomerCredential  Failed", "Unknown")
                Return Nothing
            End If

            Dim para As String(,) = {{"@BSSMasterCode", CInt(BSSMasterCode)}}
            dt = objMW.SP_Datatable("sp_GetCustomerInfra", para)
            dt.TableName = "sp_GetCustomerInfra"
            Return dt

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
     <SoapHeader("User", Required:=True)> _
    Public Function GetLOB(ByVal BSSMasterCode As String) As DataTable
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GetLOB") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetLOB  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetLOB  Failed", "Unknown")
                Return Nothing
            End If

            Dim para As String(,) = {{"@BSSMasterCode", CInt(BSSMasterCode)}}
            dt = objMW.SP_Datatable("sp_GetCustomerLOB", para)
            dt.TableName = "sp_GetCustomerLOB"
            Return dt

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
     <SoapHeader("User", Required:=True)> _
    Public Function GetServiceUnit(ByVal BSSMasterCode As String) As DataTable
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GetServiceUnit") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetServiceUnit  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetServiceUnit  Failed", "Unknown")
                Return Nothing
            End If

            Dim para As String(,) = {{"@BSSMasterCode", CInt(BSSMasterCode)}}
            dt = objMW.SP_Datatable("sp_GetCustomerServiceUnit", para)
            dt.TableName = "sp_GetCustomerServiceUnit"
            Return dt

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
     <SoapHeader("User", Required:=True)> _
    Public Function GetBandwidth(ByVal BSSMasterCode As String) As DataTable
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GetBandwidth") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetBandwidth  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetBandwidth  Failed", "Unknown")
                Return Nothing
            End If

            Dim para As String(,) = {{"@BSSMasterCode", CInt(BSSMasterCode)}}
            dt = objMW.SP_Datatable("sp_GetCustomerBandwidth", para)
            dt.TableName = "sp_GetCustomerBandwidth"
            Return dt

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
     <SoapHeader("User", Required:=True)> _
    Public Function GetCiruitsMainDetail(ByVal BSSMasterCode As String, _
                                       ByVal City As String, _
                                       ByVal Infra As String, _
                                       ByVal ServiceUnit As String, _
                                       ByVal Bandwidth As String, _
                                       ByVal Region As String, _
                                       ByVal IsFavourite As String, _
                                       ByVal UserID As String, _
                                       ByVal NMNStatus As String, _
                                       ByVal LOB As String) As DataTable
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GetCiruitsMainDetail") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetCiruitsMainDetail  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetCiruitsMainDetail  Failed", "Unknown")
                Return Nothing
            End If


            Dim para As String(,) = {{"@BSSMasterCode", (BSSMasterCode.Trim)}, _
                                      {"@City", City}, _
                                      {"@Infra", Infra}, _
                                      {"@ServiceUnit", ServiceUnit}, _
                                      {"@Bandwidth", Bandwidth}, _
                                      {"@Region", Region}, _
                                      {"@IsFavourite", CInt(IsFavourite)}, _
                                      {"@UserID", CInt(UserID)}, _
                                      {"@NMNStatus", NMNStatus}, _
                                      {"@LOB", LOB}}

            dt = objMW.SP_Datatable("sp_GetCiruitsMainDetailss", para)
            dt.TableName = "sp_GetCiruitsMainDetailss"
            Return dt

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
     <SoapHeader("User", Required:=True)> _
    Public Function GetCircuitCompleteDetail(ByVal SignupID As String) As DataTable
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GetCircuitCompleteDetail") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetCircuitCompleteDetail  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetCircuitCompleteDetail  Failed", "Unknown")
                Return Nothing
            End If

            Dim para As String(,) = {{"@SignupID", CInt(SignupID)}}
            dt = objMW.SP_Datatable("sp_GetCiruitsCompleteDetail", para)
            dt.TableName = "sp_GetCiruitsCompleteDetail"
            Return dt

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
     <SoapHeader("User", Required:=True)> _
    Public Function GetNewConnectionList(ByVal BSSMasterCode As String) As DataTable
        Dim dt As New DataTable
        Try

            If User IsNot Nothing Then
                If Not User.IsValid("GetNewConnectionList") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetNewConnectionList  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetNewConnectionList  Failed", "Unknown")
                Return Nothing
            End If

            Dim para As String(,) = {{"@BSSMasterCode", CInt(BSSMasterCode)}}
            dt = objMW.SP_Datatable("sp_GetNewConnectionList", para)
            dt.TableName = "sp_GetNewConnectionList"
            Return dt

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
     <SoapHeader("User", Required:=True)> _
    Public Function GetConnectionHistory(ByVal SignupID As String) As DataTable
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GetConnectionHistory") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetConnectionHistory  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetConnectionHistory  Failed", "Unknown")
                Return Nothing
            End If


            Dim para As String(,) = {{"@SignupID", CInt(SignupID)}}
            dt = objMW.SP_Datatable("sp_GetConnectionHistory", para)
            dt.TableName = "sp_GetConnectionHistory"
            Return dt

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
     <SoapHeader("User", Required:=True)> _
    Public Function GetCustomerMonthlyInvoices(ByVal BSSMasterCode As Integer, _
                                               ByVal GPMasterCode As String) As DataTable
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GetCustomerMonthlyInvoices") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetCustomerMonthlyInvoices  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetCustomerMonthlyInvoices  Failed", "Unknown")
                Return Nothing
            End If


            Dim para As String(,) = {{"@BSSMasterCode", CInt(BSSMasterCode)}, _
                                     {"@GPMasterCode", GPMasterCode}}
            dt = objMW.SP_Datatable("sp_GetCustomerMonthlyInvoices", para)
            dt.TableName = "sp_GetCustomerMonthlyInvoices"
            Return dt

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
     <SoapHeader("User", Required:=True)> _
    Public Function GetCustomerProdutSummary(ByVal BSSMasterCode As Integer, _
                                               ByVal GPMasterCode As String) As DataTable
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GetCustomerProdutSummary") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetCustomerProdutSummary  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetCustomerProdutSummary  Failed", "Unknown")
                Return Nothing
            End If


            Dim para As String(,) = {{"@BSSMasterCode", CInt(BSSMasterCode)}, _
                                     {"@GPMasterCode", GPMasterCode}}
            dt = objMW.SP_Datatable("sp_GetCustomerProdutSummary", para)
            dt.TableName = "sp_GetCustomerProdutSummary"
            Return dt

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    Public Function GetCustomerOutstanding(ByVal BSSMasterCode As Integer, _
                                               ByVal GPMasterCode As String) As DataTable
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GetCustomerOutstanding") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetCustomerOutstanding  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetCustomerOutstanding  Failed", "Unknown")
                Return Nothing
            End If



            Dim para As String(,) = {{"@BSSMasterCode", CInt(BSSMasterCode)}, _
                                     {"@GPMasterCode", GPMasterCode}}
            dt = objMW.SP_Datatable("sp_GetCustomerOutstanding", para)
            dt.TableName = "sp_GetCustomerOutstanding"
            Return dt

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    Public Function GetCustomerInovices(ByVal BSSMasterCode As String, _
                                           ByVal GPMasterCode As String, _
                                           ByVal Period As String, _
                                           ByVal InvoiceNo As String) As DataTable
        Dim dt As New DataTable
        Try

            If User IsNot Nothing Then
                If Not User.IsValid("GetCustomerInovices") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetCustomerInovices  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetCustomerInovices  Failed", "Unknown")
                Return Nothing
            End If


            Dim para As String(,) = {{"@BSSMasterCode", CInt(BSSMasterCode)}, _
                                     {"@GPMasterCode", GPMasterCode}, _
                                     {"@Period", Period}, _
                                     {"@InvoiceNo", InvoiceNo}}
            dt = objMW.SP_Datatable("sp_GetCustomerInovices", para)
            dt.TableName = "sp_GetCustomerInovices"
            Return dt

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
     <SoapHeader("User", Required:=True)> _
    Public Function SetFavCircuit(ByVal SignupID As String, _
                                         ByVal UserID As String) As Boolean
        Dim dt As New DataTable
        Try

            If User IsNot Nothing Then
                If Not User.IsValid("SetFavCircuit") Then
                    ObjEmail.WriteLoggedDB("MpplService", "SetFavCircuit  Failed", "Unknown")
                    Return False
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "SetFavCircuit  Failed", "Unknown")
                Return False
            End If

            Dim para As String(,) = {{"@SignupID", CInt(SignupID)}, _
                                     {"@UserID", CInt(UserID)}}
            If objMW.executeProc("sp_SetFavCircuit", para) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
     <SoapHeader("User", Required:=True)> _
    Public Function SetUnFavCircuit(ByVal SignupID As String, _
                                    ByVal UserID As String) As Boolean
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("SetUnFavCircuit") Then
                    ObjEmail.WriteLoggedDB("MpplService", "SetUnFavCircuit  Failed", "Unknown")
                    Return False
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "SetUnFavCircuit  Failed", "Unknown")
                Return False
            End If


            Dim para As String(,) = {{"@SignupID", CInt(SignupID)}, _
                                     {"@UserID", CInt(UserID)}}

            If objMW.executeProc("sp_SetUnFavCircuit", para) Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
     <SoapHeader("User", Required:=True)> _
    Public Function GetCustomerFav(ByVal UserID As String, _
                                   ByVal BSSMasterCode As String) As DataTable
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GetCustomerFav") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetCustomerFav  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetCustomerFav  Failed", "Unknown")
                Return Nothing
            End If

            Dim para As String(,) = {{"@UserID", CInt(UserID)}, _
                                     {"@BSSMasterCode", CInt(BSSMasterCode)}}

            dt = objMW.SP_Datatable("sp_GetCustomerFav", para)
            dt.TableName = "sp_GetCustomerFav"
            Return dt


        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
     <SoapHeader("User", Required:=True)> _
    Public Function GetInvoiceDetailForPDF(ByVal FROM_DATE As String, _
                                   ByVal TO_DATE As String, _
                                   ByVal SOPNUMBE As String) As DataTable
        Dim dt As New DataTable
        Try

            If User IsNot Nothing Then
                If Not User.IsValid("GetInvoiceDetailForPDF") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetInvoiceDetailForPDF  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetInvoiceDetailForPDF  Failed", "Unknown")
                Return Nothing
            End If
            Dim para As String(,) = {{"@B_POSTED", 1}, _
                                      {"@FROM_DATE", DateTime.Parse(FROM_DATE)}, _
                                      {"@TO_DATE", DateTime.Parse(TO_DATE)}, _
                                      {"@SOPNUMBE", SOPNUMBE}, _
                                      {"@FromInvPrint", DateTime.Parse(FROM_DATE)}, _
                                      {"@ToInvPrint", DateTime.Parse(TO_DATE)}}

            dt = objGP.SP_Datatable("sp_App_GetDetailsforPDF", para)
            dt.TableName = "sp_App_GetDetailsforPDF"
            Return dt
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
   <SoapHeader("User", Required:=True)> _
    Public Function GetContractInvoiceDetails(ByVal userid As String) As DataTable
        Dim dt As New DataTable
        Try

            If User IsNot Nothing Then
                If Not User.IsValid("GetContractInvoiceDetails") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetContractInvoiceDetails  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetContractInvoiceDetails  Failed", "Unknown")
                Return Nothing
            End If

            Dim para As String(,) = {{"@userid", userid}}

            dt = objGP.SP_Datatable("sp_app_ContractInvoiceDetails", para)
            dt.TableName = "sp_app_ContractInvoiceDetails"
            Return dt
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod()> _
    <SoapHeader("User", Required:=True)> _
    Public Function BSS_GetInitialStatments() As DataTable

        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("BSS_GetInitialStatments") Then
                    ObjEmail.WriteLoggedDB("MpplService", "BSS_GetInitialStatments  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "BSS_GetInitialStatments  Failed", "Unknown")
                Return Nothing
            End If

            dt = objBSS.SP_Datatable("sp_OTS_GetInitialStatement")
            dt.TableName = "sp_OTS_GetInitialStatement"
            Return dt

        Catch ex As Exception
            Throw ex
        Finally

        End Try

    End Function

    <WebMethod()> _
    <SoapHeader("User", Required:=True)> _
    Public Function BSS_GetComplains(ByVal BSSMasterCode As String, ByVal GPID As String, ByVal TicketNo As String, ByVal FromDate As String, ByVal ToDate As String) As DataTable
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("BSS_GetComplains") Then
                    ObjEmail.WriteLoggedDB("MpplService", "BSS_GetComplains  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "BSS_GetComplains  Failed", "Unknown")
                Return Nothing
            End If

            Dim para As String(,) = {{"@BSSMasterCode", CInt(BSSMasterCode)}, {"@GPID", GPID}, {"@TicketNo", TicketNo}, {"@FromDate", FromDate}, {"@ToDate", ToDate}}
            dt = objBSS.SP_Datatable("sp_app_GetComplains", para)
            dt.TableName = "sp_app_GetComplains"
            Return dt

        Catch ex As Exception
            Throw ex
        Finally

        End Try

    End Function

    <WebMethod()> _
    <SoapHeader("User", Required:=True)> _
    Public Function BSS_GetComplainWithWhere(ByVal Where As String) As DataTable
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("BSS_GetComplainWithWhere") Then
                    ObjEmail.WriteLoggedDB("MpplService", "BSS_GetComplainWithWhere  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "BSS_GetComplainWithWhere  Failed", "Unknown")
                Return Nothing
            End If

            Dim para As String(,) = {{"@Where", Where}}
            dt = objBSS.SP_Datatable("sp_app_GetComplainWithWhere", para)
            dt.TableName = "sp_app_GetComplainWithWhere"
            Return dt

        Catch ex As Exception
            Throw ex
        Finally

        End Try

    End Function

    <WebMethod()> _
    <SoapHeader("User", Required:=True)> _
    Public Function BSS_GetinterComplainsDetails(ByVal ComplaintID As Integer) As DataTable

        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("BSS_GetinterComplainsDetails") Then
                    ObjEmail.WriteLoggedDB("MpplService", "BSS_GetinterComplainsDetails  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "BSS_GetinterComplainsDetails  Failed", "Unknown")
                Return Nothing
            End If

            Dim para As String(,) = {{"@ComplaintID", ComplaintID}}
            dt = objBSS.SP_Datatable("sp_ICS_GetComplainDetails", para)
            dt.TableName = "sp_ICS_GetComplainDetails"
            Return dt

        Catch ex As Exception
            Throw ex
        Finally

        End Try

    End Function

    <WebMethod()> _
    <SoapHeader("User", Required:=True)> _
    Public Function BSS_InsertComplain(
                ByVal SignupID As Integer, _
                ByVal InitailStatementID As Integer, _
                ByVal PoCName As String, _
                ByVal PoCNumber As String, _
                ByVal Remarks As String) As Integer
        Try

            If User IsNot Nothing Then
                If Not User.IsValid("BSS_InsertComplain") Then
                    ObjEmail.WriteLoggedDB("MpplService", "BSS_InsertComplain  Failed", "Unknown")
                    Return 0
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "BSS_InsertComplain  Failed", "Unknown")
                Return 0
            End If

            Dim para As String(,) = {
             {"@SignupID", SignupID}, _
             {"@InitailStatementID", InitailStatementID}, _
             {"@PoCName", PoCName}, _
             {"@PoCNumber", PoCNumber}, _
             {"@Remarks", Remarks}, _
             {"@Flag", "WEB"}}

            Dim LastID As Integer = objBSS.InsertProc_GeComplainID("sp_app_InsertComplains", para)

            If LastID > 0 Then
                Return LastID
            Else
                Return 0

            End If


        Catch ex As Exception
            Throw ex
        End Try



    End Function

    <WebMethod()> _
    <SoapHeader("User", Required:=True)> _
    Public Function BSS_InsertComplainTroubleShooting(
                ByVal ComplainID As String, _
                ByVal LastMilePowerStatus As String, _
                ByVal FiberLEDStatus As String, _
                ByVal DeviceRebooted As String) As Boolean

        Try
            If User IsNot Nothing Then
                If Not User.IsValid("BSS_InsertComplainTroubleShooting") Then
                    ObjEmail.WriteLoggedDB("MpplService", "BSS_InsertComplainTroubleShooting  Failed", "Unknown")
                    Return False
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "BSS_InsertComplainTroubleShooting  Failed", "Unknown")
                Return False
            End If

            Dim para As String(,) = {
             {"@ComplainID", CInt(ComplainID)}, _
             {"@LastMilePowerStatus", LastMilePowerStatus}, _
             {"@FiberLEDStatus", FiberLEDStatus}, _
             {"@DeviceRebooted", DeviceRebooted}}

            If objBSS.executeProc("sp_app_InsertComplainTroubleShooting", para) Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Throw ex
        End Try



    End Function

    <WebMethod()> _
    <SoapHeader("User", Required:=True)> _
    Public Function BSS_GetComplainsByID(ByVal ComplaintID As String) As DataTable
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("BSS_GetComplainsByID") Then
                    ObjEmail.WriteLoggedDB("MpplService", "BSS_GetComplainsByID  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "BSS_GetComplainsByID  Failed", "Unknown")
                Return Nothing
            End If

            Dim para As String(,) = {{"@ComplaintID", ComplaintID}}
            dt = objBSS.SP_Datatable("sp_OTS_GetComplainByComplainID", para)
            dt.TableName = "sp_OTS_GetComplainByComplainID"
            Return dt

        Catch ex As Exception
            Throw ex
        Finally

        End Try

    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    Public Function BSS_ClosedComplain(ByVal ComplainID As String, _
                                        ByVal ComplaintStatusID As String, _
                                        ByVal CustomerFeedBack As String, _
                                        ByVal FurtherAction As String, _
                                        ByVal Remarks As String) As Boolean
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("BSS_ClosedComplain") Then
                    ObjEmail.WriteLoggedDB("MpplService", "BSS_ClosedComplain  Failed", "Unknown")
                    Return False
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "BSS_ClosedComplain  Failed", "Unknown")
                Return False
            End If
            Dim para As String(,) = {{"@ComplainID", CInt(ComplainID)}, _
                                     {"@ComplaintStatusID", CInt(ComplaintStatusID)}, _
                                     {"@CustomerFeedBack", CustomerFeedBack}, _
                                     {"@FurtherAction", FurtherAction}, _
                                     {"@Remarks", Remarks}}

            If objBSS.executeProc("sp_app_ClosedComplains", para) Then
                Return True
            Else
                Return False

            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod()> _
    <SoapHeader("User", Required:=True)> _
    Public Function BSS_InsertICSComplain(
                ByVal SignupID As Integer, _
                ByVal ComplaintTypeID As Integer, _
                ByVal LinkStatusID As Integer, _
                ByVal Remarks As String) As Integer
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("BSS_InsertICSComplain") Then
                    ObjEmail.WriteLoggedDB("MpplService", "BSS_InsertICSComplain  Failed", "Unknown")
                    Return 0
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "BSS_InsertICSComplain  Failed", "Unknown")
                Return 0
            End If

            Dim para As String(,) = {
             {"@SignupID", SignupID}, _
             {"@ComplaintTypeID", ComplaintTypeID}, _
             {"@LinkStatusID", LinkStatusID}, _
             {"@Remarks", Remarks}, _
             {"@Flag", "WEB"}}

            Dim LastID As Integer = objBSS.InsertProc_GeComplainID("sp_app_InsertICSComplains", para)

            If LastID > 0 Then
                Return LastID
            Else
                Return 0

            End If


        Catch ex As Exception
            Throw ex
        End Try



    End Function

    <WebMethod()> _
    <SoapHeader("User", Required:=True)> _
    Public Function BSS_GetLinkStatus() As DataTable

        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("BSS_GetLinkStatus") Then
                    ObjEmail.WriteLoggedDB("MpplService", "BSS_GetLinkStatus  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "BSS_GetLinkStatus  Failed", "Unknown")
                Return Nothing
            End If

            dt = objBSS.SP_Datatable("sp_ICS_GetLinkStatus")
            dt.TableName = "sp_ICS_GetLinkStatus"
            Return dt

        Catch ex As Exception
            Throw ex
        Finally

        End Try

    End Function

    <WebMethod()> _
    <SoapHeader("User", Required:=True)> _
    Public Function BSS_GetComplaintType() As DataTable

        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("BSS_GetComplaintType") Then
                    ObjEmail.WriteLoggedDB("MpplService", "BSS_GetComplaintType  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "BSS_GetComplaintType  Failed", "Unknown")
                Return Nothing
            End If

            Dim para As String(,) = {{"@ComplaintTypeID", 0}, _
                                     {"@ComplaintType", " "}, _
                                     {"@IsActive", 1}}
            dt = objBSS.SP_Datatable("sp_ICS_GetComplaintType", para)
            dt.TableName = "sp_ICS_GetComplaintType"
            Return dt

        Catch ex As Exception
            Throw ex
        Finally

        End Try

    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    Public Function CreateNewLeads(ByVal CustomerName As String, _
                                    ByVal Infra As String, _
                                    ByVal Poc_Name As String, _
                                    ByVal Poc_ContactNo As String, _
                                    ByVal Poc_Email As String, _
                                    ByVal Address As String, _
                                    ByVal City As String, _
                                    ByVal ServiceUnit As String, _
                                    ByVal Description As String) As Boolean
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("CreateNewLeads") Then
                    ObjEmail.WriteLoggedDB("MpplService", "CreateNewLeads  Failed", "Unknown")
                    Return False
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "CreateNewLeads  Failed", "Unknown")
                Return False
            End If

            If ObjEmail.CreateLeadOracle_RMDB(CustomerName, _
                                              Infra, _
                                              Poc_Name, _
                                              Poc_ContactNo, _
                                              Poc_Email, _
                                              Address, _
                                              City, _
                                              ServiceUnit, _
                                              Description) Then
                Return True
            Else
                Return True

            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    '=======================================Methods for Multinet APP Response in JSON==========================================

    <WebMethod> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    <SoapHeader("User", Required:=True)> _
    Public Function CustomerCredentialJSON(ByVal UserLogin As String, _
                                           ByVal Password As String) As String

        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("CustomerCredentialJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "CustomerCredentialJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "CustomerCredentialJSON  Failed", "Unknown")
                Return "Please provide details"
            End If

            Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
            Dim rows As New List(Of Dictionary(Of String, Object))()
            Dim row As Dictionary(Of String, Object)

            Dim para As String(,) = {{"@UserLogin", UserLogin}, {"@Password", Password}}

            dt = objMW.SP_Datatable("sp_GetUserCredentials", para)
            dt.TableName = "sp_GetUserCredentials"
            For Each dr As DataRow In dt.Rows
                row = New Dictionary(Of String, Object)()
                For Each col As DataColumn In dt.Columns
                    If col.ColumnName = "CustomerLogo" Then
                        row.Add(col.ColumnName, Convert.ToBase64String(ObjEmail.BSS_DownloadNOCFile(dr("CustomerLogo"))))
                    Else
                        row.Add(col.ColumnName, dr(col))
                    End If

                Next
                rows.Add(row)
            Next
            Return serializer.Serialize(rows)

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function GetUserAccountAccessJSON(ByVal UserID As String) As String
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GetUserAccountAccessJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetUserAccountAccessJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetUserAccountAccessJSON  Failed", "Unknown")
                Return "Please provide details"
            End If

            Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
            Dim rows As New List(Of Dictionary(Of String, Object))()
            Dim row As Dictionary(Of String, Object)

            Dim para As String(,) = {{"@UserID", CInt(UserID)}}
            dt = objMW.SP_Datatable("sp_GetUserAccountAccess", para)
            dt.TableName = "sp_GetUserAccountAccess"

            For Each dr As DataRow In dt.Rows
                row = New Dictionary(Of String, Object)()
                For Each col As DataColumn In dt.Columns
                    row.Add(col.ColumnName, dr(col))
                Next
                rows.Add(row)
            Next

            Return serializer.Serialize(rows)

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function GetRegionJSON(ByVal BSSMasterCode As String) As String
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GetRegionJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetRegionJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetRegionJSON  Failed", "Unknown")
                Return "Please provide details"
            End If

            Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
            Dim rows As New List(Of Dictionary(Of String, Object))()
            Dim row As Dictionary(Of String, Object)

            Dim para As String(,) = {{"@BSSMasterCode", CInt(BSSMasterCode)}}
            dt = objMW.SP_Datatable("sp_GetCustomerRegion", para)
            dt.TableName = "sp_GetCustomerRegion"

            For Each dr As DataRow In dt.Rows
                row = New Dictionary(Of String, Object)()
                For Each col As DataColumn In dt.Columns
                    row.Add(col.ColumnName, dr(col))
                Next
                rows.Add(row)
            Next

            Return serializer.Serialize(rows)

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
<ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function GetCitiesJSON(ByVal BSSMasterCode As String, _
                              ByVal Region As String) As String
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GetCitiesJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetCitiesJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetCitiesJSON  Failed", "Unknown")
                Return "Please provide details"
            End If

            Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
            Dim rows As New List(Of Dictionary(Of String, Object))()
            Dim row As Dictionary(Of String, Object)

            Dim para As String(,) = {{"@BSSMasterCode", CInt(BSSMasterCode)}, {"@Region", (Region)}}
            dt = objMW.SP_Datatable("sp_GetCustomerCities", para)
            dt.TableName = "sp_GetCustomerCities"

            For Each dr As DataRow In dt.Rows
                row = New Dictionary(Of String, Object)()
                For Each col As DataColumn In dt.Columns
                    row.Add(col.ColumnName, dr(col))
                Next
                rows.Add(row)
            Next

            Return serializer.Serialize(rows)

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function GetInfraJSON(ByVal BSSMasterCode As String) As String
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GetInfraJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetInfraJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetInfraJSON  Failed", "Unknown")
                Return "Please provide details"
            End If

            Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
            Dim rows As New List(Of Dictionary(Of String, Object))()
            Dim row As Dictionary(Of String, Object)

            Dim para As String(,) = {{"@BSSMasterCode", CInt(BSSMasterCode)}}
            dt = objMW.SP_Datatable("sp_GetCustomerInfra", para)
            dt.TableName = "sp_GetCustomerInfra"

            For Each dr As DataRow In dt.Rows
                row = New Dictionary(Of String, Object)()
                For Each col As DataColumn In dt.Columns
                    row.Add(col.ColumnName, dr(col))
                Next
                rows.Add(row)
            Next

            Return serializer.Serialize(rows)

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function GetLOBJSON(ByVal BSSMasterCode As String) As String
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GetLOBJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetLOBJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetLOBJSON  Failed", "Unknown")
                Return "Please provide details"
            End If

            Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
            Dim rows As New List(Of Dictionary(Of String, Object))()
            Dim row As Dictionary(Of String, Object)

            Dim para As String(,) = {{"@BSSMasterCode", CInt(BSSMasterCode)}}
            dt = objMW.SP_Datatable("sp_GetCustomerLOB", para)
            dt.TableName = "sp_GetCustomerLOB"

            For Each dr As DataRow In dt.Rows
                row = New Dictionary(Of String, Object)()
                For Each col As DataColumn In dt.Columns
                    row.Add(col.ColumnName, dr(col))
                Next
                rows.Add(row)
            Next

            Return serializer.Serialize(rows)

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function GetServiceUnitJSON(ByVal BSSMasterCode As String) As String
        Dim dt As New DataTable
        Try

            If User IsNot Nothing Then
                If Not User.IsValid("GetServiceUnitJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetServiceUnitJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetServiceUnitJSON  Failed", "Unknown")
                Return "Please provide details"
            End If

            Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
            Dim rows As New List(Of Dictionary(Of String, Object))()
            Dim row As Dictionary(Of String, Object)

            Dim para As String(,) = {{"@BSSMasterCode", CInt(BSSMasterCode)}}
            dt = objMW.SP_Datatable("sp_GetCustomerServiceUnit", para)
            dt.TableName = "sp_GetCustomerServiceUnit"

            For Each dr As DataRow In dt.Rows
                row = New Dictionary(Of String, Object)()
                For Each col As DataColumn In dt.Columns
                    row.Add(col.ColumnName, dr(col))
                Next
                rows.Add(row)
            Next

            Return serializer.Serialize(rows)

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function GetBandwidthJSON(ByVal BSSMasterCode As String) As String
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GetBandwidthJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetBandwidthJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetBandwidthJSON  Failed", "Unknown")
                Return "Please provide details"
            End If

            Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
            Dim rows As New List(Of Dictionary(Of String, Object))()
            Dim row As Dictionary(Of String, Object)

            Dim para As String(,) = {{"@BSSMasterCode", CInt(BSSMasterCode)}}
            dt = objMW.SP_Datatable("sp_GetCustomerBandwidth", para)
            dt.TableName = "sp_GetCustomerBandwidth"

            For Each dr As DataRow In dt.Rows
                row = New Dictionary(Of String, Object)()
                For Each col As DataColumn In dt.Columns
                    row.Add(col.ColumnName, dr(col))
                Next
                rows.Add(row)
            Next

            Return serializer.Serialize(rows)

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
<ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function GetCiruitsMainDetailJSON(ByVal BSSMasterCode As String, _
                                       ByVal City As String, _
                                       ByVal Infra As String, _
                                       ByVal ServiceUnit As String, _
                                       ByVal Bandwidth As String, _
                                       ByVal Region As String, _
                                       ByVal IsFavourite As String, _
                                       ByVal UserID As String, _
                                       ByVal NMNStatus As String, _
                                       ByVal LOB As String) As String
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GetCiruitsMainDetailJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetCiruitsMainDetailJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetCiruitsMainDetailJSON  Failed", "Unknown")
                Return "Please provide details"
            End If


            Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
            Dim rows As New List(Of Dictionary(Of String, Object))()
            Dim row As Dictionary(Of String, Object)

            Dim para As String(,) = {{"@BSSMasterCode", (BSSMasterCode.Trim)}, _
                                      {"@City", City}, _
                                      {"@Infra", Infra}, _
                                      {"@ServiceUnit", ServiceUnit}, _
                                      {"@Bandwidth", Bandwidth}, _
                                      {"@Region", Region}, _
                                      {"@IsFavourite", CInt(IsFavourite)}, _
                                      {"@UserID", CInt(UserID)}, _
                                      {"@NMNStatus", NMNStatus}, _
                                      {"@LOB", LOB}}

            dt = objMW.SP_Datatable("sp_GetCiruitsMainDetails", para)
            dt.TableName = "sp_GetCiruitsMainDetails"

            For Each dr As DataRow In dt.Rows
                row = New Dictionary(Of String, Object)()
                For Each col As DataColumn In dt.Columns
                    row.Add(col.ColumnName, dr(col))
                Next
                rows.Add(row)
            Next

            Return serializer.Serialize(rows)

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function GetCircuitCompleteDetailJSON(ByVal SignupID As String) As String
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GetCircuitCompleteDetailJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetCircuitCompleteDetailJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetCircuitCompleteDetailJSON  Failed", "Unknown")
                Return "Please provide details"
            End If

            Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
            Dim rows As New List(Of Dictionary(Of String, Object))()
            Dim row As Dictionary(Of String, Object)

            Dim para As String(,) = {{"@SignupID", CInt(SignupID)}}
            dt = objMW.SP_Datatable("sp_GetCiruitsCompleteDetail", para)
            dt.TableName = "sp_GetCiruitsCompleteDetail"

            For Each dr As DataRow In dt.Rows
                row = New Dictionary(Of String, Object)()
                For Each col As DataColumn In dt.Columns
                    row.Add(col.ColumnName, dr(col))
                Next
                rows.Add(row)
            Next

            Return serializer.Serialize(rows)

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function GetNewConnectionListJSON(ByVal BSSMasterCode As String) As String
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GetNewConnectionListJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetNewConnectionListJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetNewConnectionListJSON  Failed", "Unknown")
                Return "Please provide details"
            End If

            Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
            Dim rows As New List(Of Dictionary(Of String, Object))()
            Dim row As Dictionary(Of String, Object)

            Dim para As String(,) = {{"@BSSMasterCode", CInt(BSSMasterCode)}}
            dt = objMW.SP_Datatable("sp_GetNewConnectionList", para)
            dt.TableName = "sp_GetNewConnectionList"

            For Each dr As DataRow In dt.Rows
                row = New Dictionary(Of String, Object)()
                For Each col As DataColumn In dt.Columns
                    row.Add(col.ColumnName, dr(col))
                Next
                rows.Add(row)
            Next

            Return serializer.Serialize(rows)

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function GetConnectionHistoryJSON(ByVal SignupID As String) As String
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GetConnectionHistoryJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetConnectionHistoryJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetConnectionHistoryJSON  Failed", "Unknown")
                Return "Please provide details"
            End If

            Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
            Dim rows As New List(Of Dictionary(Of String, Object))()
            Dim row As Dictionary(Of String, Object)

            Dim para As String(,) = {{"@SignupID", CInt(SignupID)}}
            dt = objMW.SP_Datatable("sp_GetConnectionHistory", para)
            dt.TableName = "sp_GetConnectionHistory"

            For Each dr As DataRow In dt.Rows
                row = New Dictionary(Of String, Object)()
                For Each col As DataColumn In dt.Columns
                    row.Add(col.ColumnName, dr(col))
                Next
                rows.Add(row)
            Next

            Return serializer.Serialize(rows)

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function GetCustomerMonthlyInvoicesJSON(ByVal BSSMasterCode As Integer, _
                                               ByVal GPMasterCode As String) As String
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GetCustomerMonthlyInvoicesJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetCustomerMonthlyInvoicesJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetCustomerMonthlyInvoicesJSON  Failed", "Unknown")
                Return "Please provide details"
            End If

            Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
            Dim rows As New List(Of Dictionary(Of String, Object))()
            Dim row As Dictionary(Of String, Object)

            Dim para As String(,) = {{"@BSSMasterCode", CInt(BSSMasterCode)}, _
                                     {"@GPMasterCode", GPMasterCode}}
            dt = objMW.SP_Datatable("sp_GetCustomerMonthlyInvoices", para)
            dt.TableName = "sp_GetCustomerMonthlyInvoices"

            For Each dr As DataRow In dt.Rows
                row = New Dictionary(Of String, Object)()
                For Each col As DataColumn In dt.Columns
                    row.Add(col.ColumnName, dr(col))
                Next
                rows.Add(row)
            Next

            Return serializer.Serialize(rows)

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function GetCustomerProdutSummaryJSON(ByVal BSSMasterCode As Integer, _
                                               ByVal GPMasterCode As String) As String
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GetCustomerProdutSummaryJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetCustomerProdutSummaryJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetCustomerProdutSummaryJSON  Failed", "Unknown")
                Return "Please provide details"
            End If

            Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
            Dim rows As New List(Of Dictionary(Of String, Object))()
            Dim row As Dictionary(Of String, Object)

            Dim para As String(,) = {{"@BSSMasterCode", CInt(BSSMasterCode)}, _
                                     {"@GPMasterCode", GPMasterCode}}
            dt = objMW.SP_Datatable("sp_GetCustomerProdutSummary", para)
            dt.TableName = "sp_GetCustomerProdutSummary"

            For Each dr As DataRow In dt.Rows
                row = New Dictionary(Of String, Object)()
                For Each col As DataColumn In dt.Columns
                    row.Add(col.ColumnName, dr(col))
                Next
                rows.Add(row)
            Next

            Return serializer.Serialize(rows)

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function GetCustomerOutstandingJSON(ByVal BSSMasterCode As Integer, _
                                               ByVal GPMasterCode As String) As String
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GetCustomerOutstandingJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetCustomerOutstandingJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetCustomerOutstandingJSON  Failed", "Unknown")
                Return "Please provide details"
            End If

            Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
            Dim rows As New List(Of Dictionary(Of String, Object))()
            Dim row As Dictionary(Of String, Object)

            Dim para As String(,) = {{"@BSSMasterCode", CInt(BSSMasterCode)}, _
                                     {"@GPMasterCode", GPMasterCode}}
            dt = objMW.SP_Datatable("sp_GetCustomerOutstanding", para)
            dt.TableName = "sp_GetCustomerOutstanding"

            For Each dr As DataRow In dt.Rows
                row = New Dictionary(Of String, Object)()
                For Each col As DataColumn In dt.Columns
                    row.Add(col.ColumnName, dr(col))
                Next
                rows.Add(row)
            Next

            Return serializer.Serialize(rows)

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function GetCustomerInovicesJSON(ByVal BSSMasterCode As String, _
                                           ByVal GPMasterCode As String, _
                                           ByVal Period As String, _
                                           ByVal InvoiceNo As String) As String
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GetCustomerInovicesJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetCustomerInovicesJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetCustomerInovicesJSON  Failed", "Unknown")
                Return "Please provide details"
            End If

            Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
            Dim rows As New List(Of Dictionary(Of String, Object))()
            Dim row As Dictionary(Of String, Object)

            Dim para As String(,) = {{"@BSSMasterCode", CInt(BSSMasterCode)}, _
                                     {"@GPMasterCode", GPMasterCode}, _
                                     {"@Period", Period}, _
                                     {"@InvoiceNo", InvoiceNo}}
            dt = objMW.SP_Datatable("sp_GetCustomerInovices", para)
            dt.TableName = "sp_GetCustomerInovices"

            For Each dr As DataRow In dt.Rows
                row = New Dictionary(Of String, Object)()
                For Each col As DataColumn In dt.Columns
                    row.Add(col.ColumnName, dr(col))
                Next
                rows.Add(row)
            Next

            Return serializer.Serialize(rows)

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
<ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function SetFavCircuitJSON(ByVal SignupID As String, _
                                         ByVal UserID As String) As String
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("SetFavCircuitJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "SetFavCircuitJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "SetFavCircuitJSON  Failed", "Unknown")
                Return "Please provide details"
            End If

            Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()

            Dim para As String(,) = {{"@SignupID", CInt(SignupID)}, _
                                     {"@UserID", CInt(UserID)}}
            If objMW.executeProc("sp_SetFavCircuit", para) Then
                Return serializer.Serialize(True)
            Else
                Return serializer.Serialize(False)
            End If


        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function SetUnFavCircuitJSON(ByVal SignupID As String, _
                                    ByVal UserID As String) As String
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("SetUnFavCircuitJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "SetUnFavCircuitJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "SetUnFavCircuitJSON  Failed", "Unknown")
                Return "Please provide details"
            End If

            Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()

            Dim para As String(,) = {{"@SignupID", CInt(SignupID)}, _
                                     {"@UserID", CInt(UserID)}}

            If objMW.executeProc("sp_SetUnFavCircuit", para) Then
                Return serializer.Serialize(True)
            Else
                Return serializer.Serialize(False)
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function GetCustomerFavJSON(ByVal UserID As String, _
                                   ByVal BSSMasterCode As String) As String
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GetCustomerFavJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetCustomerFavJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetCustomerFavJSON  Failed", "Unknown")
                Return "Please provide details"
            End If

            Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
            Dim rows As New List(Of Dictionary(Of String, Object))()
            Dim row As Dictionary(Of String, Object)

            Dim para As String(,) = {{"@UserID", CInt(UserID)}, _
                                     {"@BSSMasterCode", CInt(BSSMasterCode)}}

            dt = objMW.SP_Datatable("sp_GetCustomerFav", para)
            dt.TableName = "sp_GetCustomerFav"

            For Each dr As DataRow In dt.Rows
                row = New Dictionary(Of String, Object)()
                For Each col As DataColumn In dt.Columns
                    row.Add(col.ColumnName, dr(col))
                Next
                rows.Add(row)
            Next

            Return serializer.Serialize(rows)

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function GetInvoiceDetailForPDFJSON(ByVal FROM_DATE As String, _
                                   ByVal TO_DATE As String, _
                                   ByVal SOPNUMBE As String) As String
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GetInvoiceDetailForPDFJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetInvoiceDetailForPDFJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetInvoiceDetailForPDFJSON  Failed", "Unknown")
                Return "Please provide details"
            End If

            Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
            Dim rows As New List(Of Dictionary(Of String, Object))()
            Dim row As Dictionary(Of String, Object)

            Dim para As String(,) = {{"@B_POSTED", 1}, _
                                     {"@FROM_DATE", DateTime.Parse(FROM_DATE)}, _
                                     {"@TO_DATE", DateTime.Parse(TO_DATE)}, _
                                     {"@SOPNUMBE", SOPNUMBE}, _
                                     {"@FromInvPrint", DateTime.Parse(FROM_DATE)}, _
                                     {"@ToInvPrint", DateTime.Parse(TO_DATE)}}

            dt = objGP.SP_Datatable("sp_App_GetDetailsforPDF", para)
            dt.TableName = "sp_App_GetDetailsforPDF"

            For Each dr As DataRow In dt.Rows
                row = New Dictionary(Of String, Object)()
                For Each col As DataColumn In dt.Columns
                    row.Add(col.ColumnName, dr(col))
                Next
                rows.Add(row)
            Next

            Return serializer.Serialize(rows)

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function GetCustomerActivityLog(ByVal BSSMasterCode As String, _
                                   ByVal UserLogin As String) As String
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("GetCustomerFavJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "GetCustomerActivityLog  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "GetCustomerActivityLog  Failed", "Unknown")
                Return "Please provide details"
            End If

            Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
            Dim rows As New List(Of Dictionary(Of String, Object))()
            Dim row As Dictionary(Of String, Object)

            Dim para As String(,) = {{"@BSSMasterCode", CInt(BSSMasterCode)}, _
                                     {"@UserLogin", UserLogin}}

            dt = objMW.SP_Datatable("sp_GetCustomerActivityLog", para)
            dt.TableName = "sp_GetCustomerActivityLog"

            For Each dr As DataRow In dt.Rows
                row = New Dictionary(Of String, Object)()
                For Each col As DataColumn In dt.Columns
                    row.Add(col.ColumnName, dr(col))
                Next
                rows.Add(row)
            Next

            Return serializer.Serialize(rows)

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function InsertCustomerActivityLog(ByVal UserLogin As String, _
                                    ByVal SessionStartTime As String, _
                                    ByVal SessionEndTime As String) As String
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("SetUnFavCircuitJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "InsertCustomerActivityLog  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "InsertCustomerActivityLog  Failed", "Unknown")
                Return "Please provide details"
            End If

            Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()

            Dim para As String(,) = {{"@UserLogin", UserLogin}, _
                                     {"@SessionStartTime", Convert.ToDateTime(SessionStartTime)}, _
                                     {"@SessionEndTime", Convert.ToDateTime(SessionEndTime)}}

            If objMW.executeProc("sp_InsertCustomerActivityLogin", para) Then
                Return serializer.Serialize(True)
            Else
                Return serializer.Serialize(False)
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod()> _
    <SoapHeader("User", Required:=True)> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function BSS_GetInitialStatmentsJSON() As String
        Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
        Dim rows As New List(Of Dictionary(Of String, Object))()
        Dim row As Dictionary(Of String, Object)

        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("BSS_GetInitialStatmentsJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "BSS_GetInitialStatmentsJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "BSS_GetInitialStatmentsJSON  Failed", "Unknown")
                Return "Please provide details"
            End If

            dt = objBSS.SP_Datatable("sp_OTS_GetInitialStatement")
            dt.TableName = "sp_OTS_GetInitialStatement"

            For Each dr As DataRow In dt.Rows
                row = New Dictionary(Of String, Object)()
                For Each col As DataColumn In dt.Columns
                    row.Add(col.ColumnName, dr(col))
                Next
                rows.Add(row)
            Next

            Return serializer.Serialize(rows)


        Catch ex As Exception
            Throw ex
        Finally

        End Try

    End Function

    <WebMethod()> _
    <SoapHeader("User", Required:=True)> _
  <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function BSS_GetComplainsJSON(ByVal BSSMasterCode As String, ByVal GPID As String, ByVal TicketNo As String, ByVal FromDate As String, ByVal ToDate As String) As String
        Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
        Dim rows As New List(Of Dictionary(Of String, Object))()
        Dim row As Dictionary(Of String, Object)

        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("BSS_GetComplainsJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "BSS_GetComplainsJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "BSS_GetComplainsJSON  Failed", "Unknown")
                Return "Please provide details"
            End If

            Dim para As String(,) = {{"@BSSMasterCode", CInt(BSSMasterCode)}, {"@GPID", GPID}, {"@TicketNo", TicketNo}, {"@FromDate", FromDate}, {"@ToDate", ToDate}}

            dt = objBSS.SP_Datatable("sp_app_GetComplains", para)
            dt.TableName = "sp_app_GetComplains"

            For Each dr As DataRow In dt.Rows
                row = New Dictionary(Of String, Object)()
                For Each col As DataColumn In dt.Columns
                    row.Add(col.ColumnName, dr(col))
                Next
                rows.Add(row)
            Next

            Return serializer.Serialize(rows)


        Catch ex As Exception
            Throw ex
        Finally

        End Try

    End Function

    <WebMethod()> _
    <SoapHeader("User", Required:=True)> _
    Public Function BSS_GetinterComplainsDetailsByJSON(ByVal ComplaintID As String) As String
        Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
        Dim rows As New List(Of Dictionary(Of String, Object))()
        Dim row As Dictionary(Of String, Object)
        Dim dt As New DataTable
        Try

            If User IsNot Nothing Then
                If Not User.IsValid("BSS_GetinterComplainsDetailsByJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "BSS_GetinterComplainsDetailsByJSON  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "BSS_GetinterComplainsDetailsByJSON  Failed", "Unknown")
                Return Nothing
            End If

            Dim para As String(,) = {{"@ComplaintID", CInt(ComplaintID)}}
            dt = objBSS.SP_Datatable("sp_ICS_GetComplainDetails", para)
            dt.TableName = "sp_ICS_GetComplainDetails"


            For Each dr As DataRow In dt.Rows
                row = New Dictionary(Of String, Object)()
                For Each col As DataColumn In dt.Columns
                    row.Add(col.ColumnName, dr(col))
                Next
                rows.Add(row)
            Next

            Return serializer.Serialize(rows)


        Catch ex As Exception
            Throw ex
        Finally

        End Try

    End Function

    <WebMethod()> _
    <SoapHeader("User", Required:=True)> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function BSS_GetInternationalComplainsJSON(ByVal GPID As String, ByVal TicketNo As String, ByVal FromDate As String, ByVal ToDate As String) As String
        Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
        Dim rows As New List(Of Dictionary(Of String, Object))()
        Dim row As Dictionary(Of String, Object)

        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("BSS_GetInternationalComplainsJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "BSS_GetInternationalComplainsJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "BSS_GetInternationalComplainsJSON  Failed", "Unknown")
                Return "Please provide details"
            End If

            Dim para As String(,) = {{"@GPID", GPID}, {"@TicketNo", TicketNo}, {"@FromDate", FromDate}, {"@ToDate", ToDate}}

            dt = objBSS.SP_Datatable("sp_web_SearchICSComplains", para)
            dt.TableName = "sp_web_SearchICSComplains"


            For Each dr As DataRow In dt.Rows
                row = New Dictionary(Of String, Object)()
                For Each col As DataColumn In dt.Columns
                    row.Add(col.ColumnName, dr(col))
                Next
                rows.Add(row)
            Next

            Return serializer.Serialize(rows)


        Catch ex As Exception
            Throw ex
        Finally

        End Try

    End Function

    <WebMethod()> _
    <SoapHeader("User", Required:=True)> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function BSS_InsertComplainJSON(
                ByVal SignupID As Integer, _
                ByVal InitailStatementID As Integer, _
                ByVal PoCName As String, _
                ByVal PoCNumber As String, _
                ByVal Remarks As String) As Integer

        Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()

        Try
            If User IsNot Nothing Then
                If Not User.IsValid("BSS_InsertComplainJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "BSS_InsertComplainJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "BSS_InsertComplainJSON  Failed", "Unknown")
                Return "Please provide details"
            End If

            Dim para As String(,) = {
             {"@SignupID", SignupID}, _
             {"@InitailStatementID", InitailStatementID}, _
             {"@PoCName", PoCName}, _
             {"@PoCNumber", PoCNumber}, _
             {"@Remarks", Remarks}, _
             {"@Flag", "APP"}}

            Dim LastID As Integer = objBSS.InsertProc_GeComplainID("sp_app_InsertComplains", para)

            If LastID > 0 Then
                Return serializer.Serialize(LastID)
            Else
                Return serializer.Serialize(0)

            End If


        Catch ex As Exception
            Throw ex
        End Try



    End Function

    <WebMethod()> _
    <SoapHeader("User", Required:=True)> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function BSS_InsertComplainTroubleShootingJSON(
                ByVal ComplainID As String, _
                ByVal LastMilePowerStatus As String, _
                ByVal FiberLEDStatus As String, _
                ByVal DeviceRebooted As String) As String

        Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()

        Try
            If User IsNot Nothing Then
                If Not User.IsValid("BSS_InsertComplainTroubleShootingJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "BSS_InsertComplainTroubleShootingJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "BSS_InsertComplainTroubleShootingJSON  Failed", "Unknown")
                Return "Please provide details"
            End If

            Dim para As String(,) = {
             {"@ComplainID", CInt(ComplainID)}, _
             {"@LastMilePowerStatus", LastMilePowerStatus}, _
             {"@FiberLEDStatus", FiberLEDStatus}, _
             {"@DeviceRebooted", DeviceRebooted}}

            If objBSS.executeProc("sp_app_InsertComplainTroubleShooting", para) Then
                Return serializer.Serialize(True)
            Else
                Return serializer.Serialize(False)
            End If

        Catch ex As Exception
            Throw ex
        End Try



    End Function

    <WebMethod()> _
    <SoapHeader("User", Required:=True)> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function BSS_GetComplainsByIDJSON(ByVal ComplaintID As String) As String
        Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
        Dim rows As New List(Of Dictionary(Of String, Object))()
        Dim row As Dictionary(Of String, Object)

        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("BSS_GetComplainsByIDJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "BSS_GetComplainsByIDJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "BSS_GetComplainsByIDJSON  Failed", "Unknown")
                Return "Please provide details"
            End If

            Dim para As String(,) = {{"@ComplaintID", ComplaintID}}

            dt = objBSS.SP_Datatable("sp_OTS_GetComplainByComplainID", para)
            dt.TableName = "sp_OTS_GetComplainByComplainID"


            For Each dr As DataRow In dt.Rows
                row = New Dictionary(Of String, Object)()
                For Each col As DataColumn In dt.Columns
                    row.Add(col.ColumnName, dr(col))
                Next
                rows.Add(row)
            Next

            Return serializer.Serialize(rows)


        Catch ex As Exception
            Throw ex
        Finally

        End Try

    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function BSS_ClosedComplainJSON(ByVal ComplainID As String, _
                                        ByVal ComplaintStatusID As String, _
                                        ByVal CustomerFeedBack As String, _
                                        ByVal FurtherAction As String, _
                                        ByVal Remarks As String) As String
        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("BSS_ClosedComplainJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "BSS_ClosedComplainJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "BSS_ClosedComplainJSON  Failed", "Unknown")
                Return "Please provide details"
            End If

            Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()

            Dim para As String(,) = {{"@ComplainID", CInt(ComplainID)}, _
                                     {"@ComplaintStatusID", CInt(ComplaintStatusID)}, _
                                     {"@CustomerFeedBack", CustomerFeedBack}, _
                                     {"@FurtherAction", FurtherAction}, _
                                     {"@Remarks", Remarks}}

            If objBSS.executeProc("sp_app_ClosedComplains", para) Then
                Return serializer.Serialize(True)
            Else
                Return serializer.Serialize(False)

            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    <WebMethod()> _
    <SoapHeader("User", Required:=True)> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function BSS_InsertICSComplainJSON(
                ByVal SignupID As String, _
                ByVal ComplaintTypeID As String, _
                ByVal LinkStatusID As String, _
                ByVal Remarks As String) As String

        Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()

        Try
            If User IsNot Nothing Then
                If Not User.IsValid("BSS_InsertICSComplainJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "BSS_InsertICSComplainJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "BSS_InsertICSComplainJSON  Failed", "Unknown")
                Return "Please provide details"
            End If

            Dim para As String(,) = {
             {"@SignupID", CInt(SignupID)}, _
             {"@ComplaintTypeID", CInt(ComplaintTypeID)}, _
             {"@LinkStatusID", CInt(LinkStatusID)}, _
             {"@Remarks", Remarks}, _
             {"@Flag", "APP"}}

            Dim LastID As Integer = objBSS.InsertProc_GeComplainID("sp_app_InsertICSComplains", para)

            If LastID > 0 Then
                Return serializer.Serialize(LastID)
            Else
                Return serializer.Serialize(0)

            End If


        Catch ex As Exception
            Throw ex
        End Try



    End Function

    <WebMethod()> _
    <SoapHeader("User", Required:=True)> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function BSS_GetLinkStatusJSON() As String
        Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
        Dim rows As New List(Of Dictionary(Of String, Object))()
        Dim row As Dictionary(Of String, Object)

        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("BSS_GetLinkStatusJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "BSS_GetLinkStatusJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "BSS_GetLinkStatusJSON  Failed", "Unknown")
                Return "Please provide details"
            End If

            dt = objBSS.SP_Datatable("sp_ICS_GetLinkStatus")
            dt.TableName = "sp_ICS_GetLinkStatus"

            For Each dr As DataRow In dt.Rows
                row = New Dictionary(Of String, Object)()
                For Each col As DataColumn In dt.Columns
                    row.Add(col.ColumnName, dr(col))
                Next
                rows.Add(row)
            Next

            Return serializer.Serialize(rows)


        Catch ex As Exception
            Throw ex
        Finally

        End Try

    End Function

    <WebMethod()> _
    <SoapHeader("User", Required:=True)> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function BSS_GetComplaintTypeJSON() As String
        Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
        Dim rows As New List(Of Dictionary(Of String, Object))()
        Dim row As Dictionary(Of String, Object)

        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("BSS_GetComplaintTypeJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "BSS_GetComplaintTypeJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "BSS_GetComplaintTypeJSON  Failed", "Unknown")
                Return "Please provide details"
            End If


            Dim para As String(,) = {{"@ComplaintTypeID", 0}, _
                                     {"@ComplaintType", " "}, _
                                     {"@IsActive", 1}}
            dt = objBSS.SP_Datatable("sp_ICS_GetComplaintType", para)
            dt.TableName = "sp_ICS_GetComplaintType"

            For Each dr As DataRow In dt.Rows
                row = New Dictionary(Of String, Object)()
                For Each col As DataColumn In dt.Columns
                    row.Add(col.ColumnName, dr(col))
                Next
                rows.Add(row)
            Next

            Return serializer.Serialize(rows)

        Catch ex As Exception
            Throw ex
        Finally

        End Try

    End Function

    <WebMethod()> _
   <SoapHeader("User", Required:=True)> _
   <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function BSS_GetCloseComplainsJSON(ByVal BSSMasterCode As String, _
                                                ByVal TicketNo As String) As String


        Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
        Dim rows As New List(Of Dictionary(Of String, Object))()
        Dim row As Dictionary(Of String, Object)

        Dim dt As New DataTable
        Try
            If User IsNot Nothing Then
                If Not User.IsValid("BSS_GetComplaintTypeJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "BSS_GetCloseComplainsJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "BSS_GetCloseComplainsJSON  Failed", "Unknown")
                Return "Please provide details"
            End If


            Dim para As String(,) = {{"@BSSMasterCode", CInt(BSSMasterCode)}, _
                                     {"@TicketNo", TicketNo}}

            dt = objBSS.SP_Datatable("sp_app_GetCloseComplains", para)
            dt.TableName = "sp_app_GetCloseComplains"

            For Each dr As DataRow In dt.Rows
                row = New Dictionary(Of String, Object)()
                For Each col As DataColumn In dt.Columns
                    row.Add(col.ColumnName, dr(col))
                Next
                rows.Add(row)
            Next

            Return serializer.Serialize(rows)

        Catch ex As Exception
            Throw ex
        Finally

        End Try

    End Function

    <WebMethod> _
    <SoapHeader("User", Required:=True)> _
<ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Function CreateNewLeadsJSON(ByVal CustomerName As String, _
                                    ByVal Infra As String, _
                                    ByVal Poc_Name As String, _
                                    ByVal Poc_ContactNo As String, _
                                    ByVal Poc_Email As String, _
                                    ByVal Address As String, _
                                    ByVal City As String, _
                                    ByVal ServiceUnit As String, _
                                    ByVal Description As String) As String
        Try
            ' User.IsValid("CreateNewLeadsJSON")
            If User IsNot Nothing Then
                If Not User.IsValid("CreateNewLeadsJSON") Then
                    ObjEmail.WriteLoggedDB("MpplService", "CreateNewLeadsJSON  Failed", "Unknown")
                    Return "Invalid User details!"
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "CreateNewLeadsJSON  Failed", "Unknown")
                Return "Please provide details"
            End If

            Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()

            If ObjEmail.CreateLeadOracle_RMDB(CustomerName, _
                                              Infra, _
                                              Poc_Name, _
                                              Poc_ContactNo, _
                                              Poc_Email, _
                                              Address, _
                                              City, _
                                              ServiceUnit, _
                                              Description) Then
                Return serializer.Serialize(True)
            Else
                Return serializer.Serialize(False)

            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    '=======================================Methods for BSS Esclation==========================================

    <WebMethod()> _
  <SoapHeader("User", Required:=True)> _
    Public Function BSS_EsclationCall(ByVal DepartmentID As Integer) As Boolean

        Dim dt As New DataTable
        Try
            'User.IsValid("BSS_EsclationCall")
            If User IsNot Nothing Then
                If Not User.IsValid("BSS_EsclationCall") Then
                    ObjEmail.WriteLoggedDB("MpplService", "BSS_EsclationCall  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "BSS_EsclationCall  Failed", "Unknown")
                Return Nothing
            End If

            If ObjEmail.EsclationEmail(DepartmentID) Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Return False
        Finally

        End Try

    End Function

    <WebMethod()> _
    <SoapHeader("User", Required:=True)> _
    Public Function BSS_EsclationCallForICS(ByVal Flag As String) As String

        Dim dt As New DataTable
        Try
            User.IsValid("BSS_EsclationCallForICS")
            If User IsNot Nothing Then
                If Not User.IsValid("BSS_EsclationCallForICS") Then
                    ObjEmail.WriteLoggedDB("MpplService", "BSS_EsclationCallForICS  Failed", "Unknown")
                    Return Nothing
                End If
            Else
                ObjEmail.WriteLoggedDB("MpplService", "BSS_EsclationCallForICS  Failed", "Unknown")
                Return Nothing
            End If

            Dim Result As String = ObjEmail.EsclationEmail_ICS(Flag)
            Return Result

        Catch ex As Exception
            Return False
        Finally

        End Try

    End Function

End Class