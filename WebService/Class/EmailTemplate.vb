﻿Imports System.Net.Mail
Imports System.IO
Imports System.Net
Imports Novell.Directory
Imports Novell.Directory.Ldap
Imports System.Security.Cryptography
Imports System.Text

Public Class EmailTemplate
    Dim Smtp_Server As SmtpClient
    Dim objBSS As New DBAcceses(My.Settings.conBSS, DBEngineType.SQL)
    Dim objRM As New DBAcceses(My.Settings.conRM, DBEngineType.SQL)
    Dim objMW As New DBAcceses(My.Settings.conMidd, DBEngineType.SQL)
    Dim objGP As New DBAcceses(My.Settings.conGP, DBEngineType.SQL)

    Public Function CustomerExternalEmail(ByVal TicketTypeID As Integer, ByVal CaseCategoryID As Integer, ByVal Stage As String, ByVal ComplaintID As Integer) As Boolean
        Try
            Dim para As String(,) = {{"@ComplaintID", ComplaintID}}
            Dim para1 As String(,) = {{"@Stage", Stage}}
            'Dim para2 As String(,) = {{"@Stage", "Active"}}

            Smtp_Server = New SmtpClient
            Dim msg As MailMessage
            msg = New MailMessage

            Dim dt = objBSS.SP_Datatable("sp_OTS_GetComplainByComplainID", para)
            If dt.Rows.Count = 1 Then
                msg.To.Add(dt.Rows(0)("Contact_Email"))
                ' msg.To.Add("abdul.sami@multinet.com.pk")
                ' msg.CC.Add("abdul.sami@multinet.com.pk")
                msg.CC.Add("support@multinet.com.pk")
                msg.IsBodyHtml = True
                If TicketTypeID = 6 Then
                    If CaseCategoryID = 1 Or CaseCategoryID = 2 Then
                        msg.From = New MailAddress("support@multinet.com.pk")
                        msg.Subject = "Support Request(" + dt.Rows(0)("CircuitName") + ")TT # " + dt.Rows(0)("TicketNo")
                        msg.Body = "<html><body><br><p>Attn: </strong>" & dt.Rows(0)("CircuitName") & "</strong></p><h3>Dear Valued Customer,</h3><p>Thank you for contacting Multinet support and using our services.</p><p>This is an automatic reply just to let you know that we have received your <strong>Support/Service Request</strong> and working eagerly to resolve your query.</p><p>Our support team will get back to you shortly. For your records, the details of the ticket are listed below.<br>Please make sure to keep ticket number in subject line while replying, this will ensure tracking of your replies appropriately.<br></p><table width='539' border='1'><tr><th colspan='2'><b>Complaint Information</b></th></tr><tr><td width='132'>Ticket No:</td><td width='391'>" & dt.Rows(0)("TicketNo") & "</td></tr><tr><td width='132'>Customer Name:</td><td>" & dt.Rows(0)("CircuitName") & "</td></tr><tr><td width='132'>Complaint Date:</td><td>" & dt.Rows(0)("ComplaintReceivedDate") & " </td></tr><tr><td width='132'>Call Received by: </td><td>" & dt.Rows(0)("LoggedBy") & "</td></tr></table><p>Just remember, we are just a call or an email away if you need further assistance. We will be happy to follow up with you.</p><p><i><b>Our Contact Details:</b></i><br>Helpline: 111-247-000<br>Email: support@multinet.com.pk</p><p>Regards,</p><h4>Multinet Support Group</h4></body></html>"
                    ElseIf CaseCategoryID = 3 Then
                        msg.From = New MailAddress("monitoring@multinet.com.pk")
                        msg.Subject = "Self Escalation(" + dt.Rows(0)("CircuitName") + ")TT # " + dt.Rows(0)("TicketNo")
                        msg.Body = "<html><body><br><p>Attn: </strong>" & dt.Rows(0)("CircuitName") & "</strong></p><h3>Dear Valued Customer,</h3><p>This is to inform you that the subjected link is showing down on our Network Monitoring System (NMS).<br> We request you to kindly check and confirm the current power and device status at your end in order to help pinpoint the issue.</p><p>This event has been received and logged under the ticket number <b> " & dt.Rows(0)("TicketNo") & ".</b>  Rest assured that we are currently investigating the issue, and will resolve it as soon as possible.</p><br><table width='477' border='1'><tr><th colspan='2'><b>Circuit Information</b></th></tr><tr><td width='112'>Circuit Name:</td><td width='349'>" & dt.Rows(0)("CircuitName") & "</td></tr><tr><td width='112'>Alert Received:</td><td>" & dt.Rows(0)("ComplaintReceivedDate") & "</td></tr><tr><td width='112'>Generated by: </td><td>" & dt.Rows(0)("LoggedBy") & "</td></tr></table><p>We look forward to your response.</p><p>Regards,</p> <h4>Multinet Monitoring Team</h4></body></html>"
                    End If
                Else

                    If TicketTypeID = 1 Then
                        msg.Body = ""
                    ElseIf TicketTypeID = 2 Then
                        msg.Body = ""

                    ElseIf TicketTypeID = 3 Then
                        msg.Body = ""

                    ElseIf TicketTypeID = 4 Then
                        msg.Body = ""

                    ElseIf TicketTypeID = 5 Then
                        If CaseCategoryID = 1 Or CaseCategoryID = 2 Then
                            msg.From = New MailAddress("support@multinet.com.pk")
                            msg.Subject = "Support Request(" + dt.Rows(0)("CircuitName") + ")TT # " + dt.Rows(0)("TicketNo")
                            msg.Body = "<html><body><br><p>Attn: </strong>" & dt.Rows(0)("CircuitName") & "</strong></p><h3>Dear Valued Customer,</h3><p>Thank you for contacting Multinet support and using our services.</p><p>This is an automatic reply just to let you know that we have received your <strong>Support/Service Request</strong> and working eagerly to resolve your query.</p><p>Our support team will get back to you shortly. For your records, the details of the ticket are listed below.<br>Please make sure to keep ticket number in subject line while replying, this will ensure tracking of your replies appropriately.<br></p><table width='539' border='1'><tr><th colspan='2'><b>Complaint Information</b></th></tr><tr><td width='132'>Ticket No:</td><td width='391'>" & dt.Rows(0)("TicketNo") & "</td></tr><tr><td width='132'>Customer Name:</td><td>" & dt.Rows(0)("CircuitName") & "</td></tr><tr><td width='132'>Complaint Date:</td><td>" & dt.Rows(0)("ComplaintReceivedDate") & " </td></tr><tr><td width='132'>Call Received by: </td><td>" & dt.Rows(0)("LoggedBy") & "</td></tr></table><p>Just remember, we are just a call or an email away if you need further assistance. We will be happy to follow up with you.</p><p><i><b>Our Contact Details:</b></i><br>Helpline: 111-247-000<br>Email: support@multinet.com.pk</p><p>Regards,</p><h4>Multinet Support Group</h4></body></html>"
                        ElseIf CaseCategoryID = 3 Then
                            msg.From = New MailAddress("monitoring@multinet.com.pk")
                            msg.Subject = "Self Escalation(" + dt.Rows(0)("CircuitName") + ")TT # " + dt.Rows(0)("TicketNo")
                            msg.Body = "<html><body><br><p>Attn: </strong>" & dt.Rows(0)("CircuitName") & "</strong></p><h3>Dear Valued Customer,</h3><p>This is to inform you that the subjected link is showing down on our Network Monitoring System (NMS).<br> We request you to kindly check and confirm the current power and device status at your end in order to help pinpoint the issue.</p><p>This event has been received and logged under the ticket number <b> " & dt.Rows(0)("TicketNo") & ".</b>  Rest assured that we are currently investigating the issue, and will resolve it as soon as possible.</p><br><table width='477' border='1'><tr><th colspan='2'><b>Circuit Information</b></th></tr><tr><td width='112'>Circuit Name:</td><td width='349'>" & dt.Rows(0)("CircuitName") & "</td></tr><tr><td width='112'>Alert Received:</td><td>" & dt.Rows(0)("ComplaintReceivedDate") & "</td></tr><tr><td width='112'>Generated by: </td><td>" & dt.Rows(0)("LoggedBy") & "</td></tr></table><p>We look forward to your response.</p><p>Regards,</p> <h4>Multinet Monitoring Team</h4></body></html>"
                        End If
                    End If

                End If


            End If

            msg.Priority = MailPriority.High
            Smtp_Server.Host = "202.142.160.14"
            Smtp_Server.Send(msg)
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Function ComplainInternalEmail(ByVal ComplaintID As Integer) As Boolean
        Try
            Dim para As String(,) = {{"@ComplaintID", ComplaintID}}

            Smtp_Server = New SmtpClient
            Dim msg As MailMessage
            msg = New MailMessage
            msg.From = New MailAddress("support@multinet.com.pk")
            msg.To.Add("ots@multinet.com.pk")
            msg.CC.Add("liaqat.hussain@multinet.com.pk")
            msg.CC.Add("support@multinet.com.pk")


            Dim dt = objBSS.SP_Datatable("sp_OTS_GetComplainByComplainID", para)

            If dt.Rows.Count = 1 Then

                If dt.Rows(0)("TicketTypeID") = 6 And dt.Rows(0)("ComplaintStatusID") = 2 Then
                    msg.Subject = "Complaint Open : " + dt.Rows(0)("CircuitName") + ") TT # " + dt.Rows(0)("TicketNo")
                    msg.Body = "<html xmlns='http://www.w3.org/1999/xhtml'><head><meta http-equiv='Content-Type' content='text/html; charset=utf-8' /><style type='text/css'>.style1 {color: #FFFFFF}.style10 {font-family: Georgia, 'Times New Roman', Times, serif}.style12 {font-size: 12px}.style9 {font-family: Calibri; font-size: 12px; }.style19 {font-family: Calibri}.style20 {color: #FFFFFF; font-family: Calibri; font-size: 12px; }--></style></head><body><p align='center'>&nbsp;<strong> &nbsp;&nbsp;Complaint Information </strong></p><table width='449' height='224' border='1' align='center' cellpadding='0' cellspacing='0' bordercolor='#000000'><tr><td width='157' height='28' bordercolor='#000000' bgcolor='#666666'><p class='style20'><strong>Ticket No </strong></p></td><td width='380'><p class='style9'> " + dt.Rows(0)("TicketNo") + "</p></td></tr><tr><td width='157' height='28' bordercolor='#000000' bgcolor='#666666'><p class='style20'><strong>Signup ID </strong></p></td><td width='380'><p class='style9'> " + Convert.ToString(dt.Rows(0)("SignupID")) + "</p></td></tr><tr><td width='157' height='28' bordercolor='#000000' bgcolor='#666666'><p class='style20'><strong>Customer Name </strong></p></td><td width='380'><p class='style9'> " + dt.Rows(0)("CircuitName") + "</p></td></tr><tr><td width='157' height='28' bordercolor='#000000' bgcolor='#666666'><p class='style20'><strong>Remarks</strong></p></td><td width='380'><p class='style9'> " + dt.Rows(0)("Remarks") + "</p></td></tr><tr><td width='157' height='28' bordercolor='#000000' bgcolor='#666666'><p class='style20'><strong>Status </strong></p></td><td width='380'><p class='style9'>" + dt.Rows(0)("ComplainStatus") + "</p></td></tr><tr><td width='157' height='28' bordercolor='#000000' bgcolor='#666666'><p class='style20'><strong>Forward to Department </strong></p></td><td width='380'><p class='style9'>" + dt.Rows(0)("AssignDepartment") + "</p></td></tr><tr><td width='157' height='28' bordercolor='#000000' bgcolor='#666666'><p class='style20'><strong>Date Time </strong></p></td><td width='380'><p class='style9'>" + dt.Rows(0)("LastUpdatedDate") + "</p></td></tr><tr><td width='157' height='28' bordercolor='#000000' bgcolor='#666666'><p class='style20'><strong>Updated By </strong></p></td><td width='380'><p class='style9'>" + dt.Rows(0)("TransactionByName") + "</p></td></tr></table><p><span class='style9'><span class='style10'>*This is an autogenerated email from BSS, incase of any query please contact OTS</span></span></p></body></html>"
                    msg.IsBodyHtml = True
                    msg.Priority = MailPriority.High
                    Smtp_Server.Host = "202.142.160.14"
                    Smtp_Server.Send(msg)
                    Return True
                Else
                    Return False
                End If

            End If


        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Function ValidateActiveDirectoryLogin(ByVal Username As String, ByVal Password As String) As Boolean

        Try
            Dim ldapConn As New LdapConnection()
            ldapConn.Connect("103.31.81.82", 389)
            ldapConn.Bind("cn=administrator,dc=multinet,dc=com,dc=pk", "Zim2017Dep!@#")
            Password = GetBase64EncodedSHA1Hash(Password)
            Dim Filters As String = "(&(uid=" + Username + ")(userPassword=" + Password + "))"
            Dim queue As LdapSearchQueue = ldapConn.Search("ou=people,dc=multinet,dc=com,dc=pk", LdapConnection.SCOPE_ONE, Filters, Nothing, False, DirectCast(Nothing, LdapSearchQueue), _
            DirectCast(Nothing, LdapSearchConstraints))
            Dim message As LdapMessage
            message = queue.getResponse()
            If TypeOf message Is LdapSearchResult Then
                ldapConn.Disconnect()
                Return True
            Else
                ldapConn.Disconnect()
                Return False
            End If

        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function GetBase64EncodedSHA1Hash(filename As String) As String
        Dim bytes As Byte() = Encoding.UTF8.GetBytes(filename)
        Using sha1 As New SHA1Managed()
            Return "{SHA}" + Convert.ToBase64String(sha1.ComputeHash(bytes))
            'Dim result As String = Convert.ToBase64String(sha1.ComputeHash(bytes))
            'Return GetBase64EncodedSHA256Hash(result)
        End Using
    End Function

    Public Function BSS_DownloadNOCFile(ByVal filename As String) As Byte()
        Try
            Dim fs As FileStream = Nothing
            'fs = New FileStream("D:\BSSDocs\BSS-IPCore\" & filename, FileMode.Open)
            'Dim fi As FileInfo = New FileInfo("D:\BSSDocs\BSS-IPCore\" & filename)
            fs = New FileStream("C:\inetpub\wwwroot\CustomerImages\" & filename, FileMode.Open)
            Dim fi As FileInfo = New FileInfo("C:\inetpub\wwwroot\CustomerImages\" & filename)
            Dim temp As Long = fi.Length
            Dim lung As Integer = Convert.ToInt32(temp)
            Dim picture As Byte() = New Byte(lung - 1) {}
            fs.Read(picture, 0, lung)
            fs.Close()
            Return picture

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Function CreateLeadOracle_RMDB(ByVal CustomerName As String, _
                                   ByVal Infra As String, _
                                   ByVal Poc_Name As String, _
                                   ByVal Poc_ContactNo As String, _
                                   ByVal Poc_Email As String, _
                                   ByVal Address As String, _
                                   ByVal City As String, _
                                   ByVal ServiceUnit As String, _
                                   ByVal Description As String) As Boolean

        Try
            'Dim strSOAPRequestBody1 As String = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:typ=""http://xmlns.oracle.com/apps/marketing/leadMgmt/leads/leadServiceV3/types/""" & " xmlns:lead=""http://xmlns.oracle.com/oracle/apps/marketing/leadMgmt/leads/leadService/"" xmlns:lead1=""http://xmlns.oracle.com/apps/marketing/leadMgmt/leads/leadService/""" & " xmlns:not=""http://xmlns.oracle.com/apps/crmCommon/notes/noteService""" & " xmlns:not1=""http://xmlns.oracle.com/apps/crmCommon/notes/notes/flex/noteDff/"">" & _
            '"<soapenv:Header/> " & _
            '    "<soapenv:Body> " & _
            '        "<typ:createSalesLead> " & _
            '            "<typ:salesLead> " & _
            '                "<lead:Name>" & CustomerName & "</lead:Name> " & _
            '                " <lead:JobTitle>JobTitle</lead:JobTitle>" & _
            '                " <lead:Description>" & Description & "</lead:Description>" & _
            '                "<lead:PrimaryContactEmailAddress>" & Poc_Email & "</lead:PrimaryContactEmailAddress>" & _
            '                "<lead:PrimaryContactAddress1>Address1</lead:PrimaryContactAddress1>" & _
            '                "<lead:PrimaryContactAddress2>Address2</lead:PrimaryContactAddress2>" & _
            '                "<lead:PrimaryContactAddress3>Address3</lead:PrimaryContactAddress3>" & _
            '                "<lead:PrimaryContactCity>CityName</lead:PrimaryContactCity>" & _
            '                "<lead:PrimaryContactPartyName>" & Poc_ContactNo & "</lead:PrimaryContactPartyName>" & _
            '            "</typ:salesLead> " & _
            '        "</typ:createSalesLead> " & _
            '    "</soapenv:Body> </soapenv:Envelope>"
            'Dim request2 As HttpWebRequest
            'Dim URI As String = "https://cahg.crm.ap2.oraclecloud.com/mklLeads/LeadIntegrationService?WSDL"
            'request2 = DirectCast(WebRequest.Create(URI), HttpWebRequest)
            'request2.Headers.Add("SOAPAction", "http://xmlns.oracle.com/apps/marketing/leadMgmt/leads/leadService/createSalesLead")


            'Dim ObjNetwork As New NetworkCredential()

            'ObjNetwork.UserName = "liaqat.ibu"
            'ObjNetwork.Password = "Multi@2016"


            'request2.Credentials = ObjNetwork
            'request2.Method = "POST"

            'request2.ContentType = "text/xml; charset=utf-8"
            'request2.ContentLength = strSOAPRequestBody1.Length



            'Using reqStream = request2.GetRequestStream()
            '    Dim streamWriter As New System.IO.StreamWriter(reqStream)
            '    streamWriter.Write(strSOAPRequestBody1)
            '    streamWriter.Close()
            '    'HttpWebResponse webResp = (HttpWebResponse)request2.GetResponse();
            '    'webResp.Close();

            '    Using webresponse As HttpWebResponse = DirectCast(request2.GetResponse(), HttpWebResponse)
            '        'Console.WriteLine("Added for " + Convert.ToString(dr["Lead_LeadName"]));
            'If WebResponse.StatusCode = HttpStatusCode.OK Then
            Dim para As String(,) = {{"@CustomerName", CustomerName}, _
                                    {"@Infra", Infra}, _
                                    {"@Poc_Name", Poc_Name}, _
                                    {"@Poc_ContactNo", Poc_ContactNo}, _
                                    {"@Poc_Email", Poc_Email}, _
                                    {"@Address", Address}, _
                                    {"@City", City}, _
                                    {"@ServiceUnit", ServiceUnit}, _
                                    {"@Description", Description}}

            If objRM.executeProc("sp_app_InsertLeads", para) Then
                ' ConformationEmail(CustomerName, Poc_ContactNo, Poc_Email, Description)
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function ConformationEmail(ByVal CustomerName As String, _
                                     ByVal Poc_ContactNo As String, _
                                     ByVal Poc_Email As String, _
                                     ByVal Description As String) As Boolean

        Try
            Smtp_Server = New SmtpClient
            Dim msg As MailMessage
            Dim EmailFormat As String = ""


            'Email Validation
            If validateEmail(Poc_Email) = True Then
                msg = New MailMessage
                msg.Subject = "Welcome " + (CustomerName) + " (" & Format(Date.Now, "MMM-yyyy") & ")"
                msg.From = New MailAddress("info@multinet.com.pk")
                msg.To.Add("liaqat.hussain@multinet.com.pk")
                msg.CC.Add("shahbaz.iqbal@multinet.com.pk")
                msg.Body = "Thanks for New connection"

                'EmailFormat = My.Settings.EmailFormat
                'Dim EmailFormatForAll As String
                'If My.Settings.EmailFormat <> "" Then
                '    Dim arr() As String = Split(My.Settings.EmailFormat, "!")
                '    EmailFormatForAll = arr(0).ToString & Format(Date.Now, "dd-MMM-yyyy hh:mm:ss") & arr(2).ToString & CustomerName & arr(4).ToString _
                '                         & Description & arr(6).ToString & Poc_ContactNo & arr(8).ToString & Poc_Email & arr(10).ToString _
                '                         & Format(Date.Now, "dd-MMM-yyyy") & arr(12).ToString & arr(13).ToString & arr(14).ToString
                '    msg.Body = EmailFormatForAll

                'End If

            End If
            msg.IsBodyHtml = True
            msg.Priority = MailPriority.High
            Smtp_Server.Host = "202.142.160.14"
            Smtp_Server.Send(msg)
            Return True



        Catch ex As Exception
            'Throw ex
        End Try
    End Function

    Public Function validateEmail(ByVal emailAddress As String) As Boolean
        ' Dim email As New Regex("^(?<user>[^@]+)@(?<host>.+)$")
        Dim email As New Regex("^[a-zA-Z][\w\.-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$")
        If email.IsMatch(emailAddress) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub WriteLoggedDB(ByVal ServiceName As String, ByVal MethodName As String, ByVal TransactionBy As String)
        Try

            Dim para As String(,) = {{"@ServiceName", ServiceName}, _
                                     {"@MethodName", MethodName}, _
                                     {"@TransactionBy", TransactionBy}}
            objMW.executeProcess("sp_InsertServiceLogg", para)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Function GetMethodViaUser(ByVal UserID As Integer) As DataTable
        Try

            Dim para As String(,) = {{"@UserID", UserID}}
            Dim dt As DataTable = objMW.SP_Datatable("sp_GetServiceMethodViaUsers", para)
            If dt.Rows.Count > 0 Then
                Return dt
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function EsclationEmail(ByVal DepartmentID As Integer) As Boolean
        Try
            Dim str As String = "<tr><td width='85' height='22' bgcolor='#FFFFFF' class='style9'><strong>SignupID</strong></td><td width='147' bgcolor='#FFFFFF' class='style9'><strong>Circuit Name</strong></td><td width='87' bgcolor='#FFFFFF' class='style9'><strong>Ticket No</strong></td><td width='133' bgcolor='#FFFFFF' class='style9'><strong>Response Delay(mins)</strong></td></tr>"
            Dim CircuitCount As Integer = 0
            Dim msg As MailMessage
            msg = New MailMessage
            Smtp_Server = New SmtpClient

            Dim para As String(,) = {{"@DepartmentID", DepartmentID}}
            Dim para1 As String(,)
            Dim para2 As String(,)
            Dim para3 As String(,)
            Dim Depart As String

            Dim dt = objBSS.SP_Datatable("sp_Esc_GetComplainDetails", para)

            If DepartmentID = 3 Then
                Depart = "IP-NOC"
            ElseIf DepartmentID = 15 Then
                Depart = "OTS"
            ElseIf DepartmentID = 59 Then
                Depart = "O&M South"
            ElseIf DepartmentID = 58 Then
                Depart = "O&M Central"
            ElseIf DepartmentID = 60 Then
                Depart = "O&M North"
            End If

            If dt.Rows.Count > 1 Then
                For i As Integer = 0 To dt.Rows.Count - 1
                    str = str + "<tr><td width='104' height='22' bgcolor='#FFFFFF' class='style9'>" + Convert.ToString(dt.Rows(i)("SignupID")) + "</td><td width='204' bgcolor='#FFFFFF' class='style9'>" + dt.Rows(i)("CircuitName") + "</td><td width='114' bgcolor='#FFFFFF' class='style9'>" + dt.Rows(i)("TicketNo") + "</td><td width='114' bgcolor='#FFFFFF' class='style9'>" + Convert.ToString(dt.Rows(i)("Delay in Minutes")) + " mins</td></tr>"
                    CircuitCount = CircuitCount + 1
                Next

                para1 = {{"@Stage", "ComplainEsclation"}}
                para2 = {{"@DepartmentID ", DepartmentID}, {"@Stage", "ComplainEsclation"}}
                para3 = {{"@DepartmentID ", DepartmentID}, {"@Stage", "ComplainEsclation"}}

                Dim dt_From = objBSS.SP_Datatable("GetEmailFrom", para1)
                Dim dt_To = objBSS.SP_Datatable("sp_OTS_GetEmailTO", para2)
                Dim dt_CC = objBSS.SP_Datatable("sp_GetComplainEmailCC", para3)

                msg.From = New MailAddress(dt_From.Rows(0)("FromID"), dt_From.Rows(0)("Name"))

                If dt_To.Rows.Count > 0 Then
                    For i As Integer = 0 To dt_To.Rows.Count - 1
                        msg.To.Add(dt_To.Rows(i)(0).ToString)
                    Next
                Else
                    msg.To.Add(dt_From.Rows(0)("FromID"))
                End If

                If dt_CC.Rows.Count > 0 Then
                    For i As Integer = 0 To dt_CC.Rows.Count - 1
                        msg.CC.Add(dt_CC.Rows(i)(0).ToString)
                    Next
                Else
                    msg.CC.Add(dt_From.Rows(0)("FromID"))
                End If

                msg.Subject = "Escalation Alert: Un-Resolve Complains"
                msg.Body = "<html><head><style type='text/css'>.style1 {color: #FFFFFF}.style9 {font-family: Calibri; font-size: 12px; }.style10 {font-size: 10px}.style12 {font-family: Calibri}.style15 {font-family: Calibri; font-size: 12px; font-weight: bold; }</style></head><body><h3 align='center' class='style9'> Circuit Complain  Summary </h3><table width='436' height='84' border='2' align='center' bordercolor='#333333' class='style10'><tr><td width='149' height='25'  class='style9'><p class='style9'> Complain Count : </p></td><td width='230' class='style9'><p>" & Convert.ToString(dt.Rows.Count) & "</p></td></tr><tr><td height='25' class='style9'><p class='style9'> Department : </p></td><td class='style9'><p >" & Depart & "</p></td></tr><tr><td height='22' class='style9'><p class='style9'> Current Status : </p></td><td class='style9'><p>  In Process </p></td></tr></table><h3 align='center' class='style9'>List of Pending Ticket</h3><table width='446' height='30' border='1' align='center' bordercolor='#333333' class='style10'>" & str & "</table><p>&nbsp;</p><h3 align='center' class='style9'>  <span class='style12'>*This is an autogenerated email from BSS, incase of any query please contact BSS Adminstrator</span></h3><p class='style9'><br><br></p><p class='style9'>&nbsp; </p></body></html>"
                WriteLoggedDB("MpplService", "BSS_EsclationCall Data Fetch & Body", "Unknown")
                msg.IsBodyHtml = True
                Smtp_Server.Host = "202.142.160.14"
                Smtp_Server.Send(msg)
                Return True
                WriteLoggedDB("MpplService", "BSS_EsclationCall Email Done", "Unknown")
            Else
                Return False
                WriteLoggedDB("MpplService", "BSS_EsclationCall Email Failed", "Unknown")
            End If

        Catch ex As Exception
            WriteLoggedDB("MpplService", "BSS_EsclationCall Exception Failed", "Unknown")
            Return False
        Finally

        End Try
    End Function

    Public Function EsclationEmail_ICS(ByVal Flag As String) As String
        Try
            Dim str_All As String = "<tr><td width='85' height='22' bgcolor='#FFFFFF' class='style9'><strong>Serial No</strong></td><td width='85' height='22' bgcolor='#FFFFFF' class='style9'><strong>SignupID</strong></td><td width='147' bgcolor='#FFFFFF' class='style9'><strong>Partner Name</strong></td><td width='147' bgcolor='#FFFFFF' class='style9'><strong>Circuit Name</strong></td><td width='110' bgcolor='#FFFFFF' class='style9'><strong>Ticket No</strong></td><td width='133' bgcolor='#FFFFFF' class='style9'><strong>Complaint Status</strong></td><td width='133' bgcolor='#FFFFFF' class='style9'><strong>FaultOccuredDateTime</strong></td><td width='133' bgcolor='#FFFFFF' class='style9'><strong>Response Delay(Days)</strong></td></tr>"
            Dim str1_ETTR As String = "<tr><td width='85' height='22' bgcolor='#FFFFFF' class='style9'><strong>SignupID</strong></td><td width='147' bgcolor='#FFFFFF' class='style9'><strong>Circuit Name</strong></td><td width='110' bgcolor='#FFFFFF' class='style9'><strong>Ticket No</strong></td><td width='133' bgcolor='#FFFFFF' class='style9'><strong>Complaint Status</strong></td><td width='133' bgcolor='#FFFFFF' class='style9'><strong>FaultOccuredDateTime</strong></td><td width='133' bgcolor='#FFFFFF' class='style9'><strong>ETTR DateTime</strong></td></tr>"
            Dim str1_RFO As String = "<tr><td width='85' height='22' bgcolor='#FFFFFF' class='style9'><strong>SignupID</strong></td><td width='147' bgcolor='#FFFFFF' class='style9'><strong>Circuit Name</strong></td><td width='110' bgcolor='#FFFFFF' class='style9'><strong>Ticket No</strong></td><td width='133' bgcolor='#FFFFFF' class='style9'><strong>Complaint Status</strong></td><td width='133' bgcolor='#FFFFFF' class='style9'><strong>FaultClearedDateTime</strong></td><td width='133' bgcolor='#FFFFFF' class='style9'><strong>Response Delay(Minutes)</strong></td></tr>"


            Dim InOpenCount As Integer = 0
            Dim ResolvedCount As Integer = 0
            Dim msg As MailMessage
            msg = New MailMessage
            Smtp_Server = New SmtpClient

            Dim para As String(,) = {{"@Flag", Flag}}
            Dim para1 As String(,)
            Dim para2 As String(,)

            Dim dt = objBSS.SP_Datatable("sp_Esc_GetICSComplainAlerts", para)

            If dt.Rows.Count > 0 Then
                If Flag = "PendingAll" Or Flag = "OMS" Or Flag = "OMC" Or Flag = "OMN" Then
                    For i As Integer = 0 To dt.Rows.Count - 1
                        str_All = str_All + "<tr><td width='85' height='22' bgcolor='#FFFFFF' class='style9'>" + Convert.ToString(i + 1) + "</td><td width='85' height='22' bgcolor='#FFFFFF' class='style9'>" + Convert.ToString(dt.Rows(i)("SignupID")) + "</td><td width='204' bgcolor='#FFFFFF' class='style9'>" + dt.Rows(i)("PartnerName") + "</td><td width='204' bgcolor='#FFFFFF' class='style9'>" + dt.Rows(i)("CircuitName") + "</td><td width='114' bgcolor='#FFFFFF' class='style9'>" + dt.Rows(i)("TicketNo") + "</td><td width='114' bgcolor='#FFFFFF' class='style9'>" + dt.Rows(i)("ComplainStatus") + "</td><td width='114' bgcolor='#FFFFFF' class='style9'>" + dt.Rows(i)("FaultOccuredDateTime") + "</td><td width='114' bgcolor='#FFFFFF' class='style9'>" + Convert.ToString(dt.Rows(i)("Delay in Days")) + " Days</td></tr>"
                        If dt.Rows(i)("ComplainStatus") = "Open" Or dt.Rows(i)("ComplainStatus") = "InProcess" Then
                            InOpenCount = InOpenCount + 1
                        ElseIf dt.Rows(i)("ComplainStatus") = "Resolved" Then
                            ResolvedCount = ResolvedCount + 1
                        End If

                    Next
                    msg.Subject = "GSAC PULSE: Pending Complaints in BSS"
                    msg.Body = "<html><head><style type='text/css'>.style1 {color: #FFFFFF}.style9 {font-family: Calibri; font-size: 12px; }.style10 {font-size: 10px}.style12 {font-family: Calibri}.style15 {font-family: Calibri; font-size: 12px; font-weight: bold; }</style></head><body><h3 align='center' class='style9'> Circuit Complain  Summary </h3><table width='408' height='85' border='2' align='center' bordercolor='#333333' class='style10'><tr><td width='225' height='25'  class='style9'><p class='style9'> Total Complaint : </p></td><td width='276' class='style9'><p>" & Convert.ToString(dt.Rows.Count) & "</p></td></tr><td width='225' height='25'  class='style9'><p class='style9'> InProcess/Open Complaints : </p></td><td width='276' class='style9'><p>" & Convert.ToString(InOpenCount) & "</p></td></tr><td width='225' height='25'  class='style9'><p class='style9'> Resolved Complaints: </p></td><td width='276' class='style9'><p>" & Convert.ToString(ResolvedCount) & "</p></td></tr></table><h3 align='center' class='style9'>List of Tickets</h3><table width='680' height='1028' border='1' align='center' bordercolor='#333333' class='style10'>" & str_All & "</table><p>&nbsp;</p><h3 align='center' class='style9'>  <span class='style12'>*This is an autogenerated email from BSS, incase of any query please contact BSS Administrator</span></h3><p class='style9'><br><br></p><p class='style9'>&nbsp; </p></body></html>"

                ElseIf Flag = "ETTR" Then
                    For i As Integer = 0 To dt.Rows.Count - 1
                        str1_ETTR = str1_ETTR + "<tr><td width='104' height='22' bgcolor='#FFFFFF' class='style9'>" + Convert.ToString(dt.Rows(i)("SignupID")) + "</td><td width='204' bgcolor='#FFFFFF' class='style9'>" + dt.Rows(i)("CircuitName") + "</td><td width='114' bgcolor='#FFFFFF' class='style9'>" + dt.Rows(i)("TicketNo") + "</td><td width='114' bgcolor='#FFFFFF' class='style9'>" + dt.Rows(i)("ComplainStatus") + "</td><td width='114' bgcolor='#FFFFFF' class='style9'>" + dt.Rows(i)("FaultOccuredDateTime") + "</td><td width='114' bgcolor='#FFFFFF' class='style9'>" + dt.Rows(i)("ETRR_DateTime") + "</td></tr>"

                    Next
                    msg.Subject = "Alert: ETTR Expiry"
                    msg.Body = "<html><head><style type='text/css'>.style1 {color: #FFFFFF}.style9 {font-family: Calibri; font-size: 12px; }.style10 {font-size: 10px}.style12 {font-family: Calibri}.style15 {font-family: Calibri; font-size: 12px; font-weight: bold; }</style></head><body><h3 align='center' class='style9'> Circuit Complain  Summary </h3><table width='408' height='85' border='2' align='center' bordercolor='#333333' class='style10'><tr><td width='149' height='25'  class='style9'><p class='style9'> Complain Count : </p></td><td width='230' class='style9'><p>" & Convert.ToString(dt.Rows.Count) & "</p></td></tr></table><h3 align='center' class='style9'>List of Ticket's</h3><table width='500' height='50' border='1' align='center' bordercolor='#333333' class='style10'>" & str1_ETTR & "</table><p>&nbsp;</p><h3 align='center' class='style9'>  <span class='style12'>*This is an autogenerated email from BSS, incase of any query please contact BSS Administrator</span></h3><p class='style9'><br><br></p><p class='style9'>&nbsp; </p></body></html>"

                ElseIf Flag = "RFO" Then
                    For i As Integer = 0 To dt.Rows.Count - 1
                        str1_RFO = str1_RFO + "<tr><td width='104' height='22' bgcolor='#FFFFFF' class='style9'>" + Convert.ToString(dt.Rows(i)("SignupID")) + "</td><td width='204' bgcolor='#FFFFFF' class='style9'>" + dt.Rows(i)("CircuitName") + "</td><td width='114' bgcolor='#FFFFFF' class='style9'>" + dt.Rows(i)("TicketNo") + "</td><td width='114' bgcolor='#FFFFFF' class='style9'>" + dt.Rows(i)("ComplainStatus") + "</td><td width='114' bgcolor='#FFFFFF' class='style9'>" + dt.Rows(i)("FaultClearedDateTime") + "</td><td width='114' bgcolor='#FFFFFF' class='style9'>" + dt.Rows(i)("Delay in Minutes") + "</td></tr>"
                    Next
                    msg.Subject = "Alert: Pending RFO "
                    msg.Body = "<html><head><style type='text/css'>.style1 {color: #FFFFFF}.style9 {font-family: Calibri; font-size: 12px; }.style10 {font-size: 10px}.style12 {font-family: Calibri}.style15 {font-family: Calibri; font-size: 12px; font-weight: bold; }</style></head><body><h3 align='center' class='style9'>List of Ticket's those RFO are Pending</h3><table width='446' height='30' border='1' align='center' bordercolor='#333333' class='style10'>" & str1_RFO & "</table><p>&nbsp;</p><h3 align='center' class='style9'>  <span class='style12'>*This is an autogenerated email from BSS, incase of any query please contact BSS Administrator</span></h3><p class='style9'><br><br></p><p class='style9'>&nbsp; </p></body></html>"



                End If


                para1 = {{"@Stage", "ICS_ComplainEsclation"}}

                If Flag = "ETTR" Or Flag = "RFO" Then
                    para2 = {{"@DepartmentID", 0}, {"@Stage", "ICS_ComplainEsclation_Internal"}}

                ElseIf Flag = "PendingAll" Then
                    para2 = {{"@DepartmentID", 0}, {"@Stage", "ICS_ComplainEsclation"}}

                ElseIf Flag = "OMS" Then
                    para2 = {{"@DepartmentID", 59}, {"@Stage", "ICS_ComplainEsclation"}}

                ElseIf Flag = "OMN" Then
                    para2 = {{"@DepartmentID", 60}, {"@Stage", "ICS_ComplainEsclation"}}

                ElseIf Flag = "OMC" Then
                    para2 = {{"@DepartmentID", 58}, {"@Stage", "ICS_ComplainEsclation"}}

                End If



                Dim dt_From = objBSS.SP_Datatable("GetEmailFrom", para1)
                Dim dt_To = objBSS.SP_Datatable("sp_OTS_GetEmailTO", para2)
                Dim dt_CC = objBSS.SP_Datatable("sp_GetComplainEmailCC", para2)

                msg.From = New MailAddress(dt_From.Rows(0)("FromID"), dt_From.Rows(0)("Name"))

                If dt_To.Rows.Count > 0 Then
                    For i As Integer = 0 To dt_To.Rows.Count - 1
                        msg.To.Add(dt_To.Rows(i)(0).ToString)
                    Next
                Else
                    msg.To.Add(dt_From.Rows(0)("FromID"))
                End If

                If dt_CC.Rows.Count > 0 Then
                    For i As Integer = 0 To dt_CC.Rows.Count - 1
                        msg.CC.Add(dt_CC.Rows(i)(0).ToString)
                    Next
                Else
                    msg.CC.Add(dt_From.Rows(0)("FromID"))
                End If

                WriteLoggedDB("MpplService", "ICS_ComplainEsclation Data Fetch & Body", "Unknown")
                msg.IsBodyHtml = True
                Smtp_Server.Host = "202.142.160.14"
                Smtp_Server.Send(msg)
                Return "Record Found and email Done"
            Else
                WriteLoggedDB("MpplService", "BSS_EsclationCall No Result Found", "Unknown")
                Return "No-Record"
            End If

        Catch ex As Exception
            WriteLoggedDB("MpplService", "BSS_EsclationCall Exception Failed", "Unknown")
            Return "Exception"
        Finally

        End Try
    End Function



End Class