﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.42000
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On


Namespace My
    
    <Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "12.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
    Partial Friend NotInheritable Class MySettings
        Inherits Global.System.Configuration.ApplicationSettingsBase
        
        Private Shared defaultInstance As MySettings = CType(Global.System.Configuration.ApplicationSettingsBase.Synchronized(New MySettings()),MySettings)
        
#Region "My.Settings Auto-Save Functionality"
#If _MyType = "WindowsForms" Then
    Private Shared addedHandler As Boolean

    Private Shared addedHandlerLockObject As New Object

    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(), Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)> _
    Private Shared Sub AutoSaveSettings(ByVal sender As Global.System.Object, ByVal e As Global.System.EventArgs)
        If My.Application.SaveMySettingsOnExit Then
            My.Settings.Save()
        End If
    End Sub
#End If
#End Region
        
        Public Shared ReadOnly Property [Default]() As MySettings
            Get
                
#If _MyType = "WindowsForms" Then
               If Not addedHandler Then
                    SyncLock addedHandlerLockObject
                        If Not addedHandler Then
                            AddHandler My.Application.Shutdown, AddressOf AutoSaveSettings
                            addedHandler = True
                        End If
                    End SyncLock
                End If
#End If
                Return defaultInstance
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Password=P@kistan;Persist Security Info=True;User ID=sa;Initial Catalog=MPPLCircu"& _ 
            "its;Data Source=103.31.80.118\APPDB")>  _
        Public ReadOnly Property conMidd() As String
            Get
                Return CType(Me("conMidd"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Password=Multi!@#$%;Persist Security Info=True;User ID=apps.admin;Initial Catalog"& _ 
            "=BSS;Data Source=172.30.1.127")>  _
        Public ReadOnly Property conBSS() As String
            Get
                Return CType(Me("conBSS"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Password=Multi!@#$%;Persist Security Info=True;User ID=apps.admin;Initial Catalog"& _ 
            "=Rainmaker;Data Source=172.30.1.127")>  _
        Public ReadOnly Property conRM() As String
            Get
                Return CType(Me("conRM"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Password=MsGP@197;Persist Security Info=True;User ID=MsGP;Initial Catalog=AutoEma"& _ 
            "il;Data Source=172.30.1.127")>  _
        Public ReadOnly Property Email() As String
            Get
                Return CType(Me("Email"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("<table>    <td width=""718"" background=""file:///D|/Liaqat Data D/Project Backup/MP"& _ 
            "PLBillingProjectFinal/EmailFormat/download.jpg""><p align=""right"" >&nbsp;</p>    "& _ 
            "</td> </tr>  <tr>    <td style =""font-family: calibri; font-size: 14px;"">  Dear "& _ 
            "Team,  </tr>  <tr>    <td><p style =""font-family: calibri; font-size: 14px;"">Fro"& _ 
            "m Web Portal  email   sent  by System, <strong>Follow-up  Lead, Email Alert</str"& _ 
            "ong> Details given bellow. <br />        Date: !! </p>      <table width=""444"" b"& _ 
            "order=""1"">        <tr>          <td><p align=""center"" style =""font-family: calib"& _ 
            "ri; font-size: 14px;""><strong>Lead Details </strong></p></td>          <td><p st"& _ 
            "yle=""font-family: calibri; font-size: 14px;""><strong>Transaction Information </s"& _ 
            "trong></p>          </td>        </tr>        <tr>          <td width=""195""><p s"& _ 
            "tyle =""font-family: calibri; font-size: 14px;"">Lead Name </p></td>          <td "& _ 
            "width=""233""><p style =""font-family: calibri; font-size: 14px;"">!!</td>        </"& _ 
            "tr>        <tr>          <td><p style =""font-family: calibri; font-size: 14px;"">"& _ 
            "Description</p></td>          <td><p style =""font-family: calibri; font-size: 14"& _ 
            "px;"">!!</td>        </tr>        <tr>          <td><p style =""font-family: calib"& _ 
            "ri; font-size: 14px;"">Primary Contact Name </p></td>          <td><p style =""fon"& _ 
            "t-family: calibri; font-size: 14px;"">!!</td>        </tr>        <tr>          <"& _ 
            "td><p style =""font-family: calibri; font-size: 14px;"">Contact Email Addres</p> <"& _ 
            "/td>          <td><p style =""font-family: calibri; font-size: 14px;"">!!</td>    "& _ 
            "    </tr>        <tr>          <td><p style =""font-family: calibri; font-size: 1"& _ 
            "4px;"">Transaction Date </p> </td>          <td><p style =""font-family: calibri; "& _ 
            "font-size: 14px;"">!!</td>        </tr>      </table>      </td>  </tr>  <tr>    "& _ 
            "<td><span class=""style35 style1"">!!</span></td>  </tr>  <tr>    <td style =""font"& _ 
            "-family: calibri; font-size: 14px; font-weight: bold; font-style: italic;"" >Mult"& _ 
            "inet  Pakistan Private Limited</td>  </tr>  <tr>    <td style =""font-family: cal"& _ 
            "ibri; font-size: 14px; font-style: italic;""><span class=""style55"">An Axiata Grou"& _ 
            "p  Subsidiary </span></td>  </tr>  <tr>    <td style =""font-family: calibri; fon"& _ 
            "t-size: 14px; font-style: italic;""><span class=""style55"">Address: D-17, Sector 3"& _ 
            "0,  Korangi Industrial Area, Karachi.</span></td>  </tr>  <tr>    <td style =""fo"& _ 
            "nt-family: calibri; font-size: 14px; font-style: italic;""><em><strong>URL: <a hr"& _ 
            "ef=""http://www.multi.net.pk/"" title=""http://www.multi.net.pk/"">www.multi.net.pk<"& _ 
            "/a></strong></em></td>  </tr>  <tr>    <td style =""font-family: calibri; font-si"& _ 
            "ze: 14px; font-style: italic;""><div align=""left"" class=""style52""></div></td>  </"& _ 
            "tr>   <tr>    <td style =""font-family: calibri; font-size: 14px;""><div align=""ce"& _ 
            "nter"" class=""style52""><em><strong>*** This is an automatically generated email &"& _ 
            "ndash; please do not reply to  it. ***</strong></em></div></td>  </tr></table>")>  _
        Public ReadOnly Property EmailFormat() As String
            Get
                Return CType(Me("EmailFormat"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Password=MsGP@197;Persist Security Info=True;User ID=MsGP;Initial Catalog=MLTNT;D"& _ 
            "ata Source=172.30.1.127")>  _
        Public ReadOnly Property conGP() As String
            Get
                Return CType(Me("conGP"),String)
            End Get
        End Property
    End Class
End Namespace

Namespace My
    
    <Global.Microsoft.VisualBasic.HideModuleNameAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute()>  _
    Friend Module MySettingsProperty
        
        <Global.System.ComponentModel.Design.HelpKeywordAttribute("My.Settings")>  _
        Friend ReadOnly Property Settings() As Global.WebService.My.MySettings
            Get
                Return Global.WebService.My.MySettings.Default
            End Get
        End Property
    End Module
End Namespace