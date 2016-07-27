'**************************************************************************************************
' 0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056
'**************************************************************************************************

Option Strict Off
Option Explicit On 

Public Class TWM_Maxhill
    Private WithEvents SBO_Application As SAPbouiCOM.Application

    '//**********************************************************
    '// declaring an Event filters container object and an
    '// event filter object
    '//**********************************************************
    Public gcAddOnName As String = "TWM_Maxhill"

    Private oForm As SAPbouiCOM.Form
    Public oFilters As SAPbouiCOM.EventFilters
    Public oFilter As SAPbouiCOM.EventFilter

    Public DocEntry As Long
    Public DocNum As String

    Private oCompany As SAPbobsCOM.Company
    Private vCardCode As String = ""
    Private kSQL As String = ""
    Private BaseFormName As String = ""
    Private i As Integer = 0

    Private CompanyA As Boolean
    Private ARMarkUp As Double

#Region "Application initilization and connecting to current running B1"


    Private Sub SetFilters()
        oFilters = New SAPbouiCOM.EventFilters

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_RESIZE)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_CLOSE)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)
        oFilter.AddEx("133") 'AR Inv
        oFilter.AddEx("141") 'AP Inv
        oFilter.AddEx("142") 'PO
        oFilter.AddEx("181") 'APCN
        oFilter.AddEx("3002") 'Draft Report
        oFilter.AddEx("twmICCRE")   'IC credential form
        SBO_Application.SetFilter(oFilters)
    End Sub

    ' set application
    Private Sub SetApplication()

        Dim SboGuiApi As SAPbouiCOM.SboGuiApi = New SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String = Environment.GetCommandLineArgs.GetValue(1)
        SboGuiApi.Connect(sConnectionString)
        SBO_Application = SboGuiApi.GetApplication()


    End Sub
    Public Function IsDesignTime() As Boolean
        Try
            Dim cProcess As System.Diagnostics.Process = System.Diagnostics.Process.GetCurrentProcess()
            If Right(cProcess.ProcessName, 6) = "vshost" Then 'This part is for RUNTIME Mode
                Return True
            Else
                Return False
                'xApplication.SetStatusBarMessage(cProcess.Id & " 2N " & cProcess.ProcessName & " << " & cParent.Id & " N " & cParent.ProcessName, SAPbouiCOM.BoMessageTime.bmt_Short, False)
            End If
        Catch ex As Exception
            Return False
        End Try

    End Function
    'Modified by Edy on 20130207. Requested by KK to include the licensing mechanism
    Public Sub CheckLicenceEX()
        Try
            Dim oLic As New TWM_Licence.TWM_Licence(SBO_Application, oCompany, gcAddOnName, Key)
            If Not oLic.IsValid Then
                SBO_Application.StatusBar.SetText("Could not start addon " & gcAddOnName & ". " & oLic.LastErrorDescription)
                oCompany.Disconnect()
                End
            Else
                If oLic.DaysToExpiry < 10 Then
                    If oLic.DaysToExpiry > 0 Then
                        SBO_Application.MessageBox("Your add on " & gcAddOnName & " will expire in " & oLic.DaysToExpiry & " days. Please contact support for the license.")
                    Else
                        SBO_Application.MessageBox("Your add on " & gcAddOnName & " expires today. Please contact support for the license.")
                    End If

                End If
            End If

        Catch ex As Exception
            SBO_Application.StatusBar.SetText("Unable to start addon " & gcAddOnName & ". No licence found.")
            SBO_Application.Disconnect()
            System.Windows.Forms.Application.Exit()
        End Try
    End Sub
    Public Sub New()

        SetApplication()
        SetFilters()

        oCompany = SBO_Application.Company.GetDICompany
        If oCompany.Connected = True Then
            SBO_Application.SetStatusBarMessage("TWM_Maxhill Addon is Connected", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        Else
            SBO_Application.MessageBox("Failed connecting to company DB")
            End ' Terminating the Add-On Application
        End If
        '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
        'CLog("Connected, now checking for Addon License")
        If Not IsDesignTime() Then
            CheckLicenceEX()
        End If
        '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

        SBO_Application.SetStatusBarMessage("Checking Objects Starting .. ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        If TWM_CheckObjects(oCompany) = False Then
            SBO_Application.SetStatusBarMessage("Object Creation Done", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            SBO_Application.MessageBox("Addon Application Created Several object(s), You need to restart SAP. Addon is terminating!")
            End
        Else
            SBO_Application.SetStatusBarMessage("Object Checking Done", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        End If


        'Create Authorization for form twmICCRE
        Dim oAuth As SAPbobsCOM.UserPermissionTree = Nothing
        Try
            oAuth = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserPermissionTree)
            If Not oAuth.GetByKey(gcAddOnName) Then
                oAuth.PermissionID = gcAddOnName
                oAuth.Name = "TWM Maxhill New"
                oAuth.Options = SAPbobsCOM.BoUPTOptions.bou_FullReadNone
                Dim lErr As Integer = oAuth.Add
                If lErr <> 0 Then Throw New Exception(oCompany.GetLastErrorDescription)

                oAuth = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserPermissionTree)
                oAuth.PermissionID = "twmICCRE"
                oAuth.Name = "IC Credential"
                oAuth.ParentID = gcAddOnName
                oAuth.Options = SAPbobsCOM.BoUPTOptions.bou_FullReadNone
                oAuth.UserPermissionForms.FormType = "twmICCRE"
                lErr = oAuth.Add
                If lErr <> 0 Then Throw New Exception(oCompany.GetLastErrorDescription)
            End If

        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

        'Create Menus
        If SBO_Application.Menus.Exists(gcAddOnName) Then SBO_Application.Menus.RemoveEx(gcAddOnName)
        SBO_Application.Menus.Item("43520").SubMenus.Add(gcAddOnName, "TWM Maxhill New", SAPbouiCOM.BoMenuType.mt_POPUP, 99)
        SBO_Application.Menus.Item(gcAddOnName).SubMenus.Add("twmICCRE", "Interco Credential", SAPbouiCOM.BoMenuType.mt_STRING, 0)
        Try
            If System.IO.File.Exists("Logo.jpg") Then
                SBO_Application.Menus.Item(gcAddOnName).Image = My.Application.Info.DirectoryPath & "\Logo.JPG"
            End If

        Catch ex As Exception
        End Try



        CompanyA = SystemACheck()
    End Sub
#End Region


#Region "sbo_event"
    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBO_Application.AppEvent

        Select Case EventType
            Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                If SBO_Application.Menus.Exists(gcAddOnName) Then SBO_Application.Menus.RemoveEx(gcAddOnName)
                System.Windows.Forms.Application.Exit()
            Case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition
                If SBO_Application.Menus.Exists(gcAddOnName) Then SBO_Application.Menus.RemoveEx(gcAddOnName)
                System.Windows.Forms.Application.Exit()
            Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown
                If SBO_Application.Menus.Exists(gcAddOnName) Then SBO_Application.Menus.RemoveEx(gcAddOnName)
                System.Windows.Forms.Application.Exit()
        End Select
    End Sub


    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Try
            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID
                    Case "twmICCRE"
                        DrawForm(pVal.MenuUID)
                    Case Else
                        'Diagnostics.Debug.WriteLine(pVal.MenuUID.ToString)
                End Select
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
        
    End Sub

    Private Sub kSetLineNum(ByVal kForm As SAPbouiCOM.Form, Optional ByVal kMode As SAPbouiCOM.BoFormMode = -1)
        Dim i As Integer
        For i = 1 To kForm.DataSources.DBDataSources.Item(1).Size
            kForm.DataSources.DBDataSources.Item(1).SetValue("LineID", i - 1, i)
        Next
        kForm.Update()
        If kMode <> -1 Then kForm.Mode = kMode
    End Sub
    Private Function kGetItemNamebyCode(ByVal iCode As String) As String
        Dim kRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim buffer As String = ""
        kRS.DoQuery("select ItemName from OITM where ItemCode='" & iCode & "'")
        kRS.MoveFirst()
        If Not kRS.EoF Then buffer = kRS.Fields.Item(0).Value
        Return buffer
    End Function


    Private Sub kGetNextNumber(ByVal kForm As SAPbouiCOM.Form)

        Dim oText As SAPbouiCOM.EditText = Nothing
        Dim kRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Select Case kForm.UniqueID
            Case "TWM_MUOMPL"
                oText = kForm.Items.Item("txtDocNum").Specific
                kSQL = "select coalesce(max(docentry),0) from [@TWM_UOMPL]"
            Case "TWM_MUOMOSPP"
                oText = kForm.Items.Item("txtDocNum").Specific
                kSQL = "select coalesce(max(docentry),0) from [@TWM_OSPPH]"
            Case "TWM_MUOMSPP1"
                oText = kForm.Items.Item("txtDocNum").Specific
                kSQL = "select coalesce(max(docentry),0) from [@TWM_SPP1H]"
            Case "TWM_MUOMSPP2"
                oText = kForm.Items.Item("txtDocNum").Specific
                kSQL = "select coalesce(max(docentry),0) from [@TWM_SPP2H]"
        End Select
        kRS.DoQuery(kSQL)
        oText.Value = kRS.Fields.Item(0).Value + 1
    End Sub
    Private Function UDFCheck(ByVal kFields As SAPbobsCOM.Fields, ByVal FieldName As String) As Boolean
        Dim kField As SAPbobsCOM.Field
        Dim kFlag As Boolean = False
        For Each kField In kFields
            If kField.Name = FieldName Then
                kFlag = True
            End If
        Next
        Return kFlag
    End Function

    Private Function getVATGroup(ByVal oVAT As String) As String
        Dim kRet As String = ""
        Dim kRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim kSQL As String = "SELECT Code from OVTG where Code like '" & oVAT & "'"
        kRS.DoQuery(kSQL)
        If Not kRS.EoF Then
            kRet = kRS.Fields.Item(0).Value.ToString
        End If
        Return kRet
    End Function
    Private Sub errSub(ByVal err As Long)
        Dim errmsg As String = ""
        If err <> 0 Then
            If err <> -2035 Then
                vCmp.GetLastError(err, errmsg)
                MsgBox(errmsg)
                End
            End If
        End If
    End Sub
    Private Function ocCheckTable(ByVal tableName As String, ByVal tableDesc As String, ByVal tableType As SAPbobsCOM.BoUTBTableType) As Boolean
        Dim xRet As Boolean = False
        Dim oUserTableMD As SAPbobsCOM.UserTablesMD = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
        With oUserTableMD
            If Not (.GetByKey(tableName)) Then
                .TableName = tableName
                .TableDescription = tableDesc
                .TableType = tableType
                Dim err As Long = .Add()
                errSub(err)
            Else
                xRet = True
            End If
        End With
        oUserTableMD = Nothing
        GC.Collect()
        Return xRet
    End Function

    Private Function ocGetFieldID(ByVal tableName As String, ByVal fieldName As String) As Long
        Dim cIndex As Long = -1
        Dim xRS As SAPbobsCOM.Recordset = Nothing
        xRS = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim qStr As String = "select FieldID from CUFD where TableID = '" & tableName & "' and AliasID= '" & fieldName & "'"
        xRS.DoQuery(qStr)
        xRS.MoveFirst()
        If Not xRS.EoF Then
            cIndex = xRS.Fields.Item(0).Value
        Else
            cIndex = -1
        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xRS)
        xRS = Nothing
        GC.Collect()
        Return cIndex
    End Function

    Private Function ocICGetFieldID(ByRef xCompany As SAPbobsCOM.Company, ByVal tableName As String, ByVal fieldName As String) As Long
        Dim cIndex As Long = -1
        Dim xRS As SAPbobsCOM.Recordset = Nothing
        xRS = xCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim qStr As String = "select FieldID from CUFD where TableID = '" & tableName & "' and AliasID= '" & fieldName & "'"
        xRS.DoQuery(qStr)
        xRS.MoveFirst()
        If Not xRS.EoF Then
            cIndex = xRS.Fields.Item(0).Value
        Else
            cIndex = -1
        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xRS)
        xRS = Nothing
        GC.Collect()
        Return cIndex
    End Function
    Private Function ocCheckField(ByVal tableName As String, ByVal fieldName As String, ByVal fieldDesc As String, ByVal fieldType As SAPbobsCOM.BoFieldTypes, Optional ByVal fieldSubType As SAPbobsCOM.BoFldSubTypes = SAPbobsCOM.BoFldSubTypes.st_None, Optional ByVal fieldSize As Integer = 0, Optional ByVal fieldLink As String = "", Optional ByVal kValidValue As String = "", Optional ByVal kDefault As String = "") As Boolean
        Dim xRet As Boolean = False
        Dim fID As Long = ocGetFieldID(tableName, fieldName)
        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        With oUserFieldsMD
            If fID = -1 Then 'If Not (.GetByKey(tableName, fieldName)) Then
                .TableName = tableName
                .Name = fieldName
                .Description = fieldDesc
                .Type = fieldType
                If fieldSubType <> SAPbobsCOM.BoFldSubTypes.st_None Then
                    .SubType = fieldSubType
                End If
                ''.db_Alpha db_Numeric Must set EditSize
                'If fieldType = SAPbobsCOM.BoFieldTypes.db_Alpha Or fieldType = SAPbobsCOM.BoFieldTypes.db_Numeric Then
                '    .EditSize = fieldSize
                'End If

                If fieldSize > 0 Then
                    .EditSize = fieldSize
                    .Size = fieldSize
                End If
                If fieldLink <> "" Then
                    .LinkedTable = fieldLink
                End If
                If kValidValue <> "" Then
                    Dim vValue() As String
                    Dim vItem As String
                    Dim vX As Integer = 0
                    vValue = kValidValue.Split(",")
                    For Each vItem In vValue
                        vX += 1
                        If vX Mod 2 = 0 Then
                            .ValidValues.Description = vItem
                            .ValidValues.Add()
                        Else
                            .ValidValues.Value = vItem
                        End If
                    Next
                    If kDefault <> "" Then
                        .DefaultValue = kDefault
                    End If
                End If
                Dim err As Long = .Add
                errSub(err)
            Else
                xRet = True
            End If
        End With
        oUserFieldsMD = Nothing
        GC.Collect()
        Return xRet
    End Function

    Private Function ocICCheckField(ByRef xCompany As SAPbobsCOM.Company, ByVal tableName As String, ByVal fieldName As String, ByVal fieldDesc As String, ByVal fieldType As SAPbobsCOM.BoFieldTypes, Optional ByVal fieldSubType As SAPbobsCOM.BoFldSubTypes = SAPbobsCOM.BoFldSubTypes.st_None, Optional ByVal fieldSize As Integer = 0, Optional ByVal fieldLink As String = "", Optional ByVal kValidValue As String = "", Optional ByVal kDefault As String = "") As Boolean
        Dim xRet As Boolean = False
        Dim fID As Long = ocICGetFieldID(xCompany, tableName, fieldName)
        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD = xCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        With oUserFieldsMD
            If fID = -1 Then 'If Not (.GetByKey(tableName, fieldName)) Then
                .TableName = tableName
                .Name = fieldName
                .Description = fieldDesc
                .Type = fieldType
                If fieldSubType <> SAPbobsCOM.BoFldSubTypes.st_None Then
                    .SubType = fieldSubType
                End If
                ''.db_Alpha db_Numeric Must set EditSize
                'If fieldType = SAPbobsCOM.BoFieldTypes.db_Alpha Or fieldType = SAPbobsCOM.BoFieldTypes.db_Numeric Then
                '    .EditSize = fieldSize
                'End If

                If fieldSize > 0 Then
                    .EditSize = fieldSize
                    .Size = fieldSize
                End If
                If fieldLink <> "" Then
                    .LinkedTable = fieldLink
                End If
                If kValidValue <> "" Then
                    Dim vValue() As String
                    Dim vItem As String
                    Dim vX As Integer = 0
                    vValue = kValidValue.Split(",")
                    For Each vItem In vValue
                        vX += 1
                        If vX Mod 2 = 0 Then
                            .ValidValues.Description = vItem
                            .ValidValues.Add()
                        Else
                            .ValidValues.Value = vItem
                        End If
                    Next
                    If kDefault <> "" Then
                        .DefaultValue = kDefault
                    End If
                End If
                Dim err As Long = .Add
                errSub(err)
            Else
                xRet = True
            End If
        End With
        oUserFieldsMD = Nothing
        GC.Collect()
        Return xRet
    End Function

    Private Function runQueryStr(ByVal kQuery As String) As String
        Dim kRS As SAPbobsCOM.Recordset
        kRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        kRS.DoQuery(kQuery)
        If Left(kQuery, 6) = "Insert" Then
            Return "0"
        Else
            Return kRS.Fields.Item(0).Value.ToString
        End If
    End Function
    Private Function SystemACheck() As Boolean 'pCode = ( HSBC/HSBCNet ) Customer ID
        Dim pCode As String = "MaxHillSysID"
        Dim SysID As String = ""
        Dim sysAFlag As Boolean
        Dim xQ As String = "Select count(*) from [@TWM_SS] Where [Name] Like '" & pCode & "'"
        If runQueryStr(xQ) = "0" Then
            Dim nRS As Integer = CInt(runQueryStr("Select Count(*) From [@TWM_SS]"))
            SysID = runQueryStr("Insert into [@TWM_SS] ([Code],[Name],[U_SValue]) Values ( " & nRS & ",'" & pCode & "','A')")
        End If
        SysID = runQueryStr("Select U_SValue from [@TWM_SS] where [Name] = '" & pCode & "'")
        sysAFlag = IIf(SysID = "A", True, False)
        Return sysAFlag
    End Function

    Private Function GetDefaultMarkup() As Double
        Dim pCode As String = "DefMarkup"
        Dim xsVal As String = ""
        Dim xVal As Double = 0
        Dim xQ As String = "Select count(*) from [@TWM_SS] Where [Name] Like '" & pCode & "'"
        If runQueryStr(xQ) = "0" Then
            Dim nRS As Integer = CInt(runQueryStr("Select Count(*) From [@TWM_SS]"))
            xsVal = runQueryStr("Insert into [@TWM_SS] ([Code],[Name],[U_SValue]) Values ( " & nRS & ",'" & pCode & "','4')")
        End If
        xsVal = runQueryStr("Select U_SValue from [@TWM_SS] where [Name] = '" & pCode & "'")
        xVal = (Val(xsVal) / 100) + 1
        Return xVal
    End Function

    Private Function GetMarkup(ByVal InvDate As Date) As Double
        Dim DateStr As String = InvDate.Year.ToString & Format(InvDate.Month, "00") & Format(InvDate.Day, "00")
        Dim xVal As Double = 0
        Dim xQ As String = "select U_FromDate, U_Markup from [@TWM_Markup] where U_FromDate <= '" & DateStr & "' order by U_FromDate desc"
        Dim xRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        xRS.DoQuery(xQ)
        If xRS.EoF Then
            xVal = GetDefaultMarkup()
        Else
            xVal = Val(xRS.Fields.Item(1).Value) / 100 + 1
        End If
        Return xVal
    End Function
    Public Function TWM_CheckObjects(ByVal xCmp As SAPbobsCOM.Company) As Boolean
        Dim kNoNeedRestart As Boolean = True
        GC.Collect()
        vCmp = xCmp
        Try
            Dim kRet As Boolean = True
            kRet = ocCheckTable("TWM_SS", "TWM System Settings", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            kRet = ocCheckField("@TWM_SS", "SValue", "Setting Value", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            '======================================================================================================================
            kRet = ocCheckTable("TWM_MARKUP", "TWM Markup Values", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            kRet = ocCheckField("@TWM_MARKUP", "FromDate", "From Date", SAPbobsCOM.BoFieldTypes.db_Date)
            kRet = ocCheckField("@TWM_MARKUP", "Markup", "Markup %", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage)
            '======================================================================================================================
            kRet = ocCheckField("OCRD", "TWMICDBN", "Intercompany Database Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50)
            kRet = ocCheckField("OCRD", "TWMICYN", "Intercompany Yes / No ?", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, , "0,No,1,Yes", "0")
            kRet = ocCheckField("OCRD", "TWMLCDC", "AR Default Customer ?", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, , "0,No,1,Yes", "0")
            '======================================================================================================================
            kRet = ocCheckField("OPOR", "TWMSCDBN", "Source DB Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50)
            kRet = ocCheckField("OPOR", "TWMBPR", "BP Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30)
            kRet = ocCheckField("OPOR", "TWMSDER", "Source DocEntry Ref", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11)
            kRet = ocCheckField("OPOR", "TWMSDNR", "Source DocNum Ref", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11)
            kRet = ocCheckField("OPOR", "TWMTDER", "Target DocEntry Ref", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11)
            kRet = ocCheckField("OPOR", "TWMTDNR", "Target DocNum Ref", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11)
            '======================================================================================================================
            kRet = ocCheckField("POR1", "TWMICBD", "IC Base Doc", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            kRet = ocCheckField("POR1", "TWMICBL", "IC Base Doc Line", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            kRet = ocCheckField("POR1", "TWMICBT", "IC Base Type", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            kRet = ocCheckField("POR1", "TWM_SS", "System Setting", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            '======================================================================================================================
            kRet = ocCheckField("PCH1", "TWM_BPCode", "BP Code for Target AR", SAPbobsCOM.BoFieldTypes.db_Alpha, , 15)
            kRet = ocCheckField("PCH1", "TWMTDER", "Target DocEntry Ref", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11)
            kRet = ocCheckField("PCH1", "TWMTDNR", "Target DocNum Ref", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11)
            '======================================================================================================================
            If kRet = False Then
                SBO_Application.SetStatusBarMessage("UDF for TWM_Addon are created!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                kNoNeedRestart = False
            End If

            If kNoNeedRestart = False Then
                SBO_Application.SetStatusBarMessage("Database Structure Changed, Please restart SAP and SAP Addon!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            End If
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace)
            Return False ' NeedRestart
        End Try
        Return kNoNeedRestart
    End Function

    Private Function getBP_ARINV() As String
        Dim BP As String = ""
        Dim sql As String = "Select Top 1 CardCode from OCRD where U_TWMLCDC = '1' and CardType= 'C'"
        Dim xRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        xRS.DoQuery(sql)
        If xRS.EoF Then
            BP = ""
        Else
            BP = xRS.Fields.Item(0).Value
        End If
        Return BP
    End Function

    Private Function CheckARINVdrf(ByVal APDE As String) As Boolean
        Dim sql As String = "Select * from ODRF Where ObjType = 13 and DocStatus = 'O' and U_TWMSDER=" & APDE
        Dim xRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        xRS.DoQuery(sql)
        Return xRS.EoF
    End Function
    Private Function CheckARINVdrf(ByVal APDE As String, ByVal APDN As String) As Boolean
        Dim sql As String = "Select * from ODRF Where ObjType = 13 and DocStatus = 'O' and U_TWMSDER=" & APDE
        sql &= " AND U_TWMSDNR=" & APDN
        Dim xRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        xRS.DoQuery(sql)
        Return xRS.EoF
    End Function

    Private Function CheckGoodIssue(ByVal APDE As String) As Boolean
        Dim hFlag As Boolean = False
        Dim dFlag As Boolean = False
        Dim sql As String = "Select * from OIGE Where DocStatus = 'O' and U_TWMSDER=" & APDE
        Dim xRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        xRS.DoQuery(sql)
        hFlag = xRS.EoF
        sql = "Select * from IGE1 where U_TWMICBD = " & APDE
        xRS.DoQuery(sql)
        dFlag = xRS.EoF
        If hFlag And dFlag Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Function CheckGoodReceived(ByVal APCNDE As String) As Boolean
        Dim sql As String = "Select * from OIGN Where DocStatus = 'O' and U_TWMSDER=" & APCNDE
        Dim xRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        xRS.DoQuery(sql)
        Return xRS.EoF
    End Function
    Private Function CheckICDoc(ByRef xCompany As SAPbobsCOM.Company, ByVal ARDE As String) As Boolean
        Dim sql As String = "Select ODRF.* from ODRF Inner Join OCRD on OCRD.CardCode = ODRF.CardCode "
        sql &= " Where ODRF.ObjType = 22 and DocStatus = 'O' and U_TWMSDER=" & ARDE
        sql &= " And U_TWMICDBN = '" & oCompany.CompanyDB & "'"
        Dim xRS As SAPbobsCOM.Recordset = xCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        SBO_Application.SetStatusBarMessage("Checking IC Document ...", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        Try
            xRS.DoQuery(sql)
        Catch ex As Exception
            SBO_Application.MessageBox("Error in Target Company : " & ex.Message)
        End Try
        Return xRS.EoF
    End Function
    Private Function CheckItem(ByRef xCompany As SAPbobsCOM.Company, ByVal pItem As String) As SAPbobsCOM.Recordset
        Dim xRec As SAPbobsCOM.Recordset = xCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        xRec.DoQuery("Select ItemCode from OITM where ItemCode = '" & pItem & "'")
        Return xRec
    End Function
    Private Function getFGT(ByVal oItem As String) As String
        Dim xItem As String = ""
        Dim aItem As String() = oItem.Split("-")
        If aItem.Length < 2 Then
            xItem = oItem
        Else
            Dim i As Integer
            For i = 0 To aItem.Length - 2
                xItem &= aItem(i) & "-"
            Next
            xItem &= "FGT"
        End If
        Return xItem
    End Function
    Private Function CreateICItem(ByRef xCompany As SAPbobsCOM.Company, ByVal pItem As String, Optional ByVal FG As Boolean = False) As Boolean
        Dim kReturn As Boolean = False
        Dim xItem As String = IIf(FG, getFGT(pItem), pItem)
        If CheckItem(xCompany, xItem).EoF Then
            Dim oItem As SAPbobsCOM.Items = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
            Dim nItem As SAPbobsCOM.Items = xCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)

            oItem.GetByKey(pItem)
            nItem.UserFields.Fields.Item("U_TWM_FAB").Value = oItem.ItemCode
            If FG Then
                'nItem.ItemCode = oItem.ItemCode & "-FG" '#### Old Style
                nItem.ItemCode = getFGT(oItem.ItemCode)
                nItem.ItemName = oItem.ItemName '& "-FGT" 'No need anymore
                nItem.InventoryUOM = "PCS"
            Else
                nItem.ItemCode = oItem.ItemCode
                nItem.ItemName = oItem.ItemName
                nItem.InventoryUOM = oItem.InventoryUOM
            End If

            If Right(nItem.ItemCode.ToString, 3) = "ACC" Then
                nItem.InventoryItem = SAPbobsCOM.BoYesNoEnum.tNO
            Else
                nItem.InventoryItem = SAPbobsCOM.BoYesNoEnum.tYES
            End If

            nItem.CostAccountingMethod = SAPbobsCOM.BoInventorySystem.bis_FIFO

            Dim err As Long = nItem.Add
            Dim errMsg As String = ""
            If err <> 0 Then
                xCompany.GetLastError(err, errMsg)
                SBO_Application.MessageBox("Fail to create Item : " & oItem.ItemCode & " Error : " & err & " msg : " & errMsg)
            Else
                kReturn = True
            End If
        Else
            kReturn = True
        End If
        Return kReturn
    End Function
    Private Function preProcessITEM(ByRef xCompany As SAPbobsCOM.Company, ByVal ARDE As String) As Boolean
        Dim kRet As Boolean = True
        Dim sql As String = "Select Distinct(ItemCode) from INV1 where docentry =" & ARDE
        Dim xRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim xItem As String = ""
        xRS.DoQuery(sql)
        xRS.MoveFirst()
        While Not xRS.EoF And kRet
            xItem = xRS.Fields.Item(0).Value
            kRet = CreateICItem(xCompany, xItem)
            If kRet Then kRet = CreateICItem(xCompany, xItem, True)
            xRS.MoveNext()
        End While
        Return kRet
    End Function
    Private Function CheckARBP(ByVal ARDE As String) As SAPbobsCOM.Recordset
        Dim sql As String = "select OCRD.U_TWMICDBN from OINV inner join OCRD ON OINV.CardCode = OCRD.CardCode "
        sql &= "where OCRD.U_TWMICYN=1 and OINV.DocEntry = " & ARDE
        Dim xRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        xRS.DoQuery(sql)
        Return xRS
    End Function

    Private Function CheckICBP(ByRef xCompany As SAPbobsCOM.Company, ByVal LocalCompanyName As String) As SAPbobsCOM.Recordset
        Dim sql As String = "Select CardCode from OCRD where U_TWMICYN='1' and U_TWMICDBN = '" & LocalCompanyName & "'"
        Dim xRS As SAPbobsCOM.Recordset = xCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            xRS.DoQuery(sql)
        Catch ex As Exception
            SBO_Application.MessageBox("Error in Target Company : " & ex.Message)
        End Try
        '################################################################################################
        ' This part is for ChinHo Testing
        '################################################################################################
        If xRS.EoF Then
            sql = "Select AliasName from OADM"
            xRS.DoQuery(sql)
            If Not xRS.EoF Then
                LocalCompanyName = xRS.Fields.Item(0).Value.ToString
                sql = "Select CardCode from OCRD where U_TWMICYN='1' and U_TWMICDBN = '" & LocalCompanyName & "'"
                Try
                    xRS.DoQuery(sql)
                Catch ex As Exception
                    SBO_Application.MessageBox("Error in Target Company : " & ex.Message)
                End Try
            End If
        End If
        '################################################################################################
        '################################################################################################
        Return xRS
    End Function
    Private Function getLocalCurrency(ByRef xCompany As SAPbobsCOM.Company) As String
        Dim oSBObob As SAPbobsCOM.SBObob
        Dim oRecordSet As SAPbobsCOM.Recordset

        oSBObob = xCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
        oRecordSet = xCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet = oSBObob.GetLocalCurrency
        oRecordSet.MoveFirst()
        Return oRecordSet.Fields.Item(0).Value
    End Function
    Private Sub AddICPODrf(ByRef xCompany As SAPbobsCOM.Company, ByVal ARDE As String, ByVal kCardCode As String)
        Dim oAR As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
        Dim oPO As SAPbobsCOM.Documents = xCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
        oPO.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseOrders
        oAR.GetByKey(ARDE)

        Dim LC As String = getLocalCurrency(xCompany)

        Dim oUserFields As SAPbobsCOM.UserFields
        Dim ii As Integer
        Dim rowcount As Integer
        Dim oFields As SAPbobsCOM.Fields
        Dim oField As SAPbobsCOM.Field
        Dim TxtName As String()
        Dim nVATGroup As String = ""
        Dim errormessage As String = ""

        Try
            '### Generating IC PO Draft ###
            '@@@ Header Filler Begin @@@
            oPO.CardCode = kCardCode ' "Get Default Customer for PO Draft
            oPO.DocDate = oAR.DocDate
            oPO.DocDueDate = oAR.DocDueDate
            If oAR.DocCurrency <> LC Then
                oPO.DocCurrency = oAR.DocCurrency
                oPO.DocRate = oAR.DocRate
            End If
            oPO.NumAtCard = oAR.NumAtCard
            oPO.DocType = oAR.DocType
            oPO.DiscountPercent = oAR.DiscountPercent
            If oAR.DocumentsOwner > 0 Then
                oPO.DocumentsOwner = oAR.DocumentsOwner
            End If
            If oAR.Comments <> "" Then oPO.Comments = oAR.Comments

            oUserFields = oPO.UserFields
            oFields = oUserFields.Fields
            ReDim TxtName(oFields.Count)
            ii = 0
            For Each oField In oFields
                TxtName(ii) = oField.Name
                Select Case TxtName(ii)
                    Case "U_TWMSCDBN"
                        oPO.UserFields.Fields.Item("U_TWMSCDBN").Value = vCmp.CompanyDB
                    Case "U_TWMBPR"
                        oPO.UserFields.Fields.Item("U_TWMBPR").Value = oAR.CardCode
                    Case "U_TWMSDER"
                        oPO.UserFields.Fields.Item("U_TWMSDER").Value = oAR.DocEntry
                    Case "U_TWMSDNR"
                        oPO.UserFields.Fields.Item("U_TWMSDNR").Value = oAR.DocNum
                    Case "U_TWMTDER", "U_TWMTDNR"
                        'Do Nothing
                    Case Else
                        If UDFCheck(oAR.UserFields.Fields, TxtName(ii)) = True Then
                            oPO.UserFields.Fields.Item(TxtName(ii)).Value = oAR.UserFields.Fields.Item(TxtName(ii)).Value
                        End If
                End Select
                ii = ii + 1
            Next oField
            '@@@ Header Filler END @@@


            '@@@ Detail Filler Start @@@
            If oPO.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items Then
                For rowcount = 0 To oAR.Lines.Count - 1
                    oAR.Lines.SetCurrentLine(rowcount)
                    oPO.Lines.ItemCode = oAR.Lines.ItemCode
                    oPO.Lines.ItemDescription = oAR.Lines.ItemDescription
                    oPO.Lines.Quantity = oAR.Lines.Quantity
                    oPO.Lines.UnitPrice = oAR.Lines.Price
                    If oAR.Lines.Currency <> LC Then
                        oPO.Lines.Rate = oAR.Lines.Rate
                    End If

                    oPO.Lines.ProjectCode = oAR.Lines.ProjectCode
                    If oAR.DocCurrency <> oAR.Lines.Currency Then
                        oPO.Lines.Currency = oAR.Lines.Currency
                    End If
                    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
                    'oPO.Lines.AccountCode = ICRevenue
                    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
                    oPO.Lines.DiscountPercent = oAR.Lines.DiscountPercent
                    oPO.Lines.ShipDate = oAR.Lines.ShipDate
                    If oPO.Lines.BaseType = -1 Then
                        nVATGroup = getVATGroup(oAR.Lines.VatGroup.Replace("O", "I"))
                        If nVATGroup <> "" Then oPO.Lines.VatGroup = nVATGroup
                    End If

                    ii = 0
                    ' add user-defined fields
                    oUserFields = oPO.Lines.UserFields
                    oFields = oUserFields.Fields
                    ReDim TxtName(oFields.Count)
                    For Each oField In oFields
                        TxtName(ii) = oField.Name
                        Select Case TxtName(ii)
                            Case "U_TWMICBD"
                                oPO.Lines.UserFields.Fields.Item("U_TWMICBD").Value = oAR.DocEntry.ToString
                            Case "U_TWMICBL"
                                oPO.Lines.UserFields.Fields.Item("U_TWMICBL").Value = oAR.Lines.LineNum.ToString   '  rowcount.ToString
                            Case Else
                                If UDFCheck(oAR.Lines.UserFields.Fields, TxtName(ii)) = True Then
                                    oPO.Lines.UserFields.Fields.Item(TxtName(ii)).Value = oAR.Lines.UserFields.Fields.Item(TxtName(ii)).Value
                                End If
                        End Select
                        ii = ii + 1
                    Next oField

                    If rowcount < oAR.Lines.Count - 1 Then
                        oPO.Lines.Add()
                    End If
                Next
            Else
                For rowcount = 0 To oAR.Lines.Count - 1
                    oAR.Lines.SetCurrentLine(rowcount)
                    'oPO.Lines.AccountCode = ICRevenue
                    oPO.Lines.ItemDescription = oAR.Lines.ItemDescription
                    If oAR.DocCurrency <> oAR.Lines.Currency Then
                        oPO.Lines.Currency = oAR.Lines.Currency
                    End If
                    oPO.Lines.Rate = oAR.Lines.Rate
                    oPO.Lines.ProjectCode = oAR.Lines.ProjectCode
                    oPO.Lines.Price = oAR.Lines.Price
                    oPO.Lines.DiscountPercent = oAR.Lines.DiscountPercent
                    oPO.Lines.UserFields.Fields.Item("U_TWMICBD").Value = oAR.DocEntry.ToString
                    oPO.Lines.UserFields.Fields.Item("U_TWMICBT").Value = "22"
                    oPO.Lines.UserFields.Fields.Item("U_TWM_SS").Value = "1"
                    oPO.Lines.UserFields.Fields.Item("U_TWMICBL").Value = rowcount.ToString
                    ' add user-defined fields
                    oUserFields = oPO.Lines.UserFields
                    oFields = oUserFields.Fields

                    ReDim TxtName(oFields.Count)
                    For Each oField In oFields
                        TxtName(ii) = oField.Name
                        If UDFCheck(oAR.Lines.UserFields.Fields, TxtName(ii)) Then
                            If TxtName(ii) <> "U_TWMICBD" Or TxtName(ii) <> "U_TWMICBT" Or TxtName(ii) <> "U_TWM_SS" Then
                                oPO.Lines.UserFields.Fields.Item(TxtName(ii)).Value = oAR.Lines.UserFields.Fields.Item(TxtName(ii)).Value
                            End If
                        End If
                        ii = ii + 1
                    Next oField
                    If rowcount < oAR.Lines.Count - 1 Then
                        oPO.Lines.Add()
                    End If
                Next
            End If
            '@@@ Detail Filler End @@@

            Dim Err As Long = oPO.Add
            Dim errmsg As String = ""
            If Err <> 0 Then
                xCompany.GetLastError(Err, errmsg)
                SBO_Application.MessageBox(errmsg)
                errormessage = errormessage & Chr(13) & errmsg
            Else
                Dim TargetEntry As String = ""
                xCompany.GetNewObjectCode(TargetEntry)
                oPO.GetByKey(TargetEntry)
                SBO_Application.SetStatusBarMessage("Created InterCompany PO Draft : " & oPO.DocNum, SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                oAR.UserFields.Fields.Item("U_TWMTDER").Value = TargetEntry
                oAR.UserFields.Fields.Item("U_TWMTDNR").Value = oPO.DocNum
                oAR.Update()
            End If
        Catch ex As Exception
            SBO_Application.MessageBox(ex.StackTrace & ": " & ex.Message)
        End Try

    End Sub
    Private Function getARGL(ByVal APGL As String) As String
        Dim kGL As String = runQueryStr("Select AcctCode from OACT where ACCTCODE = '4" & Mid(APGL, 2) & "'")
        Return kGL
    End Function
    Private Sub UpdateAPRef(ByVal APDE As String, ByVal kCardCode As String, ByVal APDraft As Boolean, ByVal ARDE As String, ByVal ARDN As String)
        Dim xRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim sql As String = ""
        If APDraft Then
            'AP Draft
            sql = "Update DRF1 Set U_TWMTDER = '" & ARDE & "', U_TWMTDNR = '" & ARDN & "'"
            sql &= " Where DRF1.DocEntry = '" & APDE & "' And U_TWM_BPCode = '" & kCardCode & "'"
        Else
            'AP
            sql = "Update PCH1 Set U_TWMTDER = '" & ARDE & "', U_TWMTDNR = '" & ARDN & "'"
            sql &= " Where PCH1.DocEntry = '" & APDE & "' And U_TWM_BPCode = '" & kCardCode & "'"
        End If
        xRS.DoQuery(sql)
    End Sub
    Private Sub AddARInvDrf_byBP(ByVal APDE As String, ByVal kCardCode As String, Optional ByVal APDraft As Boolean = False)
        SBO_Application.SetStatusBarMessage("Please wait Generating AR Invoice Draft for BP : " & kCardCode, SAPbouiCOM.BoMessageTime.bmt_Short, False)
        Dim oAR As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
        oAR.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices
        Dim oAP As SAPbobsCOM.Documents = Nothing
        If APDraft Then
            oAP = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
            oAP.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices
        Else
            oAP = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        End If

        oAP.GetByKey(APDE)

        Dim oUserFields As SAPbobsCOM.UserFields
        Dim ii As Integer
        Dim rowcount As Integer
        Dim ARrowFlag As Boolean = False
        Dim oFields As SAPbobsCOM.Fields
        Dim oField As SAPbobsCOM.Field
        Dim TxtName As String()
        Dim nVATGroup As String = ""
        Dim errormessage As String = ""
        Try
            '### Generating AR Invoice Draft ###
            '@@@ Header Filler Begin @@@
            oAR.CardCode = kCardCode ' "Get Default Customer for AR"
            oAR.DocDate = oAP.DocDate
            oAR.DocDueDate = oAP.DocDueDate
            oAR.DocCurrency = oAP.DocCurrency
            oAR.DocRate = oAP.DocRate
            oAR.NumAtCard = oAP.DocNum
            oAR.DocType = oAP.DocType
            oAR.DiscountPercent = oAP.DiscountPercent
            If oAP.DocumentsOwner > 0 Then
                oAR.DocumentsOwner = oAP.DocumentsOwner
            End If

            ARMarkUp = GetMarkup(oAP.DocDate)

            For i As Integer = 0 To oAR.UserFields.Fields.Count - 1
                If oAR.UserFields.Fields.Item(i).Name.ToString() <> "U_TWMTDER" And oAR.UserFields.Fields.Item(i).Name.ToString() <> "U_TWMTDNR" Then
                    oAR.UserFields.Fields.Item(i).Value = oAP.UserFields.Fields.Item(i).Value
                End If
            Next i

            oAR.UserFields.Fields.Item("U_TWMSCDBN").Value = vCmp.CompanyDB
            oAR.UserFields.Fields.Item("U_TWMBPR").Value = oAP.CardCode
            oAR.UserFields.Fields.Item("U_TWMSDER").Value = oAP.DocEntry
            oAR.UserFields.Fields.Item("U_TWMSDNR").Value = oAP.DocNum

            'oUserFields = oAR.UserFields
            'oFields = oUserFields.Fields
            'ReDim TxtName(oFields.Count)
            'ii = 0
            'For Each oField In oFields
            '    TxtName(ii) = oField.Name
            '    Select Case TxtName(ii)
            '        Case "U_TWMSCDBN"
            '            oAR.UserFields.Fields.Item("U_TWMSCDBN").Value = vCmp.CompanyDB
            '        Case "U_TWMBPR"
            '            oAR.UserFields.Fields.Item("U_TWMBPR").Value = oAP.CardCode
            '        Case "U_TWMSDER"
            '            oAR.UserFields.Fields.Item("U_TWMSDER").Value = oAP.DocEntry
            '        Case "U_TWMSDNR"
            '            oAR.UserFields.Fields.Item("U_TWMSDNR").Value = oAP.DocNum
            '        Case "U_TWMTDER", "U_TWMTDNR"
            '            'Do Nothing
            '        Case Else
            '            If UDFCheck(oAP.UserFields.Fields, TxtName(ii)) = True Then
            '                oAR.UserFields.Fields.Item(TxtName(ii)).Value = oAP.UserFields.Fields.Item(TxtName(ii)).Value
            '            End If
            '    End Select
            '    ii = ii + 1
            'Next oField

            '@@@ Header Filler END @@@

            '@@@ Detail Filler Start @@@
            If oAR.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items Then

                'Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'Dim fieldsquery As String = "SELECT 'U_' + AliasID FieldName FROM CUFD WHERE TableID = 'INV1'"
                'oRS.DoQuery(fieldsquery)

                For rowcount = 0 To oAP.Lines.Count - 1
                    oAP.Lines.SetCurrentLine(rowcount)
                    If oAP.Lines.UserFields.Fields.Item("U_TWM_BPCode").Value = kCardCode Then
                        'Add this line to AR
                        If ARrowFlag Then 'Add New Row from 2nd Onward
                            oAR.Lines.Add()
                        Else 'Come in As False for the First Time
                            ARrowFlag = True
                        End If
                        oAR.Lines.ItemCode = oAP.Lines.ItemCode
                        oAR.Lines.ItemDescription = oAP.Lines.ItemDescription
                        oAR.Lines.Quantity = oAP.Lines.Quantity
                        oAR.Lines.UnitPrice = oAP.Lines.Price * ARMarkUp
                        oAR.Lines.Rate = oAP.Lines.Rate
                        oAR.Lines.WarehouseCode = oAP.Lines.WarehouseCode

                        oAR.Lines.ProjectCode = oAP.Lines.ProjectCode
                        If oAP.DocCurrency <> oAP.Lines.Currency Then
                            oAR.Lines.Currency = oAP.Lines.Currency
                        End If
                        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
                        'oAR.Lines.AccountCode = ICRevenue
                        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
                        oAR.Lines.DiscountPercent = oAP.Lines.DiscountPercent
                        oAR.Lines.ShipDate = oAP.Lines.ShipDate
                        If oAR.Lines.BaseType = -1 Then
                            nVATGroup = getVATGroup(oAP.Lines.VatGroup.Replace("I", "O"))
                            If nVATGroup <> "" Then oAR.Lines.VatGroup = nVATGroup
                        End If
                        ii = 0
                        ' add user-defined fields

                        For i As Integer = 0 To oAR.Lines.UserFields.Fields.Count - 1
                            oAR.Lines.UserFields.Fields.Item(i).Value = oAP.Lines.UserFields.Fields.Item(i).Value
                        Next i
                        oAR.Lines.UserFields.Fields.Item("U_TWMICBD").Value = oAP.DocEntry.ToString
                        oAR.Lines.UserFields.Fields.Item("U_TWMICBL").Value = oAP.Lines.LineNum.ToString

                        'While Not oRS.EoF
                        '    Dim fieldName As String = oRS.Fields.Item("FieldName").Value.ToString()
                        '    Select Case fieldName
                        '        Case "U_TWMICBD"
                        '            oAR.Lines.UserFields.Fields.Item("U_TWMICBD").Value = oAP.DocEntry.ToString

                        '        Case "U_TWMICBL"
                        '            oAR.Lines.UserFields.Fields.Item("U_TWMICBL").Value = oAP.Lines.LineNum.ToString
                        '        Case Else
                        '            If UDFCheck(oAP.Lines.UserFields.Fields, fieldName) = True Then
                        '                oAR.Lines.UserFields.Fields.Item(fieldName).Value = oAP.Lines.UserFields.Fields.Item(fieldName).Value
                        '            End If
                        '    End Select

                        '    oRS.MoveNext()
                        'End While


                        'oUserFields = oAR.Lines.UserFields
                        'oFields = oUserFields.Fields
                        'ReDim TxtName(oFields.Count)
                        'For Each oField In oFields
                        '    TxtName(ii) = oField.Name
                        '    Select Case TxtName(ii)
                        '        Case "U_TWMICBD"
                        '            oAR.Lines.UserFields.Fields.Item("U_TWMICBD").Value = oAP.DocEntry.ToString

                        '        Case "U_TWMICBL"
                        '            oAR.Lines.UserFields.Fields.Item("U_TWMICBL").Value = oAP.Lines.LineNum.ToString   '  rowcount.ToString
                        '        Case Else
                        '            If UDFCheck(oAP.Lines.UserFields.Fields, TxtName(ii)) = True Then
                        '                oAR.Lines.UserFields.Fields.Item(TxtName(ii)).Value = oAP.Lines.UserFields.Fields.Item(TxtName(ii)).Value
                        '            End If
                        '    End Select
                        '    ii = ii + 1
                        'Next oField


                    End If
                Next
            Else
                For rowcount = 0 To oAP.Lines.Count - 1
                    oAP.Lines.SetCurrentLine(rowcount)
                    If oAP.Lines.UserFields.Fields.Item("U_TWM_BPCode").Value = kCardCode Then
                        'Add this line to AR
                        If ARrowFlag Then 'Add New Row from 2nd Onward
                            oAR.Lines.Add()
                        Else 'Come in As False for the First Time
                            ARrowFlag = True
                        End If
                        oAR.Lines.AccountCode = getARGL(oAP.Lines.AccountCode.ToString)
                        If oAR.Lines.AccountCode = "" Then
                            SBO_Application.MessageBox("There is no relating ARGL code for Line " & oAR.Lines.LineNum & " APGL Code : " & oAP.Lines.AccountCode)
                        End If
                        oAR.Lines.ItemDescription = oAP.Lines.ItemDescription
                        If oAP.DocCurrency <> oAP.Lines.Currency Then
                            oAR.Lines.Currency = oAP.Lines.Currency
                        End If
                        oAR.Lines.Rate = oAP.Lines.Rate
                        oAR.Lines.ProjectCode = oAP.Lines.ProjectCode
                        oAR.Lines.Price = oAP.Lines.Price
                        oAR.Lines.DiscountPercent = oAP.Lines.DiscountPercent
                        oAR.Lines.LineTotal = oAP.Lines.LineTotal

                        For i As Integer = 0 To oAR.Lines.UserFields.Fields.Count - 1
                            oAR.Lines.UserFields.Fields.Item(i).Value = oAP.Lines.UserFields.Fields.Item(i).Value
                        Next i

                        oAR.Lines.UserFields.Fields.Item("U_TWMICBD").Value = oAP.DocEntry.ToString
                        oAR.Lines.UserFields.Fields.Item("U_TWMICBT").Value = "22"
                        oAR.Lines.UserFields.Fields.Item("U_TWM_SS").Value = "1"
                        oAR.Lines.UserFields.Fields.Item("U_TWMICBL").Value = rowcount.ToString
                        ' add user-defined fields

                        'oUserFields = oAR.Lines.UserFields
                        'oFields = oUserFields.Fields
                        'ii = 0
                        'ReDim TxtName(oFields.Count)
                        'For Each oField In oFields
                        '    TxtName(ii) = oField.Name
                        '    If UDFCheck(oAP.Lines.UserFields.Fields, TxtName(ii)) Then
                        '        If TxtName(ii) <> "U_TWMICBD" Or TxtName(ii) <> "U_TWMICBT" Or TxtName(ii) <> "U_TWM_SS" Then
                        '            oAR.Lines.UserFields.Fields.Item(TxtName(ii)).Value = oAP.Lines.UserFields.Fields.Item(TxtName(ii)).Value
                        '        End If
                        '    End If
                        '    ii = ii + 1
                        'Next oField
                    End If
                Next
            End If
            '@@@ Detail Filler End @@@
            'vCmp.StartTransaction()
            Dim Err As Long = oAR.Add
            Dim errmsg As String = ""
            Dim msg As String = ""
            If Err <> 0 Then
                vCmp.GetLastError(Err, errmsg)
                SBO_Application.MessageBox(errmsg)
                errormessage = errormessage & Chr(13) & errmsg
                'vCmp.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            Else
                Dim TargetEntry As String = ""
                vCmp.GetNewObjectCode(TargetEntry)
                oAR.GetByKey(TargetEntry)
                msg = "Created AR Invoice Draft : " & TargetEntry & " DocNum : " & oAR.DocNum
                SBO_Application.SetStatusBarMessage(msg, SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                oAP.UserFields.Fields.Item("U_TWMTDER").Value = TargetEntry
                oAP.UserFields.Fields.Item("U_TWMTDNR").Value = oAR.DocNum
                oAP.Update()
                UpdateAPRef(APDE, kCardCode, APDraft, TargetEntry, oAR.DocNum)
                'vCmp.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.StackTrace & ": " & ex.Message)
        End Try
    End Sub
    Private Sub AddARInvDrf(ByVal APDE As String, ByVal kCardCode As String, Optional ByVal APDraft As Boolean = False)
        Dim oAR As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
        oAR.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices
        Dim oAP As SAPbobsCOM.Documents = Nothing
        If APDraft Then
            oAP = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
            oAP.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices
        Else
            oAP = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        End If

        oAP.GetByKey(APDE)

        Dim oUserFields As SAPbobsCOM.UserFields
        Dim ii As Integer
        Dim rowcount As Integer
        Dim oFields As SAPbobsCOM.Fields
        Dim oField As SAPbobsCOM.Field
        Dim TxtName As String()
        Dim nVATGroup As String = ""
        Dim errormessage As String = ""
        Try
            '### Generating AR Invoice Draft ###
            '@@@ Header Filler Begin @@@
            oAR.CardCode = kCardCode ' "Get Default Customer for AR"
            oAR.DocDate = oAP.DocDate
            oAR.DocDueDate = oAP.DocDueDate
            oAR.DocCurrency = oAP.DocCurrency
            oAR.DocRate = oAP.DocRate
            oAR.NumAtCard = oAP.DocNum
            oAR.DocType = oAP.DocType
            oAR.DiscountPercent = oAP.DiscountPercent
            If oAP.DocumentsOwner > 0 Then
                oAR.DocumentsOwner = oAP.DocumentsOwner
            End If

            ARMarkUp = GetMarkup(oAP.DocDate)

            oUserFields = oAR.UserFields
            oFields = oUserFields.Fields
            ReDim TxtName(oFields.Count)
            ii = 0
            For Each oField In oFields
                TxtName(ii) = oField.Name
                Select Case TxtName(ii)
                    Case "U_TWMSCDBN"
                        oAR.UserFields.Fields.Item("U_TWMSCDBN").Value = vCmp.CompanyDB
                    Case "U_TWMBPR"
                        oAR.UserFields.Fields.Item("U_TWMBPR").Value = oAP.CardCode
                    Case "U_TWMSDER"
                        oAR.UserFields.Fields.Item("U_TWMSDER").Value = oAP.DocEntry
                    Case "U_TWMSDNR"
                        oAR.UserFields.Fields.Item("U_TWMSDNR").Value = oAP.DocNum
                    Case "U_TWMTDER", "U_TWMTDNR"
                        'Do Nothing
                    Case Else
                        If UDFCheck(oAP.UserFields.Fields, TxtName(ii)) = True Then
                            oAR.UserFields.Fields.Item(TxtName(ii)).Value = oAP.UserFields.Fields.Item(TxtName(ii)).Value
                        End If
                End Select
                ii = ii + 1
            Next oField
            '@@@ Header Filler END @@@

            '@@@ Detail Filler Start @@@
            If oAR.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items Then
                For rowcount = 0 To oAP.Lines.Count - 1
                    oAP.Lines.SetCurrentLine(rowcount)
                    oAR.Lines.ItemCode = oAP.Lines.ItemCode
                    oAR.Lines.ItemDescription = oAP.Lines.ItemDescription
                    oAR.Lines.Quantity = oAP.Lines.Quantity
                    oAR.Lines.UnitPrice = oAP.Lines.Price * ARMarkUp
                    oAR.Lines.Rate = oAP.Lines.Rate
                    oAR.Lines.WarehouseCode = oAP.Lines.WarehouseCode

                    oAR.Lines.ProjectCode = oAP.Lines.ProjectCode
                    If oAP.DocCurrency <> oAP.Lines.Currency Then
                        oAR.Lines.Currency = oAP.Lines.Currency
                    End If
                    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
                    'oAR.Lines.AccountCode = ICRevenue
                    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
                    oAR.Lines.DiscountPercent = oAP.Lines.DiscountPercent
                    oAR.Lines.ShipDate = oAP.Lines.ShipDate
                    If oAR.Lines.BaseType = -1 Then
                        nVATGroup = getVATGroup(oAP.Lines.VatGroup.Replace("I", "O"))
                        If nVATGroup <> "" Then oAR.Lines.VatGroup = nVATGroup
                    End If
                    ii = 0
                    ' add user-defined fields
                    oUserFields = oAR.Lines.UserFields
                    oFields = oUserFields.Fields
                    ReDim TxtName(oFields.Count)
                    For Each oField In oFields
                        TxtName(ii) = oField.Name
                        Select Case TxtName(ii)
                            Case "U_TWMICBD"
                                oAR.Lines.UserFields.Fields.Item("U_TWMICBD").Value = oAP.DocEntry.ToString
                            Case "U_TWMICBL"
                                oAR.Lines.UserFields.Fields.Item("U_TWMICBL").Value = oAP.Lines.LineNum.ToString   '  rowcount.ToString
                            Case Else
                                If UDFCheck(oAP.Lines.UserFields.Fields, TxtName(ii)) = True Then
                                    oAR.Lines.UserFields.Fields.Item(TxtName(ii)).Value = oAP.Lines.UserFields.Fields.Item(TxtName(ii)).Value
                                End If
                        End Select
                        ii = ii + 1
                    Next oField

                    If rowcount < oAP.Lines.Count - 1 Then
                        oAR.Lines.Add()
                    End If
                Next
            Else
                For rowcount = 0 To oAP.Lines.Count - 1
                    oAP.Lines.SetCurrentLine(rowcount)
                    oAR.Lines.AccountCode = getARGL(oAP.Lines.AccountCode.ToString)
                    If oAR.Lines.AccountCode = "" Then
                        SBO_Application.MessageBox("There is no relating ARGL code for Line " & oAR.Lines.LineNum & " APGL Code : " & oAP.Lines.AccountCode)
                    End If
                    oAR.Lines.ItemDescription = oAP.Lines.ItemDescription
                    If oAP.DocCurrency <> oAP.Lines.Currency Then
                        oAR.Lines.Currency = oAP.Lines.Currency
                    End If
                    oAR.Lines.Rate = oAP.Lines.Rate
                    oAR.Lines.ProjectCode = oAP.Lines.ProjectCode
                    oAR.Lines.Price = oAP.Lines.Price
                    oAR.Lines.DiscountPercent = oAP.Lines.DiscountPercent
                    oAR.Lines.LineTotal = oAP.Lines.LineTotal
                    oAR.Lines.UserFields.Fields.Item("U_TWMICBD").Value = oAP.DocEntry.ToString
                    oAR.Lines.UserFields.Fields.Item("U_TWMICBT").Value = "22"
                    oAR.Lines.UserFields.Fields.Item("U_TWM_SS").Value = "1"
                    oAR.Lines.UserFields.Fields.Item("U_TWMICBL").Value = rowcount.ToString
                    ' add user-defined fields
                    oUserFields = oAR.Lines.UserFields
                    oFields = oUserFields.Fields
                    ii = 0
                    ReDim TxtName(oFields.Count)
                    For Each oField In oFields
                        TxtName(ii) = oField.Name
                        If UDFCheck(oAP.Lines.UserFields.Fields, TxtName(ii)) Then
                            If TxtName(ii) <> "U_TWMICBD" Or TxtName(ii) <> "U_TWMICBT" Or TxtName(ii) <> "U_TWM_SS" Then
                                oAR.Lines.UserFields.Fields.Item(TxtName(ii)).Value = oAP.Lines.UserFields.Fields.Item(TxtName(ii)).Value
                            End If
                        End If
                        ii = ii + 1
                    Next oField
                    If rowcount < oAP.Lines.Count - 1 Then
                        oAR.Lines.Add()
                    End If
                Next
            End If
            '@@@ Detail Filler End @@@
            'vCmp.StartTransaction()
            Dim Err As Long = oAR.Add
            Dim errmsg As String = ""
            If Err <> 0 Then
                vCmp.GetLastError(Err, errmsg)
                SBO_Application.MessageBox(errmsg)
                errormessage = errormessage & Chr(13) & errmsg
                'vCmp.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            Else
                Dim TargetEntry As String = ""
                vCmp.GetNewObjectCode(TargetEntry)
                oAR.GetByKey(TargetEntry)
                SBO_Application.SetStatusBarMessage("Created AR Invoice Draft : " & oAR.DocNum, SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                oAP.UserFields.Fields.Item("U_TWMTDER").Value = TargetEntry
                oAP.UserFields.Fields.Item("U_TWMTDNR").Value = oAR.DocNum
                oAP.Update()
                'vCmp.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.StackTrace & ": " & ex.Message)
        End Try
    End Sub

    Private Function LookUpItem(ByRef iCode As String, ByRef iCon As Double) As Boolean
        Dim kReturn As Boolean = False
        If Right(iCode, 3) = "FGT" Then
            Dim kRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            kRS.DoQuery("Select U_TWM_Conversion from OITM where Itemcode = '" & iCode & "'")
            iCon = IIf(kRS.EoF, 0, Val(kRS.Fields.Item(0).Value))

            'iCode = Left(iCode, iCode.Length - 3) & "FAB"
            'kRS.DoQuery("Select Itemcode from OITM where ItemCode = '" & iCode & "'")
            kRS.DoQuery("Select ISNULL(U_TWM_FAB,'') from OITM where ItemCode = '" & iCode & "'")
            If kRS.EoF Then
                kReturn = False
            Else
                iCode = kRS.Fields.Item(0).Value
                If iCode = "" Then
                    kReturn = False
                Else
                    kReturn = True
                End If
            End If
        End If
        Return kReturn
    End Function
    Private Sub AddGoodIssue(ByVal APDE As String)

        Dim oAP As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        Dim oGI As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
        oAP.GetByKey(APDE)

        Dim oUserFields As SAPbobsCOM.UserFields
        Dim ii As Integer
        Dim rowcount As Integer
        Dim oFields As SAPbobsCOM.Fields
        Dim oField As SAPbobsCOM.Field
        Dim TxtName As String()
        Dim nVATGroup As String = ""
        Dim errormessage As String = ""
        Dim iItemCode As String = ""
        Dim iCon As Double = 0

        Dim postFlag As Boolean = False

        Try
            '### Generating Good Issue ###
            '@@@ Header Filler Begin @@@
            oGI.DocDate = oAP.DocDate
            oGI.TaxDate = oAP.TaxDate

            oUserFields = oGI.UserFields
            oFields = oUserFields.Fields
            ReDim TxtName(oFields.Count)
            ii = 0
            For Each oField In oFields
                TxtName(ii) = oField.Name
                Select Case TxtName(ii)
                    Case "U_TWMSCDBN"
                        oGI.UserFields.Fields.Item("U_TWMSCDBN").Value = vCmp.CompanyDB
                    Case "U_TWMBPR"
                        oGI.UserFields.Fields.Item("U_TWMBPR").Value = oAP.CardCode
                    Case "U_TWMSDER"
                        oGI.UserFields.Fields.Item("U_TWMSDER").Value = oAP.DocEntry
                    Case "U_TWMSDNR"
                        oGI.UserFields.Fields.Item("U_TWMSDNR").Value = oAP.DocNum
                    Case "U_TWMTDER", "U_TWMTDNR"
                        'Do Nothing
                    Case Else
                        If UDFCheck(oAP.UserFields.Fields, TxtName(ii)) = True Then
                            oGI.UserFields.Fields.Item(TxtName(ii)).Value = oAP.UserFields.Fields.Item(TxtName(ii)).Value
                        End If
                End Select
                ii = ii + 1
            Next oField
            '@@@ Header Filler END @@@

            '@@@ Detail Filler Start @@@
            For rowcount = 0 To oAP.Lines.Count - 1
                oAP.Lines.SetCurrentLine(rowcount)
                iItemCode = oAP.Lines.ItemCode
                If LookUpItem(iItemCode, iCon) Then
                    postFlag = True
                    'oGI.Lines.ItemCode = oAP.Lines.ItemCode
                    oGI.Lines.ItemCode = iItemCode

                    '20130220 Modified by Edy as per KK email.
                    'oGI.Lines.Quantity = oAP.Lines.Quantity * iCon
                    '## Wrong UOM value ( MUST PICK UP FROM UDF U_UOMSP )
                    'oGI.Lines.Quantity = oAP.Lines.Quantity * iCon * oAP.Lines.UnitsOfMeasurment

                    '######## Pick value from UDF and made sure it's at least 1
                    Dim oAP_Lines_UOMSP As Double = Val(oAP.Lines.UserFields.Fields.Item("U_TWM_UOMSP").Value)
                    If oAP_Lines_UOMSP = 0 Then oAP_Lines_UOMSP = 1

                    oGI.Lines.Quantity = oAP.Lines.Quantity * iCon * oAP_Lines_UOMSP

                    oGI.Lines.ProjectCode = oAP.Lines.ProjectCode
                    oGI.Lines.COGSCostingCode = oAP.Lines.COGSCostingCode
                    oGI.Lines.WarehouseCode = oAP.Lines.WarehouseCode
                    ii = 0
                    ' add user-defined fields
                    oUserFields = oGI.Lines.UserFields
                    oFields = oUserFields.Fields
                    ReDim TxtName(oFields.Count)
                    For Each oField In oFields
                        TxtName(ii) = oField.Name
                        Select Case TxtName(ii)
                            Case "U_TWMICBD"
                                oGI.Lines.UserFields.Fields.Item("U_TWMICBD").Value = oAP.DocEntry.ToString
                            Case "U_TWMICBL"
                                oGI.Lines.UserFields.Fields.Item("U_TWMICBL").Value = oAP.Lines.LineNum.ToString   '  rowcount.ToString
                            Case Else
                                If UDFCheck(oAP.Lines.UserFields.Fields, TxtName(ii)) = True Then
                                    oGI.Lines.UserFields.Fields.Item(TxtName(ii)).Value = oAP.Lines.UserFields.Fields.Item(TxtName(ii)).Value
                                End If
                        End Select
                        ii = ii + 1
                    Next oField

                    If rowcount < oAP.Lines.Count - 1 Then
                        oGI.Lines.Add()
                    End If
                End If
            Next
            '@@@ Detail Filler End @@@
            If postFlag Then
                Dim Err As Long = oGI.Add
                Dim errmsg As String = ""
                If Err <> 0 Then
                    vCmp.GetLastError(Err, errmsg)
                    SBO_Application.MessageBox(errmsg)
                    errormessage = errormessage & Chr(13) & errmsg
                Else
                    Dim TargetEntry As String = ""
                    vCmp.GetNewObjectCode(TargetEntry)
                    oGI.GetByKey(TargetEntry)
                    SBO_Application.SetStatusBarMessage("Created Good Issue : " & oGI.DocNum, SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    oAP.UserFields.Fields.Item("U_TWMTDER").Value = TargetEntry
                    oAP.UserFields.Fields.Item("U_TWMTDNR").Value = oGI.DocNum
                    oAP.Update()
                End If
            Else
                SBO_Application.SetStatusBarMessage("No Line to Generate Good Issue.", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            End If


        Catch ex As Exception

        End Try
    End Sub

    Private Function getFABPrice(ByVal iCode As String, ByRef oPrice As Double, ByVal APDE As Integer, ByVal APLN As Integer) As Double
        Dim xPrice As Double = oPrice
        Dim kRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        kRS.DoQuery("select StockPrice from IGE1 where U_TWMICBD = " & APDE & " AND U_TWMICBL= " & APLN & " And ItemCode = '" & iCode & "'")
        If Not kRS.EoF Then
            xPrice = kRS.Fields.Item(0).Value
        End If
        Return xPrice
    End Function

    Private Sub AddGoodReceived(ByVal APCNDE As String)

        Dim oAPCN As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
        Dim oGR As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)
        oAPCN.GetByKey(APCNDE)

        Dim oUserFields As SAPbobsCOM.UserFields
        Dim ii As Integer
        Dim rowcount As Integer
        Dim oFields As SAPbobsCOM.Fields
        Dim oField As SAPbobsCOM.Field
        Dim TxtName As String()
        Dim nVATGroup As String = ""
        Dim errormessage As String = ""
        Dim iItemCode As String = ""
        Dim iCon As Double = 0

        Dim postFlag As Boolean = False

        Try
            '### Generating Good Issue ###
            '@@@ Header Filler Begin @@@
            oGR.DocDate = oAPCN.DocDate
            oGR.TaxDate = oAPCN.TaxDate

            oUserFields = oGR.UserFields
            oFields = oUserFields.Fields
            ReDim TxtName(oFields.Count)
            ii = 0
            For Each oField In oFields
                TxtName(ii) = oField.Name
                Select Case TxtName(ii)
                    Case "U_TWMSCDBN"
                        oGR.UserFields.Fields.Item("U_TWMSCDBN").Value = vCmp.CompanyDB
                    Case "U_TWMBPR"
                        oGR.UserFields.Fields.Item("U_TWMBPR").Value = oAPCN.CardCode
                    Case "U_TWMSDER"
                        oGR.UserFields.Fields.Item("U_TWMSDER").Value = oAPCN.DocEntry
                    Case "U_TWMSDNR"
                        oGR.UserFields.Fields.Item("U_TWMSDNR").Value = oAPCN.DocNum
                    Case "U_TWMTDER", "U_TWMTDNR"
                        'Do Nothing
                    Case Else
                        If UDFCheck(oAPCN.UserFields.Fields, TxtName(ii)) = True Then
                            oGR.UserFields.Fields.Item(TxtName(ii)).Value = oAPCN.UserFields.Fields.Item(TxtName(ii)).Value
                        End If
                End Select
                ii = ii + 1
            Next oField
            '@@@ Header Filler END @@@

            '@@@ Detail Filler Start @@@
            For rowcount = 0 To oAPCN.Lines.Count - 1
                oAPCN.Lines.SetCurrentLine(rowcount)
                iItemCode = oAPCN.Lines.ItemCode
                If LookUpItem(iItemCode, iCon) Then
                    postFlag = True
                    'oGI.Lines.ItemCode = oAP.Lines.ItemCode
                    oGR.Lines.ItemCode = iItemCode
                    'oGR.Lines.Quantity = oAPCN.Lines.Quantity * iCon 
                    'oGR.Lines.Quantity = oAPCN.Lines.Quantity * iCon * oAPCN.Lines.UnitsOfMeasurment

                    '######## Pick value from UDF and made sure it's at least 1
                    Dim oAPCN_Lines_UOMSP As Double = Val(oAPCN.Lines.UserFields.Fields.Item("U_TWM_UOMSP").Value)
                    If oAPCN_Lines_UOMSP = 0 Then oAPCN_Lines_UOMSP = 1
                    oGR.Lines.Quantity = oAPCN.Lines.Quantity * iCon * oAPCN_Lines_UOMSP

                    oGR.Lines.ProjectCode = oAPCN.Lines.ProjectCode
                    oGR.Lines.COGSCostingCode = oAPCN.Lines.COGSCostingCode
                    oGR.Lines.Price = getFABPrice(iItemCode, oAPCN.Lines.Price, oAPCN.Lines.BaseEntry, oAPCN.Lines.BaseLine)
                    oGR.Lines.WarehouseCode = oAPCN.Lines.WarehouseCode

                    ii = 0
                    ' add user-defined fields
                    oUserFields = oGR.Lines.UserFields
                    oFields = oUserFields.Fields
                    ReDim TxtName(oFields.Count)
                    For Each oField In oFields
                        TxtName(ii) = oField.Name
                        Select Case TxtName(ii)
                            Case "U_TWMICBD"
                                oGR.Lines.UserFields.Fields.Item("U_TWMICBD").Value = oAPCN.DocEntry.ToString
                            Case "U_TWMICBL"
                                oGR.Lines.UserFields.Fields.Item("U_TWMICBL").Value = oAPCN.Lines.LineNum.ToString   '  rowcount.ToString
                            Case Else
                                If UDFCheck(oAPCN.Lines.UserFields.Fields, TxtName(ii)) = True Then
                                    oGR.Lines.UserFields.Fields.Item(TxtName(ii)).Value = oAPCN.Lines.UserFields.Fields.Item(TxtName(ii)).Value
                                End If
                        End Select
                        ii = ii + 1
                    Next oField

                    If rowcount < oAPCN.Lines.Count - 1 Then
                        oGR.Lines.Add()
                    End If
                End If
            Next
            '@@@ Detail Filler End @@@

            If postFlag Then
                Dim Err As Long = oGR.Add
                Dim errmsg As String = ""
                If Err <> 0 Then
                    vCmp.GetLastError(Err, errmsg)
                    SBO_Application.MessageBox(errmsg)
                    errormessage = errormessage & Chr(13) & errmsg
                Else
                    Dim TargetEntry As String = ""
                    vCmp.GetNewObjectCode(TargetEntry)
                    oGR.GetByKey(TargetEntry)
                    SBO_Application.SetStatusBarMessage("Created Good Received : " & oGR.DocNum, SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    oAPCN.UserFields.Fields.Item("U_TWMTDER").Value = TargetEntry
                    oAPCN.UserFields.Fields.Item("U_TWMTDNR").Value = oGR.DocNum
                    oAPCN.Update()
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub
    Private Sub Filter_APInv_GoodIssue(ByVal xForm As SAPbouiCOM.Form)

        Dim xRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim sql As String = "" '"select top 1 docentry, doctype from OPCH WHERE usersign=" & oCompany.UserSignature & " order by docentry desc "
        Dim APInvDE As String = ""
        Dim APInvDT As String = ""

        Select Case xForm.Mode
            Case SAPbouiCOM.BoFormMode.fm_ADD_MODE
                APInvDE = DocEntry
                sql = "Select doctype from OPCH where DocEntry = '" & APInvDE & "'"
                xRS.DoQuery(sql)
                If Not xRS.EoF Then
                    APInvDT = xRS.Fields.Item(0).Value
                    If APInvDT <> "I" Then APInvDE = ""
                    '### Generating Good Issue ###
                    SBO_Application.SetStatusBarMessage("Add Mode for " & DocNum & " <" & DocEntry & ">", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                End If

                'xRS.DoQuery(sql)
                'If Not xRS.EoF Then
                '    APInvDT = xRS.Fields.Item(1).Value
                '    APInvDE = xRS.Fields.Item(0).Value
                '    If APInvDT <> "I" Then APInvDE = ""
                '    '### Generating Good Issue ###
                '    SBO_Application.SetStatusBarMessage("Add Mode for " & APInvDE, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                'End If
            Case SAPbouiCOM.BoFormMode.fm_UPDATE_MODE, SAPbouiCOM.BoFormMode.fm_OK_MODE
                APInvDT = xForm.DataSources.DBDataSources.Item(0).GetValue("DocType", 0).ToString
                APInvDE = xForm.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0).ToString
                If APInvDT <> "I" Then APInvDE = ""
                SBO_Application.SetStatusBarMessage("Update Mode for " & APInvDE, SAPbouiCOM.BoMessageTime.bmt_Short, False)
        End Select

        If APInvDE <> "" Then
            If CheckGoodIssue(APInvDE) Then
                AddGoodIssue(APInvDE) 'xyz
            Else
                SBO_Application.MessageBox("Good Issue based on the current AP Invoice was already generated.")
            End If
        End If
    End Sub

    Private Sub Filter_APCN_GoodReceived(ByVal xForm As SAPbouiCOM.Form)

        Dim xRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim sql As String = "" '"select top 1 docentry, doctype from ORPC WHERE usersign=" & oCompany.UserSignature & " order by docentry desc "
        Dim APCNDE As String = ""
        Dim APCNDT As String = ""
        Select Case xForm.Mode
            Case SAPbouiCOM.BoFormMode.fm_ADD_MODE
                APCNDE = DocEntry
                sql = "Select doctype from ORPC where docentry = '" & APCNDE & "'"
                xRS.DoQuery(sql)
                If Not xRS.EoF Then
                    APCNDT = xRS.Fields.Item(0).Value
                    If APCNDT <> "I" Then APCNDE = ""
                    '### Generating Good Issue ###
                    SBO_Application.SetStatusBarMessage("Add Mode for " & DocNum & " <" & APCNDE & ">", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                End If

                'xRS.DoQuery(sql)
                'If Not xRS.EoF Then
                '    APCNDT = xRS.Fields.Item(1).Value
                '    APCNDE = xRS.Fields.Item(0).Value
                '    If APCNDT <> "I" Then APCNDE = ""
                '    ### Generating Good Issue ###
                '    SBO_Application.SetStatusBarMessage("Add Mode for " & APCNDE, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                'End If
            Case SAPbouiCOM.BoFormMode.fm_UPDATE_MODE, SAPbouiCOM.BoFormMode.fm_OK_MODE
                APCNDT = xForm.DataSources.DBDataSources.Item(0).GetValue("DocType", 0).ToString
                APCNDE = xForm.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0).ToString
                If APCNDT <> "I" Then APCNDE = ""
                SBO_Application.SetStatusBarMessage("Update Mode for " & APCNDE, SAPbouiCOM.BoMessageTime.bmt_Short, False)
        End Select

        If APCNDE <> "" Then
            If CheckGoodReceived(APCNDE) Then
                AddGoodReceived(APCNDE) 'xyz
            Else
                SBO_Application.MessageBox("Good Received based on the current APCN is already Generated.")
            End If
        End If
    End Sub

    Private Sub Filter_AP_ARINVDrf(ByVal xForm As SAPbouiCOM.Form)
        Dim xRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim sql As String = "" '"SELECT max(docentry) FROM OPCH T0  WHERE usersign=" & oCompany.UserSignature
        Dim APDE As String = ""
        Dim kCardCode As String = ""

        Select Case xForm.Mode
            Case SAPbouiCOM.BoFormMode.fm_ADD_MODE
                APDE = DocEntry
                SBO_Application.SetStatusBarMessage("Add Mode for " & DocNum & " <" & DocEntry & ">", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                'xRS.DoQuery(sql)
                'If Not xRS.EoF Then
                '    APDE = xRS.Fields.Item(0).Value
                '    '### Generating AR Invoice Draft ###
                '    SBO_Application.SetStatusBarMessage("Add Mode for " & APDE, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                'End If
            Case SAPbouiCOM.BoFormMode.fm_UPDATE_MODE, SAPbouiCOM.BoFormMode.fm_OK_MODE
                APDE = xForm.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0).ToString
                SBO_Application.SetStatusBarMessage("Update Mode for " & APDE, SAPbouiCOM.BoMessageTime.bmt_Short, False)
        End Select

        If APDE <> "" Then
            If CheckARINVdrf(APDE) Then
                'Get BP Group
                If APBPCheck(APDE) Then
                    xRS = getGrpAPBP(APDE)
                    While Not xRS.EoF
                        kCardCode = xRS.Fields.Item(0).Value
                        AddARInvDrf_byBP(APDE, kCardCode)
                        xRS.MoveNext()
                    End While
                Else
                    'Some Ref BP are missing 
                    SBO_Application.MessageBox("Some rows are having invalid/blank User Defined BPCodes")
                End If
                'kCardCode = getBP_ARINV()
                'If kCardCode = "" Then
                '    SBO_Application.MessageBox("Default Customer for AR Invoice Draft not Found!")
                'Else
                '    AddARInvDrf(APDE, kCardCode)
                'End If
            Else
                SBO_Application.MessageBox("AR Draft based on the current AP was generated.")
            End If
        End If
    End Sub
    Private Function getGrpAPBP(ByVal DocEntry As String, Optional ByVal isDraft As Boolean = False) As SAPbobsCOM.Recordset
        Dim vRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim sql As String = ""
        If isDraft Then
            sql = "select U_TWM_BPCODE, Isnull(CardCode,'') BP  from DRF1 left outer join OCRD on U_TWM_BPCode = CardCode "
            sql &= " where DRF1.DocEntry = '" & DocEntry & "' Group by U_TWM_BPCode, CardCode"
        Else
            sql = "select U_TWM_BPCODE, Isnull(CardCode,'') BP  from PCH1 left outer join OCRD on U_TWM_BPCode = CardCode "
            sql &= " where PCH1.DocEntry = '" & DocEntry & "' Group by U_TWM_BPCode, CardCode"
        End If
        vRS.DoQuery(sql)
        Return vRS
    End Function
    Private Function APBPCheck(ByVal DocEntry As String, Optional ByVal isDraft As Boolean = False) As Boolean
        Dim xRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim sql As String = ""
        If isDraft Then
            sql = "Select U_TWM_BPCODE from DRF1 where DRF1.DocEntry = '" & Docentry & "'"
        Else
            sql = "select U_TWM_BPCODE from PCH1 where PCH1.DocEntry = '" & Docentry & "'"
        End If
        sql &= "And Not Exists ( Select CardCode from OCRD where CardCode = U_TWM_BPCode )"
        xRS.DoQuery(sql)
        Return xRS.EoF
    End Function
    Private Sub Filter_APDrf_ARINVDrf(ByVal APDE As String, ByVal APDN As String)
        '### Generating AR Invoice Draft ###
        SBO_Application.SetStatusBarMessage("Add Mode for DE : " & APDE, SAPbouiCOM.BoMessageTime.bmt_Short, False)

        If CheckARINVdrf(APDE, APDN) Then
            'Get BP Group
            If APBPCheck(APDE, True) Then
                Dim kCardCode As String
                Dim xRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                xRS = getGrpAPBP(APDE, True)
                While Not xRS.EoF
                    kCardCode = xRS.Fields.Item(0).Value
                    AddARInvDrf_byBP(APDE, kCardCode, True)
                    xRS.MoveNext()
                End While
            Else
                'Some Ref BP are missing 
                SBO_Application.MessageBox("Some rows are having invalid/blank User Defined BPCodes")
            End If

            'Dim kCardCode As String = getBP_ARINV()
            'If kCardCode = "" Then
            '    SBO_Application.MessageBox("Default Customer for AR Invoice Draft not Found!")
            'Else
            '    AddARInvDrf(APDE, kCardCode, True)
            'End If
        Else
            SBO_Application.MessageBox("AR Draft based on the current AP was already generated.")
        End If
    End Sub

    Private Function SetInterCompanyConnectionContext() As Integer
        SendCmp = New SAPbobsCOM.Company
        Dim sCookie As String = SendCmp.GetContextCookie
        Dim sConnectionContext As String = SBO_Application.Company.GetConnectionContext(sCookie)
        If SendCmp.Connected = True Then
            SendCmp.Disconnect()
        End If
        SetInterCompanyConnectionContext = SendCmp.SetSboLoginContext(sConnectionContext) '// Set the connection context information to the DI API.
    End Function
    Public Function ConnectToSelectedCompanyDB(ByVal CompanyName As String) As Boolean

        Dim err As Long
        Dim errmsg As String = ""

        Try
            SBO_Application.SetStatusBarMessage("Please wait while Connecting to Target Database ...", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
            GC.Collect()
            If SendCmp.Connected = True Then
                SendCmp.Disconnect()
            End If
            GC.Collect()
            If SetInterCompanyConnectionContext() <> 0 Then
                SBO_Application.MessageBox("Failed setting a connection to DI API")
                End ' Terminating the Add-On Application
            End If

            'Get UserName and Pasword
            Dim sUN As String = GetSetting("ICCRED_UN")
            Dim sPW As String = GetSetting("ICCRED_PW")
            Dim oLic As New TWM_Licence.TWM_SAP(Key)
            If sPW <> "" Then
                Try
                    sPW = oLic.Decrypt(sPW)
                Catch ex As Exception
                    sPW = ""
                End Try
            End If

            SendCmp.UserName = sUN
            SendCmp.Password = sPW
            SendCmp.CompanyDB = CompanyName
            err = SendCmp.Connect
            GC.Collect()
            If err <> 0 Then
                GC.Collect()
                SendCmp.GetLastError(err, errmsg)
                SBO_Application.MessageBox(err & ": " & errmsg)
                Return False
            Else
                GC.Collect()
                Return True
            End If

        Catch ex As Exception
            MessageBox.Show(ex.StackTrace & ": " & ex.Message)
            Return False
        End Try
    End Function

    Private Sub Filter_IC_AR_PODrf(ByVal xForm As SAPbouiCOM.Form)

        Dim xRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim sql As String = "" '"SELECT max(docentry) FROM OINV T0  WHERE usersign=" & oCompany.UserSignature
        Dim ARDE As String = ""
        Dim jCardCode As String = ""
        Dim kCardCode As String = ""
        Dim kDBName As String = ""

        Select Case xForm.Mode
            Case SAPbouiCOM.BoFormMode.fm_ADD_MODE
                ARDE = DocEntry
                SBO_Application.SetStatusBarMessage("Add Mode for " & DocNum & " <" & DocEntry & ">", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                'xRS.DoQuery(sql)
                'If Not xRS.EoF Then
                '    ARDE = xRS.Fields.Item(0).Value
                '    '### Generating AR Invoice Draft ###
                '    SBO_Application.SetStatusBarMessage("Add Mode for " & ARDE, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                'End If
            Case SAPbouiCOM.BoFormMode.fm_UPDATE_MODE, SAPbouiCOM.BoFormMode.fm_OK_MODE
                ARDE = xForm.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0).ToString
                SBO_Application.SetStatusBarMessage("Update Mode for " & ARDE, SAPbouiCOM.BoMessageTime.bmt_Short, False)
        End Select

        If ARDE <> "" Then
            xRS = CheckARBP(ARDE)
            If xRS.EoF Then
                SBO_Application.MessageBox("AR Invoice Customer is not an Intercompany Customer!")
            Else
                kDBName = xRS.Fields.Item(0).Value 'Get Target DB Name 
                If ConnectToSelectedCompanyDB(kDBName) Then 'Check DB Connection
                    '## Start Intercompany Checks
                    SBO_Application.SetStatusBarMessage("Start Intercompany Checks", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                    'Check Default BP
                    xRS = CheckICBP(SendCmp, vCmp.CompanyDB) 'Check IC BP
                    If xRS.EoF Then
                        SBO_Application.SetStatusBarMessage("No Intercompany BP in Target Database", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Else
                        kCardCode = xRS.Fields.Item(0).Value
                        SBO_Application.SetStatusBarMessage("Intercompany Bp is " & kCardCode, SAPbouiCOM.BoMessageTime.bmt_Short, False)

                        If CheckICDoc(SendCmp, ARDE) Then 'Check Existing Doc
                            If ocICGetFieldID(SendCmp, "OITM", "TWM_FAB") = -1 Then
                                SBO_Application.SetStatusBarMessage("Target Company missing UDF U_TWM_FAB", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            Else
                                SBO_Application.SetStatusBarMessage("Intercompany Item Checking Start now", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                If preProcessITEM(SendCmp, ARDE) Then 'Check Items ( List )
                                    AddICPODrf(SendCmp, ARDE, kCardCode)
                                Else
                                    SBO_Application.SetStatusBarMessage("Intercompany Item Creation Failed", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                End If
                            End If
                        Else
                            SBO_Application.MessageBox("PO Draft based on the current AR was already existed in Intercompany database.")
                        End If
                    End If
                Else
                    SBO_Application.SetStatusBarMessage("Target Database Not Found", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                End If
            End If
        End If
    End Sub
    Private Sub myButton(ByVal xForm As SAPbouiCOM.Form, ByVal bID As String, ByVal refID As String, Optional ByVal AddFlag As Boolean = False, Optional ByVal bName As String = "")
        Dim refItem, myItem As SAPbouiCOM.Item
        Dim oButton As SAPbouiCOM.Button

        refItem = xForm.Items.Item(refID)

        If AddFlag Then
            myItem = xForm.Items.Add(bID, SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oButton = myItem.Specific
            oButton.Caption = bName
            myItem.Width = 150
            myItem.Height = refItem.Height
        Else
            myItem = xForm.Items.Item(bID)
        End If

        myItem.Top = refItem.Top
        myItem.Left = refItem.Left + refItem.Width + 5
    End Sub
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Dim oForm As SAPbouiCOM.Form
        Dim fState As String = ""
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing
        Try
            If pVal.BeforeAction = True Then
                Select pVal.FormTypeEx
                    Case "twmICCRE"
                        If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            oForm = SBO_Application.Forms.Item(FormUID)
                            SaveCredential(oForm)
                        End If

                End Select
            Else

                Select Case pVal.FormTypeEx
                    Case 141 'AP Inv 
                        'Case 141 'AP Invoice for CompanyA INSTEAD of 'Case 142 'PO
                        If CompanyA Then
                            oForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                            Select Case pVal.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                    myButton(oForm, "btnAP", "2", True, "Generate AR Invoice Draft")
                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                                    myButton(oForm, "btnAP", "2")
                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    oForm = SBO_Application.Forms.Item(FormUID)
                                    If pVal.ItemUID = "btnAP" Then
                                        Select Case oForm.Mode
                                            Case SAPbouiCOM.BoFormMode.fm_ADD_MODE
                                                'Get Document Status
                                                fState = oForm.Items.Item("81").Specific.selected.value
                                                If fState = "6" Then 'It's Draft
                                                    'Go Back to Draft Report
                                                    SBO_Application.MessageBox("Go Back to Draft Report to generate AR Invoice Draft from AP Invoice Draft")
                                                Else
                                                    SBO_Application.MessageBox("Adding AP INVOICE will generate AR Invoice Draft Automatically.")
                                                End If
                                            Case SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                SBO_Application.MessageBox("AR Invoice Draft Cannot be Updated.")
                                            Case SAPbouiCOM.BoFormMode.fm_OK_MODE
                                                SBO_Application.SetStatusBarMessage("Attempt Adding Again", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                Filter_AP_ARINVDrf(oForm)
                                            Case Else
                                                'Do Nothing
                                        End Select
                                    ElseIf pVal.ItemUID = "1" Then
                                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.ActionSuccess Then
                                            'SBO_Application.MessageBox("Adding AR Inv Draft")
                                            Filter_AP_ARINVDrf(oForm)
                                        End If
                                    End If
                            End Select
                        End If
                        '############################################################################################################
                        '###                    REMOVE Good Issue Function  by  Request ( 2013 December )                         ###
                        '############################################################################################################
                        'If Not CompanyA Then 'ONLY FOR COMPANY B
                        '    oForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                        '    Select Case pVal.EventType
                        '        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        '            myButton(oForm, "btnAPINV", "2", True, "Generate Good Issue")
                        '        Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                        '            myButton(oForm, "btnAPINV", "2")
                        '        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        '            oForm = SBO_Application.Forms.Item(FormUID)
                        '            If pVal.ItemUID = "btnAPINV" Then
                        '                Select Case oForm.Mode
                        '                    Case SAPbouiCOM.BoFormMode.fm_ADD_MODE
                        '                        SBO_Application.MessageBox("Adding AP Invoice will generate Good Issue Document Automatically.")
                        '                    Case SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        '                        SBO_Application.MessageBox("Good Issue Document Cannot be Updated.")
                        '                    Case SAPbouiCOM.BoFormMode.fm_OK_MODE
                        '                        'SBO_Application.SetStatusBarMessage("Attempt to do Good Issue Again", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                        '                        'Filter_APInv_GoodIssue(oForm)
                        '                    Case Else
                        '                        'Do Nothing
                        '                End Select
                        '            ElseIf pVal.ItemUID = "1" Then
                        '                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.ActionSuccess Then
                        '                    Filter_APInv_GoodIssue(oForm)
                        '                End If
                        '            End If
                        '    End Select
                        'End If
                        '############################################################################################################
                        '############################################################################################################
                    Case 181 'APCN
                        'If Not CompanyA Then 'ONLY FOR COMPANY B
                        '    oForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                        '    Select Case pVal.EventType
                        '        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        '            myButton(oForm, "btnAPCN", "2", True, "Generate Good Received")
                        '        Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                        '            myButton(oForm, "btnAPCN", "2")
                        '        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        '            oForm = SBO_Application.Forms.Item(FormUID)
                        '            If pVal.ItemUID = "btnAPCN" Then
                        '                Select Case oForm.Mode
                        '                    Case SAPbouiCOM.BoFormMode.fm_ADD_MODE
                        '                        SBO_Application.MessageBox("Adding APCN will generate Good Received Document Automatically.")
                        '                    Case SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        '                        SBO_Application.MessageBox("Good Received Document Cannot be Updated.")
                        '                    Case SAPbouiCOM.BoFormMode.fm_OK_MODE
                        '                        SBO_Application.SetStatusBarMessage("Attempt to do Good Received Again", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                        '                        Filter_APCN_GoodReceived(oForm)
                        '                    Case Else
                        '                        'Do Nothing
                        '                End Select
                        '            ElseIf pVal.ItemUID = "1" Then
                        '                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.ActionSuccess Then
                        '                    Filter_APCN_GoodReceived(oForm)
                        '                End If
                        '            End If
                        '    End Select
                        'End If
                    Case 3002 'Draft Report
                        If CompanyA Then
                            oForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                            Select Case pVal.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                    myButton(oForm, "btnAP", "2", True, "Generate AR Invoice Draft")
                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                                    myButton(oForm, "btnAP", "2")
                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    oForm = SBO_Application.Forms.Item(FormUID)
                                    If pVal.ItemUID = "btnAP" Then
                                        oMatrix = oForm.Items.Item("3").Specific
                                        Dim HiLine As Integer = oMatrix.GetNextSelectedRow
                                        If HiLine > -1 Then
                                            Dim xBox As SAPbouiCOM.EditText = oMatrix.Columns.Item("1").Cells.Item(HiLine).Specific
                                            Dim DE As Integer = xBox.Value
                                            xBox = oMatrix.Columns.Item("3").Cells.Item(HiLine).Specific
                                            Dim DN As Integer = xBox.Value

                                            xBox = oMatrix.Columns.Item("2").Cells.Item(HiLine).Specific
                                            If xBox.Value = 18 Then
                                                'SBO_Application.MessageBox("Line:" & HiLine & " Obj:" & xCbo.Selected.Value & " " & DE & "/" & DN)
                                                Filter_APDrf_ARINVDrf(DE, DN)
                                            Else
                                                SBO_Application.MessageBox("Highlighted Line is not a Purchase Invoice.")
                                            End If

                                            'Dim xCbo As SAPbouiCOM.ComboBox = oMatrix.Columns.Item("2").Cells.Item(HiLine).Specific
                                            'If xCbo.Selected.Value = 18 Then
                                            '    'SBO_Application.MessageBox("Line:" & HiLine & " Obj:" & xCbo.Selected.Value & " " & DE & "/" & DN)
                                            '    Filter_APDrf_ARINVDrf(DE, DN)
                                            'Else
                                            '    SBO_Application.MessageBox("Highlighted Line is not a Purchase Invoice.")
                                            'End If
                                        Else
                                            SBO_Application.MessageBox("Choose AP Invoice Draft to Generate AR Invoice Draft")
                                        End If

                                    End If
                            End Select
                        End If
                    Case 133 'AR INV
                        If CompanyA Then
                            oForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                            Select Case pVal.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                    myButton(oForm, "btnAR", "2", True, "Generate InterCo PO Draft")
                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                                    myButton(oForm, "btnAR", "2")
                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    oForm = SBO_Application.Forms.Item(FormUID)
                                    If pVal.ItemUID = "btnAR" Then
                                        Select Case oForm.Mode
                                            Case SAPbouiCOM.BoFormMode.fm_ADD_MODE
                                                SBO_Application.MessageBox("Adding AR Invoice will generate Inter Company PO Draft Automatically.")
                                            Case SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                SBO_Application.MessageBox("Inter Company PO Draft Cannot be Updated.")
                                            Case SAPbouiCOM.BoFormMode.fm_OK_MODE
                                                SBO_Application.SetStatusBarMessage("Attempt Adding Again", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                Filter_IC_AR_PODrf(oForm)
                                            Case Else
                                                'Do Nothing
                                        End Select
                                    ElseIf pVal.ItemUID = "1" Then
                                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.ActionSuccess Then
                                            'SBO_Application.MessageBox("Adding AR Inv Draft")
                                            Filter_IC_AR_PODrf(oForm)
                                        End If
                                    End If
                            End Select
                        End If
                End Select
            End If
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex.Message)
        End Try
    End Sub

    Private Sub SaveCredential(ByRef oForm As SAPbouiCOM.Form)
        Dim sUN As String = oForm.Items.Item("txtUN").Specific.String
        Dim sPlainPW As String = oForm.Items.Item("txtPW").Specific.String
        Dim sEnryptedPW As String = ""
        If (sPlainPW <> "") Then
            Dim oLic As New TWM_Licence.TWM_SAP(Key)
            Try
                sEnryptedPW = oLic.Encrypt(sPlainPW)
            Catch ex As Exception
            End Try
        End If

        'Write to database.

        Dim sSQL As String = ""
        Dim oRS As SAPbobsCOM.Recordset = Nothing
        Try
            oCompany.StartTransaction()

            'UN
            UpdateSetting("ICCRED_UN", sUN)

            'PW
            UpdateSetting("ICCRED_PW", sEnryptedPW)
            

            If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
        Catch ex As Exception
            Throw ex
        Finally
            If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
        End Try


    End Sub

    Private Sub UpdateSetting(ByVal Name As String, ByVal Value As String)
        Dim sSQL As String = String.Format("SELECT Code FROM [@TWM_SS] WHERE Name = '{0}'", Name)
        Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRS.DoQuery(sSQL)
        If oRS.RecordCount > 0 Then
            'Update
            Dim sCode As String = oRS.Fields.Item(0).Value
            sSQL = String.Format("UPDATE [@TWM_SS] SET U_SValue = '{1}' WHERE Code = '{0}'", sCode, Value)
            oRS.DoQuery(sSQL)
        Else
            'Insert
            sSQL = "SELECT ISNULL(MAX(CONVERT(INT,Code)),0)+1 FROM [@TWM_SS] WHERE CODE LIKE '%[1-9]%'"
            oRS.DoQuery(sSQL)
            Dim sCode As String = oRS.Fields.Item(0).Value
            sSQL = String.Format("INSERT INTO [@TWM_SS] (Code, Name, U_SValue) VALUES('{0}', '{1}', '{2}')", sCode, Name, Value)
            oRS.DoQuery(sSQL)
        End If
    End Sub

    Private Function GetSetting(ByVal Name As String) As String
        Dim sSQL As String = String.Format("SELECT U_SValue FROM [@TWM_SS] WHERE Name = '{0}'", Name.Replace("'", "''"))
        Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRS.DoQuery(sSQL)
        If (oRS.Fields.Count > 0) Then
            Return oRS.Fields.Item(0).Value
        Else
            Return Nothing
        End If
    End Function

#End Region


#Region "helper method"
    Private Function DrawForm(ByRef FileName As String) As SAPbouiCOM.Form
        Dim oForm As SAPbouiCOM.Form = Nothing
        Try
            LoadFromXML(FileName & ".xml")
            oForm = SBO_Application.Forms.ActiveForm
            If (oForm.TypeEx <> FileName) Then Throw New Exception("Invalid Form")
            InitForm(oForm)


        Catch ex As Exception
            Throw ex
        End Try


    End Function
    Private Sub InitForm(ByRef oForm As SAPbouiCOM.Form)
        Select Case oForm.TypeEx
            Case "twmICCRE"
                oForm.DataSources.UserDataSources.Add("txtUN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50)
                oForm.DataSources.UserDataSources.Add("txtPW", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
                Dim oText As SAPbouiCOM.EditText = oForm.Items.Item("txtUN").Specific
                oText.DataBind.SetBound(True, "", "txtUN")
                oText = oForm.Items.Item("txtPW").Specific
                oText.DataBind.SetBound(True, "", "txtPW")

                Dim sSQL As String = "SELECT U_SValue FROM [@TWM_SS] WHERE Name = 'ICCRED_UN'"
                Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRS.DoQuery(sSQL)
                If oRS.RecordCount > 0 Then oForm.Items.Item("txtUN").Specific.String = oRS.Fields.Item(0).Value

                sSQL = "SELECT U_SValue FROM [@TWM_SS] WHERE Name = 'ICCRED_PW'"
                oRS.DoQuery(sSQL)
                Dim PlainPW As String = ""
                If oRS.RecordCount > 0 Then
                    Dim EncryptedPW As String = oRS.Fields.Item(0).Value
                    If EncryptedPW <> "" Then
                        Try
                            Dim oLic As New TWM_Licence.TWM_SAP(Key)
                            PlainPW = oLic.Decrypt(EncryptedPW)
                        Catch ex As Exception
                            PlainPW = ""
                        End Try
                    End If
                End If
                oForm.Items.Item("txtPW").Specific.String = PlainPW

                oForm.ActiveItem = "txtUN"


        End Select
    End Sub
    Private Sub LoadFromXML(ByRef FileName As String)

        'Dim oXmlDoc As Xml.XmlDocument

        'oXmlDoc = New Xml.XmlDocument

        ''// load the content of the XML File
        'Dim sPath As String

        'sPath = Application.StartupPath.ToString

        'oXmlDoc.Load(sPath & "\" & FileName)

        Dim sResourceName As String = String.Format("{0}.{1}", System.Reflection.Assembly.GetExecutingAssembly.GetName.Name, FileName)
        Dim oStream As System.IO.Stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(sResourceName)
        If Not oStream Is Nothing Then
            '// load the form to the SBO application in one batch
            Dim sXML As String = ""
            Using oReader As New IO.StreamReader(oStream)
                sXML = oReader.ReadToEnd
            End Using
            SBO_Application.LoadBatchActions(sXML)
            Dim sResult As String = SBO_Application.GetLastBatchResults()
        End If

        

    End Sub
#End Region

    Private Sub SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.FormDataEvent
        If BusinessObjectInfo.ActionSuccess Then
            If BusinessObjectInfo.BeforeAction = False Then
                If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                    Dim xForm As SAPbouiCOM.Form = SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                    Dim bisObj As SAPbouiCOM.BusinessObject = xForm.BusinessObject
                    Dim uid As String = bisObj.Key
                    Dim oDocument As SAPbobsCOM.Documents = Nothing
                    Dim goFlag As Boolean = True
                    Select Case BusinessObjectInfo.Type
                        Case SAPbobsCOM.BoObjectTypes.oPurchaseInvoices, SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes, SAPbobsCOM.BoObjectTypes.oInvoices
                            oDocument = oCompany.GetBusinessObject(BusinessObjectInfo.Type)
                        Case Else
                            goFlag = False
                    End Select
                    If goFlag Then
                        oDocument.Browser.GetByKeys(BusinessObjectInfo.ObjectKey)
                        DocEntry = oDocument.DocEntry
                        DocNum = oDocument.DocNum
                    End If
                End If
            End If
        End If
        
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
