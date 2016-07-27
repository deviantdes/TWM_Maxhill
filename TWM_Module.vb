Option Strict Off
Option Explicit On

Module TWM_Module
    Public vCmp As New SAPbobsCOM.Company
    Public SendCmp As New SAPbobsCOM.Company

    Public Sub SetToolBarEnabled(ByVal xform As SAPbouiCOM.Form)
        xform.EnableMenu("1281", True) 'Find
        xform.EnableMenu("1282", True) 'Add
        xform.EnableMenu("1283", True) 'Remove Menu
        xform.EnableMenu("1288", True) 'Next Rec
        xform.EnableMenu("1289", True) 'Previous Rec
        xform.EnableMenu("1290", True) 'First Rec
        xform.EnableMenu("1291", True) 'Last Rec
        xform.EnableMenu("6913", True)  'Disable UDF updating
    End Sub

    Public Sub SetToolBarDisabled(ByVal xform As SAPbouiCOM.Form)
        xform.EnableMenu("1281", False) 'Find
        xform.EnableMenu("1282", False) 'Add
        xform.EnableMenu("1283", False) 'Remove Menu
        xform.EnableMenu("1288", False) 'Next Rec
        xform.EnableMenu("1289", False) 'Previous Rec
        xform.EnableMenu("1290", False) 'First Rec
        xform.EnableMenu("1291", False) 'Last Rec
        xform.EnableMenu("6913", False)  'Disable UDF updating
    End Sub

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
    Private Function CreateUDOUOMType() As Boolean
        Dim kRet As Integer
        Dim kReturn As Boolean = False
        Dim kUDO As SAPbobsCOM.UserObjectsMD = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
        If kUDO.GetByKey("TWM_UOMTYPE") = True Then
            kReturn = True
        Else
            With kUDO
                'ocCheckUDO("TWM_UOMTYPE", "TWM_UOMTYPE", "TWM_UOMTYPE", SAPbobsCOM.BoUDOObjType.boud_MasterData, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO, "")
                .Code = "TWM_UOMTYPE"
                .Name = "TWM_UOMTYPE"
                .TableName = "TWM_UOMTYPE"
                .ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData

                .CanFind = SAPbobsCOM.BoYesNoEnum.tYES

                .FindColumns.ColumnAlias = "Code"
                .FindColumns.ColumnDescription = "UOM Code"
                .FindColumns.Add()

                .FindColumns.ColumnAlias = "Name"
                .FindColumns.ColumnDescription = "UOM Name"
                .FindColumns.Add()

                .FindColumns.ColumnAlias = "U_UOMQty"
                .FindColumns.ColumnDescription = "UOM Qty"
                .FindColumns.Add()

                .CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
                .CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
                .CanClose = SAPbobsCOM.BoYesNoEnum.tNO
                .ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO
                .CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                .CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES

                .FormColumns.FormColumnAlias = "Code"
                .FormColumns.FormColumnDescription = "UOM Code"
                .FormColumns.Add()

                .FormColumns.FormColumnAlias = "Name"
                .FormColumns.FormColumnDescription = "UOM Name"
                .FormColumns.Add()

                .FormColumns.FormColumnAlias = "U_UOMQty"
                .FormColumns.FormColumnDescription = "UOM Qty"
                .FormColumns.Add()

                .CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                .LogTableName = "" 'Not support Log Table now

                kRet = .Add
            End With
            errSub(kRet)
        End If
        kUDO = Nothing
        Return kReturn
    End Function
    Private Function ocCheckUDO(ByVal kUDOCode As String, ByVal kUDOName As String, ByVal kUDOTableName As String, ByVal kUDOType As SAPbobsCOM.BoUDOObjType, _
       ByVal kCanFind As SAPbobsCOM.BoYesNoEnum, ByVal kCanDelete As SAPbobsCOM.BoYesNoEnum, ByVal kCanCancel As SAPbobsCOM.BoYesNoEnum, _
       ByVal kCanClose As SAPbobsCOM.BoYesNoEnum, ByVal kManageSeries As SAPbobsCOM.BoYesNoEnum, ByVal kCanYrTransfer As SAPbobsCOM.BoYesNoEnum, _
       ByVal kCanDefaultForm As SAPbobsCOM.BoYesNoEnum, ByVal kCanLog As SAPbobsCOM.BoYesNoEnum, _
       ByVal kChildTableName As String) As Boolean
        Dim kRet As Integer
        Dim kReturn As Boolean = False

        Dim kUDO As SAPbobsCOM.UserObjectsMD = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
        If kUDO.GetByKey(kUDOCode) = True Then
            kReturn = True
        Else
            With kUDO
                .Code = kUDOCode
                .Name = kUDOName
                .TableName = kUDOTableName
                .ObjectType = kUDOType

                .CanFind = kCanFind
                'Ignore Find Columns
                .CanDelete = kCanDelete
                .CanCancel = kCanCancel
                .CanClose = kCanClose
                .ManageSeries = kManageSeries
                .CanYearTransfer = kCanYrTransfer
                .CanCreateDefaultForm = kCanDefaultForm
                .CanLog = kCanLog
                .LogTableName = "" 'Not support Log Table now

                If kChildTableName <> "" Then
                    .ChildTables.TableName = kChildTableName
                    .ChildTables.Add()
                End If

                kRet = .Add
            End With
            errSub(kRet)
        End If
        kUDO = Nothing
        Return kReturn
    End Function
    
   
End Module


