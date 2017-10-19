Imports System.Data.SqlClient

Public Class ImportData

    Dim ClsClobalFunction As New GlobalFunction
    Dim ClsEncrypDecryp As New EncrypDecryp
    Dim ClsConn As New ConnectDatabase
    Dim ConnStr As String = ClsConn.ConnStr

    Public Function InsertData(TableName As String, dt As DataTable) As String
        Dim conn As New SqlConnection(ConnStr)
        conn.Open()
        Dim trans As SqlTransaction
        trans = conn.BeginTransaction

        'Dim gmodby As String = "IEReplaceCOCO"
        Dim sql As String = ""
        Try

            Select Case TableName
                Case "APP_DATA"
#Region "APP_DATA"
                    Dim AID, DEPOT, SITENAME, SITEADD, SITETEL, SITEFAX, SITEZIPCODE, VATNO, LOCALVAT, IS_COCO, IS_POSCASH, BUS_PLACE, VOLUMEFORMAT, VALUEFORMAT, SITENAME2,
                        C_SERVICE, COM_NAME, COM_BRANCH, LOCAL_DIFFERENCE, TOBACCO_TAX, CREATEDATE, MODDATE, MODBY As String


                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandText = sql
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        With dt
                            AID = IIf(.Rows(i)("AID").ToString = "", "null", "'" & .Rows(i)("AID").ToString & "'")
                            DEPOT = IIf(.Rows(i)("DEPOT").ToString = "", "null", "'" & .Rows(i)("DEPOT").ToString & "'")
                            SITENAME = IIf(.Rows(i)("SITENAME").ToString = "", "null", "'" & .Rows(i)("SITENAME").ToString & "'")
                            SITEADD = IIf(.Rows(i)("SITEADD").ToString = "", "null", "'" & .Rows(i)("SITEADD").ToString & "'")
                            SITETEL = IIf(.Rows(i)("SITETEL").ToString = "", "null", "'" & .Rows(i)("SITETEL").ToString & "'")
                            SITEFAX = IIf(.Rows(i)("SITEFAX").ToString = "", "null", "'" & .Rows(i)("SITEFAX").ToString & "'")
                            SITEZIPCODE = IIf(.Rows(i)("SITEZIPCODE").ToString = "", "null", "'" & .Rows(i)("SITEZIPCODE").ToString & "'")
                            VATNO = IIf(.Rows(i)("VATNO").ToString = "", "null", "'" & .Rows(i)("VATNO").ToString & "'")
                            LOCALVAT = IIf(.Rows(i)("LOCALVAT").ToString = "", "null", "'" & .Rows(i)("LOCALVAT").ToString & "'")
                            IS_COCO = IIf(.Rows(i)("IS_COCO").ToString = "", "null", "'" & .Rows(i)("IS_COCO").ToString & "'")
                            IS_POSCASH = IIf(.Rows(i)("IS_POSCASH").ToString = "", "null", "'" & .Rows(i)("IS_POSCASH").ToString & "'")
                            BUS_PLACE = IIf(.Rows(i)("BUS_PLACE").ToString = "", "null", "'" & .Rows(i)("BUS_PLACE").ToString & "'")
                            VOLUMEFORMAT = IIf(.Rows(i)("VOLUMEFORMAT").ToString = "", "null", "'" & .Rows(i)("VOLUMEFORMAT").ToString & "'")
                            VALUEFORMAT = IIf(.Rows(i)("VALUEFORMAT").ToString = "", "null", "'" & .Rows(i)("VALUEFORMAT").ToString & "'")
                            SITENAME2 = IIf(.Rows(i)("SITENAME2").ToString = "", "null", "'" & .Rows(i)("SITENAME2").ToString & "'")
                            C_SERVICE = IIf(.Rows(i)("C_SERVICE").ToString = "", "null", "'" & .Rows(i)("C_SERVICE").ToString & "'")
                            COM_NAME = "null" 'IIf(.Rows(i)("COM_NAME").ToString = "", "null", "'" & .Rows(i)("COM_NAME").ToString & "'")
                            COM_BRANCH = "null" 'IIf(.Rows(i)("COM_BRANCH").ToString = "", "null", "'" & .Rows(i)("COM_BRANCH").ToString & "'")
                            LOCAL_DIFFERENCE = "0" 'IIf(.Rows(i)("LOCAL_DIFFERENCE").ToString = "", "null", "'" & .Rows(i)("LOCAL_DIFFERENCE").ToString & "'")
                            TOBACCO_TAX = "7" 'IIf(.Rows(i)("TOBACCO_TAX").ToString = "", "null", "'" & .Rows(i)("TOBACCO_TAX").ToString & "'")
                        End With
                        CREATEDATE = IIf(dt.Rows(i)("CREATEDATE").ToString = "", "null", "" & dt.Rows(i)("CREATEDATE").ToString & "")
                        If CREATEDATE <> "null" Then
                            CREATEDATE = ClsClobalFunction.ConvertDateTime(CREATEDATE)
                        End If

                        MODDATE = IIf(dt.Rows(i)("MODDATE").ToString = "", "null", "" & dt.Rows(i)("MODDATE").ToString & "")
                        If MODDATE <> "null" Then
                            MODDATE = ClsClobalFunction.ConvertDateTime(MODDATE)
                        End If
                        MODBY = IIf(dt.Rows(i)("MODBY").ToString = "", "null", "'" & dt.Rows(i)("MODBY").ToString & "'")


                        sql = "INSERT INTO [dbo].[APP_DATA] "
                        sql &= " ([AID] "
                        sql &= " ,[DEPOT]"
                        sql &= " ,[SITENAME]"
                        sql &= " ,[SITEADD]"
                        sql &= " ,[SITETEL]"
                        sql &= " ,[SITEFAX]"
                        sql &= " ,[SITEZIPCODE]"
                        sql &= " ,[VATNO]"
                        sql &= " ,[LOCALVAT]"
                        sql &= " ,[IS_COCO]"
                        sql &= " ,[IS_POSCASH]"
                        sql &= " ,[CREATEDATE]"
                        sql &= " ,[MODDATE]"
                        sql &= " ,[MODBY]"
                        sql &= " ,[BUS_PLACE]"
                        sql &= " ,[VOLUMEFORMAT]"
                        sql &= " ,[VALUEFORMAT]"
                        sql &= " ,[SITENAME2]"
                        sql &= " ,[C_SERVICE]"
                        sql &= " ,[COM_NAME]"
                        sql &= " ,[COM_BRANCH]"
                        sql &= " ,[LOCAL_DIFFERENCE]"
                        sql &= " ,[TOBACCO_TAX])"
                        sql &= "  VALUES"
                        sql &= "  (" & AID & ""
                        sql &= "  ," & DEPOT & ""
                        sql &= "  ," & SITENAME & ""
                        sql &= "  ," & SITEADD & ""
                        sql &= "  ," & SITETEL & ""
                        sql &= "  ," & SITEFAX & ""
                        sql &= "  ," & SITEZIPCODE & ""
                        sql &= "  ," & VATNO & ""
                        sql &= "  ," & LOCALVAT & ""
                        sql &= "  ," & IS_COCO & ""
                        sql &= "  ," & IS_POSCASH & ""
                        sql &= "  ," & CREATEDATE & ""
                        sql &= "  ," & MODDATE & ""
                        sql &= "  ," & MODBY & ""
                        sql &= "  ," & BUS_PLACE & ""
                        sql &= "  ," & VOLUMEFORMAT & ""
                        sql &= "  ," & VALUEFORMAT & ""
                        sql &= "  ," & SITENAME2 & ""
                        sql &= "  ," & C_SERVICE & ""
                        sql &= "  ," & COM_NAME & ""
                        sql &= "  ," & COM_BRANCH & ""
                        sql &= "  ," & LOCAL_DIFFERENCE & ""
                        sql &= "  ," & TOBACCO_TAX & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()

                    Next
#End Region

                Case "TBPOS_PUMP_ALLOW"
#Region "TBPOS_PUMP_ALLOW"
                    Dim POS_ID, PUMP_ID, CREATEDATE, MODDATE, MODBY As String

                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandText = sql
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        POS_ID = dt.Rows(i)("POS_ID").ToString
                        PUMP_ID = dt.Rows(i)("PUMP_ID").ToString

                        CREATEDATE = IIf(dt.Rows(i)("CREATEDATE").ToString = "", "null", "" & dt.Rows(i)("CREATEDATE").ToString & "")
                        If CREATEDATE <> "null" Then
                            CREATEDATE = ClsClobalFunction.ConvertDateTime(CREATEDATE)
                        End If

                        MODDATE = IIf(dt.Rows(i)("MODDATE").ToString = "", "null", "" & dt.Rows(i)("MODDATE").ToString & "")
                        If MODDATE <> "null" Then
                            MODDATE = ClsClobalFunction.ConvertDateTime(MODDATE)
                        End If
                        MODBY = IIf(dt.Rows(i)("MODBY").ToString = "", "null", "'" & dt.Rows(i)("MODBY").ToString & "'")

                        sql = " INSERT INTO [dbo].[TBPOS_PUMP_ALLOW]"
                        sql &= " ([POS_ID],[PUMP_ID],CREATEDATE,MODDATE,MODBY)"
                        sql &= " VALUES(" & IIf(POS_ID = "", "null", "'" & POS_ID & "'") & ""
                        sql &= "," & IIf(PUMP_ID = "", "null", "'" & PUMP_ID & "'") & ""
                        sql &= "," & CREATEDATE & "," & MODDATE & "," & MODBY & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()

                    Next
#End Region

                Case "TBMATERIAL"
#Region "TBMATERIAL"
                    Dim MAT_ID, MAT_NAME, MAT_ID2, MAT_NAME2, MAT_NAME3, MAT_BARCODE, QTY, UOM, MOVING_AVG_PRICE, STOCK, STOCK_MIN, STOCK_MAX, STOCK_LOCATION_ID, TAX_CLASS, MAT_GROUP, MAT_GROUP3,
                                          DIVISION_ID, PRICE0, PRICE1, PRICE2, PRICE3, PRICE4, PRICE5, PRICE6, PRICE7, PRICE8, PRICE9, PRICE10, PRICE11, PRICE12, TIMEOFSALE, LAST_SALE, LAST_RECEIVE, BLOCK,
                                          PRICINGDATE, PRICINGMODBY, LOCATION_ID, MATCOLOR, OBJ_ID, OBJ_ID_MAT_GROUP3, OBJ_ID_DIVISION_ID, CREATEDATE, MODDATE, MODBY As String

                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        With dt
                            MAT_ID = IIf(.Rows(i)("MAT_ID").ToString = "", "null", "'" & .Rows(i)("MAT_ID").ToString & "'")
                            MAT_NAME = IIf(.Rows(i)("MAT_NAME").ToString = "", "null", "'" & .Rows(i)("MAT_NAME").ToString & "'")
                            MAT_ID2 = IIf(.Rows(i)("MAT_ID2").ToString = "", "null", "'" & .Rows(i)("MAT_ID2").ToString & "'")
                            MAT_NAME2 = IIf(.Rows(i)("MAT_NAME2").ToString = "", "null", "'" & .Rows(i)("MAT_NAME2").ToString & "'")
                            MAT_NAME3 = IIf(.Rows(i)("MAT_NAME3").ToString = "", "null", "'" & .Rows(i)("MAT_NAME3").ToString & "'")
                            MAT_BARCODE = IIf(.Rows(i)("MAT_BARCODE").ToString = "", "null", "'" & .Rows(i)("MAT_BARCODE").ToString & "'")
                            QTY = IIf(.Rows(i)("QTY").ToString = "", "null", "'" & .Rows(i)("QTY").ToString & "'")
                            UOM = IIf(.Rows(i)("UOM").ToString = "", "null", "'" & .Rows(i)("UOM").ToString & "'")
                            MOVING_AVG_PRICE = IIf(.Rows(i)("MOVING_AVG_PRICE").ToString = "", "null", "'" & .Rows(i)("MOVING_AVG_PRICE").ToString & "'")
                            STOCK = IIf(.Rows(i)("STOCK").ToString = "", "null", "'" & .Rows(i)("STOCK").ToString & "'")
                            STOCK_MIN = IIf(.Rows(i)("STOCK_MIN").ToString = "", "null", "'" & .Rows(i)("STOCK_MIN").ToString & "'")
                            STOCK_MAX = IIf(.Rows(i)("STOCK_MAX").ToString = "", "null", "'" & .Rows(i)("STOCK_MAX").ToString & "'")
                            STOCK_LOCATION_ID = IIf(.Rows(i)("STOCK_LOCATION_ID").ToString = "", "null", "'" & .Rows(i)("STOCK_LOCATION_ID").ToString & "'")
                            TAX_CLASS = IIf(.Rows(i)("TAX_CLASS").ToString = "", "null", "'" & .Rows(i)("TAX_CLASS").ToString & "'")
                            MAT_GROUP = IIf(.Rows(i)("MAT_GROUP").ToString = "", "null", "'" & .Rows(i)("MAT_GROUP").ToString & "'")
                            MAT_GROUP3 = IIf(.Rows(i)("MAT_GROUP3").ToString = "", "null", "'" & .Rows(i)("MAT_GROUP3").ToString & "'")
                            DIVISION_ID = IIf(.Rows(i)("DIVISION_ID").ToString = "DIVISION_ID", "null", "'" & .Rows(i)("DIVISION_ID").ToString & "'")
                            PRICE0 = IIf(.Rows(i)("PRICE0").ToString = "", "null", "'" & .Rows(i)("PRICE0").ToString & "'")
                            PRICE1 = IIf(.Rows(i)("PRICE1").ToString = "", "null", "'" & .Rows(i)("PRICE1").ToString & "'")
                            PRICE2 = IIf(.Rows(i)("PRICE2").ToString = "", "null", "'" & .Rows(i)("PRICE2").ToString & "'")
                            PRICE3 = IIf(.Rows(i)("PRICE3").ToString = "", "null", "'" & .Rows(i)("PRICE3").ToString & "'")
                            PRICE4 = IIf(.Rows(i)("PRICE4").ToString = "", "null", "'" & .Rows(i)("PRICE4").ToString & "'")
                            PRICE5 = IIf(.Rows(i)("PRICE5").ToString = "", "null", "'" & .Rows(i)("PRICE5").ToString & "'")
                            PRICE6 = IIf(.Rows(i)("PRICE6").ToString = "", "null", "'" & .Rows(i)("PRICE6").ToString & "'")
                            PRICE7 = IIf(.Rows(i)("PRICE7").ToString = "", "null", "'" & .Rows(i)("PRICE7").ToString & "'")
                            PRICE8 = IIf(.Rows(i)("PRICE8").ToString = "", "null", "'" & .Rows(i)("PRICE8").ToString & "'")
                            PRICE9 = IIf(.Rows(i)("PRICE9").ToString = "", "null", "'" & .Rows(i)("PRICE9").ToString & "'")
                            PRICE10 = IIf(.Rows(i)("PRICE10").ToString = "", "null", "'" & .Rows(i)("PRICE10").ToString & "'")
                            PRICE11 = IIf(.Rows(i)("PRICE11").ToString = "", "null", "'" & .Rows(i)("PRICE11").ToString & "'")
                            PRICE12 = IIf(.Rows(i)("PRICE12").ToString = "", "null", "'" & .Rows(i)("PRICE12").ToString & "'")
                            TIMEOFSALE = IIf(.Rows(i)("TIMEOFSALE").ToString = "", "null", "'" & .Rows(i)("TIMEOFSALE").ToString & "'")

                            LAST_SALE = IIf(.Rows(i)("LAST_SALE").ToString = "", "null", "" & .Rows(i)("LAST_SALE").ToString & "")
                            If LAST_SALE <> "null" Then
                                LAST_SALE = ClsClobalFunction.ConvertDateTime(LAST_SALE)
                            End If

                            LAST_RECEIVE = IIf(.Rows(i)("LAST_RECEIVE").ToString = "", "null", "" & .Rows(i)("LAST_RECEIVE").ToString & "")
                            If LAST_RECEIVE <> "null" Then
                                LAST_RECEIVE = ClsClobalFunction.ConvertDate(LAST_RECEIVE)
                            End If

                            BLOCK = IIf(.Rows(i)("BLOCK").ToString = "", "null", "'" & .Rows(i)("BLOCK").ToString & "'")
                            PRICINGDATE = IIf(.Rows(i)("PRICINGDATE").ToString = "", "null", "" & .Rows(i)("PRICINGDATE").ToString & "")
                            If PRICINGDATE <> "null" Then
                                PRICINGDATE = ClsClobalFunction.ConvertDate(PRICINGDATE)
                            End If

                            PRICINGMODBY = IIf(.Rows(i)("PRICINGMODBY").ToString = "", "null", "'" & .Rows(i)("PRICINGMODBY").ToString & "'")
                            LOCATION_ID = IIf(.Rows(i)("LOCATION_ID").ToString = "", "null", "'" & .Rows(i)("LOCATION_ID").ToString & "'")
                            MATCOLOR = IIf(.Rows(i)("MATCOLOR").ToString = "", "null", "'" & .Rows(i)("MATCOLOR").ToString & "'")
                            OBJ_ID = "null" 'IIf(.Rows(i)("OBJ_ID").ToString = "", "null", "'" & .Rows(i)("OBJ_ID").ToString & "'")
                            OBJ_ID_MAT_GROUP3 = "null" 'IIf(.Rows(i)("OBJ_ID_MAT_GROUP3").ToString = "", "null", "'" & .Rows(i)("OBJ_ID_MAT_GROUP3").ToString & "'")
                            OBJ_ID_DIVISION_ID = "null" 'IIf(.Rows(i)("OBJ_ID_DIVISION_ID").ToString = "", "null", "'" & .Rows(i)("OBJ_ID_DIVISION_ID").ToString & "'")
                        End With

                        CREATEDATE = IIf(dt.Rows(i)("CREATEDATE").ToString = "", "null", "" & dt.Rows(i)("CREATEDATE").ToString & "")
                        If CREATEDATE <> "null" Then
                            CREATEDATE = ClsClobalFunction.ConvertDateTime(CREATEDATE)
                        End If

                        MODDATE = IIf(dt.Rows(i)("MODDATE").ToString = "", "null", "" & dt.Rows(i)("MODDATE").ToString & "")
                        If MODDATE <> "null" Then
                            MODDATE = ClsClobalFunction.ConvertDateTime(MODDATE)
                        End If

                        MODBY = IIf(dt.Rows(i)("MODBY").ToString = "", "null", "'" & dt.Rows(i)("MODBY").ToString & "'")

                        sql = "INSERT INTO [dbo].[TBMATERIAL]"
                        sql &= "([MAT_ID]"
                        sql &= "  , [MAT_NAME]"
                        sql &= "  , [MAT_ID2]"
                        sql &= "  , [MAT_NAME2]"
                        sql &= "  , [MAT_NAME3]"
                        sql &= "  , [MAT_BARCODE]"
                        sql &= "  , [QTY]"
                        sql &= "  , [UOM]"
                        sql &= "  , [MOVING_AVG_PRICE]"
                        sql &= "  , [STOCK]"
                        sql &= "  , [STOCK_MIN]"
                        sql &= "  , [STOCK_MAX]"
                        sql &= "  , [STOCK_LOCATION_ID]"
                        sql &= "  , [TAX_CLASS]"
                        sql &= "  , [MAT_GROUP]"
                        sql &= "  , [MAT_GROUP3]"
                        sql &= "  , [DIVISION_ID]"
                        sql &= "  , [PRICE0]"
                        sql &= "  , [PRICE1]"
                        sql &= "  , [PRICE2]"
                        sql &= "  , [PRICE3]"
                        sql &= "  ,[PRICE4]"
                        sql &= "  ,[PRICE5]"
                        sql &= "  ,[PRICE6]"
                        sql &= "  ,[PRICE7]"
                        sql &= "  ,[PRICE8]"
                        sql &= "  ,[PRICE9]"
                        sql &= "  ,[PRICE10]"
                        sql &= "  ,[PRICE11]"
                        sql &= "  ,[PRICE12]"
                        sql &= "  ,[TIMEOFSALE]"
                        sql &= "  ,[LAST_SALE]"
                        sql &= "  ,[LAST_RECEIVE]"
                        sql &= "  ,[BLOCK]"
                        sql &= "  ,[PRICINGDATE]"
                        sql &= "  ,[PRICINGMODBY]"
                        sql &= "  ,[LOCATION_ID]"
                        sql &= "  ,[MATCOLOR]"
                        sql &= "  ,[CREATEDATE]"
                        sql &= "  ,[MODDATE]"
                        sql &= "  ,[MODBY]"
                        sql &= "  ,[OBJ_ID]"
                        sql &= "  ,[OBJ_ID_MAT_GROUP3]"
                        sql &= "  ,[OBJ_ID_DIVISION_ID])"
                        sql &= "   VALUES"
                        sql &= "     (" & MAT_ID & ""
                        sql &= "     ," & MAT_NAME & ""
                        sql &= "     ," & MAT_ID2 & ""
                        sql &= "     ," & MAT_NAME2 & ""
                        sql &= "     ," & MAT_NAME3 & ""
                        sql &= "     ," & MAT_BARCODE & ""
                        sql &= "     ," & QTY & ""
                        sql &= "     ," & UOM & ""
                        sql &= "     ," & MOVING_AVG_PRICE & ""
                        sql &= "     ," & STOCK & ""
                        sql &= "     ," & STOCK_MIN & ""
                        sql &= "     ," & STOCK_MAX & ""
                        sql &= "     ," & STOCK_LOCATION_ID & ""
                        sql &= "     ," & TAX_CLASS & ""
                        sql &= "     ," & MAT_GROUP & ""
                        sql &= "     ," & MAT_GROUP3 & ""
                        sql &= "     ," & DIVISION_ID & ""
                        sql &= "     ," & PRICE0 & ""
                        sql &= "     ," & PRICE1 & ""
                        sql &= "     ," & PRICE2 & ""
                        sql &= "     ," & PRICE3 & ""
                        sql &= "     ," & PRICE4 & ""
                        sql &= "     ," & PRICE5 & ""
                        sql &= "     ," & PRICE6 & ""
                        sql &= "     ," & PRICE7 & ""
                        sql &= "     ," & PRICE8 & ""
                        sql &= "     ," & PRICE9 & ""
                        sql &= "     ," & PRICE10 & ""
                        sql &= "     ," & PRICE11 & ""
                        sql &= "     ," & PRICE12 & ""
                        sql &= "     ," & TIMEOFSALE & ""
                        sql &= "     ," & LAST_SALE & ""
                        sql &= "     ," & LAST_RECEIVE & ""
                        sql &= "     ," & BLOCK & ""
                        sql &= "     ," & PRICINGDATE & ""
                        sql &= "     ," & PRICINGMODBY & ""
                        sql &= "     ," & LOCATION_ID & ""
                        sql &= "     ," & MATCOLOR & ""
                        sql &= "     ," & CREATEDATE & ""
                        sql &= "     ," & MODDATE & ""
                        sql &= "     ," & MODBY & ""
                        sql &= "     ," & OBJ_ID & ""
                        sql &= "     ," & OBJ_ID_MAT_GROUP3 & ""
                        sql &= "     ," & OBJ_ID_DIVISION_ID & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()
                    Next
#End Region

                Case "TBBOM_USAGE"
#Region "TBBOM_USAGE"
                    Dim MAT_ID, BASE_QTY, BASE_UOM, COMPONENT, QTY, UOM, BLOCK, DAY_ID, SHIFT_NO, ALT, BOM_USG, CREATEDATE, MODDATE, MODBY As String
                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        With dt
                            MAT_ID = IIf(.Rows(i)("MAT_ID").ToString = "", "null", "'" & .Rows(i)("MAT_ID").ToString & "'")
                            BASE_QTY = IIf(.Rows(i)("BASE_QTY").ToString = "", "null", "'" & .Rows(i)("BASE_QTY").ToString & "'")
                            BASE_UOM = IIf(.Rows(i)("BASE_UOM").ToString = "", "null", "'" & .Rows(i)("BASE_UOM").ToString & "'")
                            COMPONENT = IIf(.Rows(i)("COMPONENT").ToString = "", "null", "'" & .Rows(i)("COMPONENT").ToString & "'")
                            QTY = IIf(.Rows(i)("QTY").ToString = "", "null", "'" & .Rows(i)("QTY").ToString & "'")
                            UOM = IIf(.Rows(i)("UOM").ToString = "", "null", "'" & .Rows(i)("UOM").ToString & "'")
                            BLOCK = IIf(.Rows(i)("BLOCK").ToString = "", "null", "'" & .Rows(i)("BLOCK").ToString & "'")
                            DAY_ID = IIf(.Rows(i)("DAY_ID").ToString = "", "null", "'" & .Rows(i)("DAY_ID").ToString & "'")
                            SHIFT_NO = IIf(.Rows(i)("SHIFT_NO").ToString = "", "null", "'" & .Rows(i)("SHIFT_NO").ToString & "'")
                            ALT = IIf(.Rows(i)("ALT").ToString = "", "null", "'" & .Rows(i)("ALT").ToString & "'")
                            BOM_USG = IIf(.Rows(i)("BOM_USG").ToString = "", "null", "'" & .Rows(i)("BOM_USG").ToString & "'")
                        End With

                        CREATEDATE = IIf(dt.Rows(i)("CREATEDATE").ToString = "", "null", "" & dt.Rows(i)("CREATEDATE").ToString & "")
                        If CREATEDATE <> "null" Then
                            CREATEDATE = ClsClobalFunction.ConvertDateTime(CREATEDATE)
                        End If

                        MODDATE = IIf(dt.Rows(i)("MODDATE").ToString = "", "null", "" & dt.Rows(i)("MODDATE").ToString & "")
                        If MODDATE <> "null" Then
                            MODDATE = ClsClobalFunction.ConvertDateTime(MODDATE)
                        End If
                        MODBY = IIf(dt.Rows(i)("MODBY").ToString = "", "null", "'" & dt.Rows(i)("MODBY").ToString & "'")

                        sql = "INSERT INTO [dbo].[TBBOM_USAGE]"
                        sql &= "([MAT_ID]"
                        sql &= "    ,[BASE_QTY]"
                        sql &= "    ,[BASE_UOM]"
                        sql &= "    ,[COMPONENT]"
                        sql &= "    ,[QTY]"
                        sql &= "    ,[UOM]"
                        sql &= "    ,[BLOCK]"
                        sql &= "    ,[DAY_ID]"
                        sql &= "    ,[SHIFT_NO]"
                        sql &= "    ,[ALT]"
                        sql &= "    ,[BOM_USG]"
                        sql &= "    ,[CREATEDATE]"
                        sql &= "    ,[MODDATE]"
                        sql &= "    ,[MODBY])"
                        sql &= "     VALUES"
                        sql &= "    (" & MAT_ID & ""
                        sql &= "    ," & BASE_QTY & ""
                        sql &= "    ," & BASE_UOM & ""
                        sql &= "    ," & COMPONENT & ""
                        sql &= "    ," & QTY & ""
                        sql &= "    ," & UOM & ""
                        sql &= "    ," & BLOCK & ""
                        sql &= "    ," & DAY_ID & ""
                        sql &= "    ," & SHIFT_NO & ""
                        sql &= "    ," & ALT & ""
                        sql &= "    ," & BOM_USG & ""
                        sql &= "    ," & CREATEDATE & ""
                        sql &= "    ," & MODDATE & ""
                        sql &= "    ," & MODBY & ")"


                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()

                    Next
#End Region

                Case "TBMATTERIAL_SITE"
#Region "TBMATTERIAL_SITE"
                    Dim MAT_ID, CREATEDATE, MODDATE, MODBY As String
                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        MAT_ID = IIf(dt.Rows(i)("MAT_ID").ToString = "", "null", dt.Rows(i)("MAT_ID").ToString)
                        CREATEDATE = IIf(dt.Rows(i)("CREATEDATE").ToString = "", "null", "" & dt.Rows(i)("CREATEDATE").ToString & "")
                        If CREATEDATE <> "null" Then
                            CREATEDATE = ClsClobalFunction.ConvertDateTime(CREATEDATE)
                        End If

                        MODDATE = IIf(dt.Rows(i)("MODDATE").ToString = "", "null", "" & dt.Rows(i)("MODDATE").ToString & "")
                        If MODDATE <> "null" Then
                            MODDATE = ClsClobalFunction.ConvertDateTime(MODDATE)
                        End If
                        MODBY = IIf(dt.Rows(i)("MODBY").ToString = "", "null", "'" & dt.Rows(i)("MODBY").ToString & "'")

                        sql = "INSERT INTO [dbo].[TBMATTERIAL_SITE]"
                        sql &= "([MAT_ID]"
                        sql &= "  ,[CREATEDATE]"
                        sql &= "  ,[MODDATE]"
                        sql &= "  ,[MODBY])"
                        sql &= "   VALUES"
                        sql &= "   (" & MAT_ID & ""
                        sql &= "   ," & CREATEDATE & ""
                        sql &= "   ," & MODDATE & ""
                        sql &= "   ," & MODBY & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()

                    Next
#End Region

                Case "TBCONVERSION"
#Region "TBCONVERSION"
                    Dim MAT_ID, BASE_UOM, ALTERNATIVE_UOM, NUMERATOR, DENOMINATOR, NORMAL_SIZE_L, BASE_UOM_BLOCK, DAY_ID, SHIFT_NO, CREATEDATE, MODDATE, MODBY As String
                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        With dt
                            MAT_ID = IIf(.Rows(i)("MAT_ID").ToString = "", "null", "'" & .Rows(i)("MAT_ID").ToString & "'")
                            BASE_UOM = IIf(.Rows(i)("BASE_UOM").ToString = "", "null", "'" & .Rows(i)("BASE_UOM").ToString & "'")
                            ALTERNATIVE_UOM = IIf(.Rows(i)("ALTERNATIVE_UOM").ToString = "", "null", "'" & .Rows(i)("ALTERNATIVE_UOM").ToString & "'")
                            NUMERATOR = IIf(.Rows(i)("NUMERATOR").ToString = "", "null", "" & .Rows(i)("NUMERATOR").ToString & "")
                            DENOMINATOR = IIf(.Rows(i)("DENOMINATOR").ToString = "", "null", "" & .Rows(i)("DENOMINATOR").ToString & "")
                            NORMAL_SIZE_L = IIf(.Rows(i)("NORMAL_SIZE_L").ToString = "", "null", "'" & .Rows(i)("NORMAL_SIZE_L").ToString & "'")
                            BASE_UOM_BLOCK = IIf(.Rows(i)("BASE_UOM_BLOCK").ToString = "", "null", "'" & .Rows(i)("BASE_UOM_BLOCK").ToString & "'")
                            DAY_ID = IIf(.Rows(i)("DAY_ID").ToString = "", "null", "" & .Rows(i)("DAY_ID").ToString & "")
                            SHIFT_NO = IIf(.Rows(i)("SHIFT_NO").ToString = "", "null", "" & .Rows(i)("SHIFT_NO").ToString & "")
                        End With
                        CREATEDATE = IIf(dt.Rows(i)("CREATEDATE").ToString = "", "null", "" & dt.Rows(i)("CREATEDATE").ToString & "")
                        If CREATEDATE <> "null" Then
                            CREATEDATE = ClsClobalFunction.ConvertDateTime(CREATEDATE)
                        End If

                        MODDATE = IIf(dt.Rows(i)("MODDATE").ToString = "", "null", "" & dt.Rows(i)("MODDATE").ToString & "")
                        If MODDATE <> "null" Then
                            MODDATE = ClsClobalFunction.ConvertDateTime(MODDATE)
                        End If
                        MODBY = IIf(dt.Rows(i)("MODBY").ToString = "", "null", "'" & dt.Rows(i)("MODBY").ToString & "'")

                        sql = "INSERT INTO [dbo].[TBCONVERSION]"
                        sql &= "([MAT_ID]"
                        sql &= "  ,[BASE_UOM]"
                        sql &= "  ,[ALTERNATIVE_UOM]"
                        sql &= "  ,[NUMERATOR]"
                        sql &= "  ,[DENOMINATOR]"
                        sql &= "  ,[NORMAL_SIZE_L]"
                        sql &= "  ,[BASE_UOM_BLOCK]"
                        sql &= "  ,[DAY_ID]"
                        sql &= "  ,[SHIFT_NO]"
                        sql &= "  ,[CREATEDATE]"
                        sql &= "  ,[MODDATE]"
                        sql &= "  ,[MODBY])"
                        sql &= "  VALUES"
                        sql &= "  (" & MAT_ID & ""
                        sql &= "  ," & BASE_UOM & ""
                        sql &= "  ," & ALTERNATIVE_UOM & ""
                        sql &= "  ," & NUMERATOR & ""
                        sql &= "  ," & DENOMINATOR & ""
                        sql &= "  ," & NORMAL_SIZE_L & ""
                        sql &= "  ," & BASE_UOM_BLOCK & ""
                        sql &= "  ," & DAY_ID & ""
                        sql &= "  ," & SHIFT_NO & ""
                        sql &= "  ," & CREATEDATE & ""
                        sql &= "  ," & MODDATE & ""
                        sql &= "  ," & MODBY & ")"


                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()

                    Next
#End Region

                Case "LKMAT_GROUP3"
#Region "LKMAT_GROUP3"
                    Dim MAT_GROUP3, GROUP3_NAME, CREATEDATE, MODDATE, MODBY As String
                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        MAT_GROUP3 = IIf(dt.Rows(i)("MAT_GROUP3").ToString = "", "null", "'" & dt.Rows(i)("MAT_GROUP3").ToString & "'")
                        GROUP3_NAME = IIf(dt.Rows(i)("GROUP3_NAME").ToString = "", "null", "'" & dt.Rows(i)("GROUP3_NAME").ToString & "'")
                        CREATEDATE = IIf(dt.Rows(i)("CREATEDATE").ToString = "", "null", "" & dt.Rows(i)("CREATEDATE").ToString & "")
                        If CREATEDATE <> "null" Then
                            CREATEDATE = ClsClobalFunction.ConvertDateTime(CREATEDATE)
                        End If

                        MODDATE = IIf(dt.Rows(i)("MODDATE").ToString = "", "null", "" & dt.Rows(i)("MODDATE").ToString & "")
                        If MODDATE <> "null" Then
                            MODDATE = ClsClobalFunction.ConvertDateTime(MODDATE)
                        End If
                        MODBY = IIf(dt.Rows(i)("MODBY").ToString = "", "null", "'" & dt.Rows(i)("MODBY").ToString & "'")

                        sql = "INSERT INTO [dbo].[LKMAT_GROUP3]"
                        sql &= "([MAT_GROUP3]"
                        sql &= "  ,[GROUP3_NAME]"
                        sql &= "  ,[CREATEDATE]"
                        sql &= "  ,[MODDATE]"
                        sql &= "  ,[MODBY])"
                        sql &= "   VALUES"
                        sql &= " (" & MAT_GROUP3 & ""
                        sql &= "  ," & GROUP3_NAME & ""
                        sql &= "  ," & CREATEDATE & ""
                        sql &= "  ," & MODDATE & ""
                        sql &= "  ," & MODBY & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()

                    Next
#End Region

                Case "LKDIVISION"
#Region "LKDIVISION"
                    Dim DIVISION_ID, DIVISION_NAME, CAN_RETURN, CREATEDATE, MODDATE, MODBY As String
                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1

                        DIVISION_ID = IIf(dt.Rows(i)("DIVISION_ID").ToString = "", "null", "'" & dt.Rows(i)("DIVISION_ID").ToString & "'")
                        DIVISION_NAME = IIf(dt.Rows(i)("DIVISION_NAME").ToString = "", "null", "'" & dt.Rows(i)("DIVISION_NAME").ToString & "'")
                        CAN_RETURN = IIf(dt.Rows(i)("CAN_RETURN").ToString = "", "null", "'" & dt.Rows(i)("CAN_RETURN").ToString & "'")
                        CREATEDATE = IIf(dt.Rows(i)("CREATEDATE").ToString = "", "null", "" & dt.Rows(i)("CREATEDATE").ToString & "")
                        If CREATEDATE <> "null" Then
                            CREATEDATE = ClsClobalFunction.ConvertDateTime(CREATEDATE)
                        End If

                        MODDATE = IIf(dt.Rows(i)("MODDATE").ToString = "", "null", "" & dt.Rows(i)("MODDATE").ToString & "")
                        If MODDATE <> "null" Then
                            MODDATE = ClsClobalFunction.ConvertDateTime(MODDATE)
                        End If
                        MODBY = IIf(dt.Rows(i)("MODBY").ToString = "", "null", "'" & dt.Rows(i)("MODBY").ToString & "'")

                        sql = "INSERT INTO [dbo].[LKDIVISION]"
                        sql &= "([DIVISION_ID]"
                        sql &= "  ,[DIVISION_NAME]"
                        sql &= "  ,[CAN_RETURN]"
                        sql &= "  ,[CREATEDATE]"
                        sql &= "  ,[MODDATE]"
                        sql &= "  ,[MODBY])"
                        sql &= "   VALUES"
                        sql &= "  (" & DIVISION_ID & ""
                        sql &= "  ," & DIVISION_NAME & ""
                        sql &= "  ," & CAN_RETURN & ""
                        sql &= "  ," & CREATEDATE & ""
                        sql &= "  ," & MODDATE & ""
                        sql &= "  ," & MODBY & ")"


                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()

                    Next
#End Region

                Case "TBMAT_RECOMMEND"
#Region "TBMAT_RECOMMEND"
                    Dim MAT_ID As String
                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        MAT_ID = IIf(dt.Rows(i)("MAT_ID").ToString = "", "null", "'" & dt.Rows(i)("MAT_ID").ToString & "'")

                        sql = "INSERT INTO [dbo].[TBMAT_RECOMMEND]"
                        sql &= "([MAT_ID])"
                        sql &= " VALUES"
                        sql &= "(" & MAT_ID & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()
                    Next
#End Region

                Case "APP_CONFIG"
#Region "APP_CONFIG"
                    'Dim CONFIG_KEY, CONFIG_DESC, CONFIG_TYPE, MODULE_TYPE, CONFIG_VALUE, ADMIN_ONLY As String
                    'Dim cmd As New SqlCommand
                    'With cmd
                    '    .CommandType = CommandType.Text
                    '    .Connection = trans.Connection
                    '    .Transaction = trans
                    'End With
                    'For i As Integer = 0 To dt.Rows.Count - 1
                    '    CONFIG_KEY = IIf(dt.Rows(i)("CONFIG_KEY").ToString = "", "null", "'" & dt.Rows(i)("CONFIG_KEY").ToString & "'")
                    '    CONFIG_DESC = IIf(dt.Rows(i)("CONFIG_DESC").ToString = "", "null", "'" & dt.Rows(i)("CONFIG_DESC").ToString & "'")
                    '    CONFIG_TYPE = IIf(dt.Rows(i)("CONFIG_TYPE").ToString = "", "null", "'" & dt.Rows(i)("CONFIG_TYPE").ToString & "'")
                    '    MODULE_TYPE = IIf(dt.Rows(i)("MODULE_TYPE").ToString = "", "null", "'" & dt.Rows(i)("MODULE_TYPE").ToString & "'")
                    '    CONFIG_VALUE = IIf(dt.Rows(i)("CONFIG_VALUE").ToString = "", "null", "'" & dt.Rows(i)("CONFIG_VALUE").ToString & "'")
                    '    ADMIN_ONLY = "0" ' IIf(dt.Rows(i)("ADMIN_ONLY").ToString = "", "null", "'" & dt.Rows(i)("ADMIN_ONLY").ToString & "'")

                    '    sql = "INSERT INTO [dbo].[APP_CONFIG]
                    '        ([CONFIG_KEY]
                    '        ,[CONFIG_DESC]
                    '        ,[CONFIG_TYPE]
                    '        ,[MODULE_TYPE]
                    '        ,[CONFIG_VALUE]
                    '        ,[CREATEDATE]
                    '        ,[MODDATE]
                    '        ,[MODBY]
                    '        ,[ADMIN_ONLY])
                    '    VALUES
                    '        (" & CONFIG_KEY & "
                    '        ," & CONFIG_DESC & "
                    '        ," & CONFIG_TYPE & "
                    '        ," & MODULE_TYPE & "
                    '        ," & CONFIG_VALUE & "
                    '        ,getdate()
                    '        ,getdate()
                    '        ,'" & modby & "'
                    '        ," & ADMIN_ONLY & ")"

                    '    cmd.CommandText = sql
                    '    cmd.ExecuteNonQuery()
                    'Next
#End Region

                Case "POS_CONFIG"
#Region "POS_CONFIG"
                    Dim POS_ID, MODULE_TYPE, POS_NO, POS_IP, TERMINAL_ID, RDCODE, ISMAIN, MAXMONEY, EDC1_PORTNAME, EDC2_PORTNAME As String
                    Dim EDC_SPEED, EDC_PARITY, EDC_TIMEOUT, CLEAR_SALE_INFO, SHOW_PUMP_INFO, CASHDRAWERPRINT, AUTOPRINT, CREATEDATE, MODDATE, MODBY As String
                    Dim EDC_MODEL, L_MID, L_TID As String

                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        POS_ID = IIf(dt.Rows(i)("POS_ID").ToString = "", "null", "'" & dt.Rows(i)("POS_ID").ToString & "'")
                        MODULE_TYPE = IIf(dt.Rows(i)("MODULE_TYPE").ToString = "", "null", "'" & dt.Rows(i)("MODULE_TYPE").ToString & "'")
                        POS_NO = IIf(dt.Rows(i)("POS_NO").ToString = "", "null", "'" & dt.Rows(i)("POS_NO").ToString & "'")
                        POS_IP = IIf(dt.Rows(i)("POS_IP").ToString = "", "null", "'" & dt.Rows(i)("POS_IP").ToString & "'")
                        TERMINAL_ID = IIf(dt.Rows(i)("TERMINAL_ID").ToString = "", "null", "'" & dt.Rows(i)("TERMINAL_ID").ToString & "'")
                        RDCODE = IIf(dt.Rows(i)("RDCODE").ToString = "", "null", "'" & dt.Rows(i)("RDCODE").ToString & "'")
                        ISMAIN = IIf(dt.Rows(i)("ISMAIN").ToString = "", "null", "'" & dt.Rows(i)("ISMAIN").ToString & "'")
                        MAXMONEY = IIf(dt.Rows(i)("MAXMONEY").ToString = "", "null", "'" & dt.Rows(i)("MAXMONEY").ToString & "'")
                        EDC1_PORTNAME = IIf(dt.Rows(i)("EDC1_PORTNAME").ToString = "", "null", "'" & dt.Rows(i)("EDC1_PORTNAME").ToString & "'")
                        EDC2_PORTNAME = IIf(dt.Rows(i)("EDC2_PORTNAME").ToString = "", "null", "'" & dt.Rows(i)("EDC2_PORTNAME").ToString & "'")
                        EDC_SPEED = IIf(dt.Rows(i)("EDC_SPEED").ToString = "", "null", "'" & dt.Rows(i)("EDC_SPEED").ToString & "'")
                        EDC_PARITY = IIf(dt.Rows(i)("EDC_PARITY").ToString = "", "null", "'" & dt.Rows(i)("EDC_PARITY").ToString & "'")
                        EDC_TIMEOUT = IIf(dt.Rows(i)("EDC_TIMEOUT").ToString = "", "null", "'" & dt.Rows(i)("EDC_TIMEOUT").ToString & "'")
                        CLEAR_SALE_INFO = IIf(dt.Rows(i)("CLEAR_SALE_INFO").ToString = "", "null", "'" & dt.Rows(i)("CLEAR_SALE_INFO").ToString & "'")
                        SHOW_PUMP_INFO = IIf(dt.Rows(i)("SHOW_PUMP_INFO").ToString = "", "null", "'" & dt.Rows(i)("SHOW_PUMP_INFO").ToString & "'")
                        CASHDRAWERPRINT = IIf(dt.Rows(i)("CASHDRAWERPRINT").ToString = "", "null", "'" & dt.Rows(i)("CASHDRAWERPRINT").ToString & "'")
                        AUTOPRINT = IIf(dt.Rows(i)("AUTOPRINT").ToString = "", "null", "'" & dt.Rows(i)("AUTOPRINT").ToString & "'")
                        EDC_MODEL = IIf(dt.Rows(i)("EDC_MODEL").ToString = "", "null", "'" & dt.Rows(i)("EDC_MODEL").ToString & "'")
                        L_MID = IIf(dt.Rows(i)("L_MID").ToString = "", "null", "'" & dt.Rows(i)("L_MID").ToString & "'")
                        L_TID = IIf(dt.Rows(i)("L_TID").ToString = "", "null", "'" & dt.Rows(i)("L_TID").ToString & "'")
                        CREATEDATE = IIf(dt.Rows(i)("CREATEDATE").ToString = "", "null", "" & dt.Rows(i)("CREATEDATE").ToString & "")
                        If CREATEDATE <> "null" Then
                            CREATEDATE = ClsClobalFunction.ConvertDateTime(CREATEDATE)
                        End If

                        MODDATE = IIf(dt.Rows(i)("MODDATE").ToString = "", "null", "" & dt.Rows(i)("MODDATE").ToString & "")
                        If MODDATE <> "null" Then
                            MODDATE = ClsClobalFunction.ConvertDateTime(MODDATE)
                        End If
                        MODBY = IIf(dt.Rows(i)("MODBY").ToString = "", "null", "'" & dt.Rows(i)("MODBY").ToString & "'")

                        sql = " INSERT INTO [dbo].[POS_CONFIG]"
                        sql &= " ([POS_ID]"
                        sql &= " ,[MODULE_TYPE]"
                        sql &= " ,[POS_NO]"
                        sql &= " ,[POS_IP]"
                        sql &= " ,[TERMINAL_ID]"
                        sql &= " ,[RDCODE]"
                        sql &= " ,[ISMAIN]"
                        sql &= " ,[MAXMONEY]"
                        sql &= " ,[EDC1_PORTNAME]"
                        sql &= " ,[EDC2_PORTNAME]"
                        sql &= " ,[EDC_SPEED]"
                        sql &= " ,[EDC_PARITY]"
                        sql &= " ,[EDC_TIMEOUT]"
                        sql &= " ,[CLEAR_SALE_INFO]"
                        sql &= " ,[SHOW_PUMP_INFO]"
                        sql &= " ,[CASHDRAWERPRINT]"
                        sql &= " ,[AUTOPRINT]"
                        sql &= " ,[CREATEDATE]"
                        sql &= " ,[MODDATE]"
                        sql &= " ,[MODBY]"
                        sql &= " ,[EDC_MODEL]"
                        sql &= " ,[L_MID]"
                        sql &= " ,[L_TID])"
                        sql &= " VALUES"
                        sql &= "(" & POS_ID & ""
                        sql &= " ," & MODULE_TYPE & ""
                        sql &= " ," & POS_NO & ""
                        sql &= " ," & POS_IP & ""
                        sql &= " ," & TERMINAL_ID & ""
                        sql &= " ," & RDCODE & ""
                        sql &= " ," & ISMAIN & ""
                        sql &= " ," & MAXMONEY & ""
                        sql &= " ," & EDC1_PORTNAME & ""
                        sql &= " ," & EDC2_PORTNAME & ""
                        sql &= " ," & EDC_SPEED & ""
                        sql &= " ," & EDC_PARITY & ""
                        sql &= " ," & EDC_TIMEOUT & ""
                        sql &= " ," & CLEAR_SALE_INFO & ""
                        sql &= " ," & SHOW_PUMP_INFO & ""
                        sql &= " ," & CASHDRAWERPRINT & ""
                        sql &= " ," & AUTOPRINT & ""
                        sql &= " ," & CREATEDATE & ""
                        sql &= " ," & MODDATE & ""
                        sql &= " ," & MODBY & ""
                        sql &= " ," & EDC_MODEL & ""
                        sql &= " ," & L_MID & ""
                        sql &= " ," & L_TID & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()
                    Next

#End Region

                Case "TBUSER"
#Region "TBUSER"
                    Dim USERNAME, PASSWORD, USERDESC, EXPIRE_DATE, ISUSER, POSITION_ID, ISAUTOCLEAR, USER_ID, F_NAME, L_NAME, CREATEDATE, MODDATE, MODBY As String

                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        USERNAME = IIf(dt.Rows(i)("USERNAME").ToString = "", "null", "'" & dt.Rows(i)("USERNAME").ToString & "'")
                        PASSWORD = IIf(dt.Rows(i)("PASSWORD").ToString = "", "null", "'" & ClsEncrypDecryp.base64Encode(dt.Rows(i)("PASSWORD").ToString, dt.Rows(i)("USERNAME").ToString) & "'")
                        USERDESC = IIf(dt.Rows(i)("USERDESC").ToString = "", "null", "'" & dt.Rows(i)("USERDESC").ToString & "'")
                        EXPIRE_DATE = IIf(dt.Rows(i)("EXPIRE_DATE").ToString = "", "null", "" & dt.Rows(i)("EXPIRE_DATE").ToString & "")
                        If EXPIRE_DATE <> "null" Then
                            EXPIRE_DATE = ClsClobalFunction.ConvertDate(EXPIRE_DATE)
                        End If

                        ISUSER = IIf(dt.Rows(i)("ISUSER").ToString = "", "null", "'" & dt.Rows(i)("ISUSER").ToString & "'")

                        Dim retPositionDesc As String = dt.Rows(i)("POSITION_DESC").ToString
                        If retPositionDesc = "" Then retPositionDesc = "Cashier"
                        Dim retPosition_ID As String = ClsClobalFunction.GET_ROLE_ID(retPositionDesc)
                        POSITION_ID = IIf(retPosition_ID = "", "null", "'" & retPosition_ID & "'")
                        ISAUTOCLEAR = "0"
                        USER_ID = "29"
                        F_NAME = "null"
                        L_NAME = "null"

                        CREATEDATE = IIf(dt.Rows(i)("CREATEDATE").ToString = "", "null", "" & dt.Rows(i)("CREATEDATE").ToString & "")
                        If CREATEDATE <> "null" Then
                            CREATEDATE = ClsClobalFunction.ConvertDateTime(CREATEDATE)
                        End If

                        MODDATE = IIf(dt.Rows(i)("MODDATE").ToString = "", "null", "" & dt.Rows(i)("MODDATE").ToString & "")
                        If MODDATE <> "null" Then
                            MODDATE = ClsClobalFunction.ConvertDateTime(MODDATE)
                        End If
                        MODBY = IIf(dt.Rows(i)("MODBY").ToString = "", "null", "'" & dt.Rows(i)("MODBY").ToString & "'")

                        sql = "INSERT INTO [dbo].[TBUSER]"
                        sql &= "([USERNAME]"
                        sql &= " ,[PASSWORD]"
                        sql &= " ,[USERDESC]"
                        sql &= " ,[EXPIRE_DATE]"
                        sql &= " ,[ISUSER]"
                        sql &= " ,[POSITION_ID]"
                        sql &= " ,[CREATEDATE]"
                        sql &= " ,[MODDATE]"
                        sql &= " ,[MODBY]"
                        sql &= " ,[ISAUTOCLEAR]"
                        sql &= " ,[F_NAME]"
                        sql &= " ,[L_NAME])"
                        sql &= " VALUES"
                        sql &= " (" & USERNAME & ""
                        sql &= " ," & PASSWORD & ""
                        sql &= " ," & USERDESC & ""
                        sql &= " ," & EXPIRE_DATE & ""
                        sql &= " ," & ISUSER & ""
                        sql &= " ," & POSITION_ID & ""
                        sql &= " ," & CREATEDATE & ""
                        sql &= " ," & MODDATE & ""
                        sql &= " ," & MODBY & ""
                        sql &= " ," & ISAUTOCLEAR & ""
                        sql &= " ," & F_NAME & ""
                        sql &= " ," & L_NAME & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()
                    Next
#End Region

                Case "TSJOURNAL"
#Region "TSJOURNAL"

                    sql = "If Not EXISTS(Select 1 FROM sys.columns "
                    sql &= " WHERE Name = N'IS_VOID_EDCWIFI'"
                    sql &= " And Object_ID = Object_ID(N'dbo.TSJOURNAL'))"
                    sql &= " BEGIN "
                    sql &= " ALTER TABLE TSJOURNAL "
                    sql &= " ADD IS_VOID_EDCWIFI CHAR(1) "
                    sql &= " End "
                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandText = sql
                        .CommandType = CommandType.Text
                        .Transaction = trans
                        .Connection = conn
                        .ExecuteNonQuery()
                    End With


                    sql = "If Not EXISTS(SELECT 1 FROM sys.columns "
                    sql &= " WHERE Name = N'IS_VOID_EDCWIFI'"
                    sql &= " And Object_ID = Object_ID(N'dbo.BKJOURNAL'))"
                    sql &= " BEGIN "
                    sql &= " ALTER TABLE BKJOURNAL "
                    sql &= " ADD IS_VOID_EDCWIFI CHAR(1) "
                    sql &= " End "

                    cmd.CommandText = sql
                    cmd.ExecuteNonQuery()

                    Dim JOURNAL_ID, POS_ID, USERNAME, TAX_NO, DAY_ID, SHIFT_ID, SALE_TYPE, CUS_ID, CUR_VAT As String
                    Dim TOTAL, VATTOTAL, DC, GRANDTOTAL, REFUND, VEHICLE_ID, CAR_TYPE, CARD_NO, CARD_TYPE As String
                    Dim INVOICE_NO, APPROVE_CODE, Signature, PRINT_TIMES, REF_JOURNAL_ID, PRICE_ID As String
                    Dim DEPT, DEPT1, COUNTER, EMP_NUMNER, DISTANCE, ACC_NUMBER, CARD_EXPIRE, TAX_CLASS As String
                    Dim DOCNO, LICENCENO, TAX_INVOICE, MCARDNO, LCARDNO, LCARDDATA, LREPOINT, LTRANS_NO As String
                    Dim LBATCH_NO, LCUSTOMER, LBALANCE, LPOINTTODAY, LREMARK, LPAY, LSTAND_ID As String
                    Dim LREDEEM_TRAN_ID, FLEET_HOST_ID, FLEET_CUST_TAX_ID, FLEET_CUST_BRANCH_NBR As String
                    Dim FLEET_CUST_NAME, FLEET_CUST_ADDRESS, FLEET_CAR_PLATE, IS_VOID_EDC As String
                    Dim AVAILABLE_CREDIT, FG_REF_NO, FLEET_DOCNO, FLEET_CUS_ID, REASON_ID, REASON_DESC As String
                    Dim ORIGINAL_TAX, DOC_TYPE, STATUS_FLAG, TOTAL_BALANCE, ORIGINAL_TAX_ID, FULLTAX_BOOKING_NUM As String
                    Dim DC_BILL_VALUE, DC_BILL_AMOUNT, DC_BILL_TYPE, TRANSACTION_ID, LRESCODE, MethodParam As String
                    Dim RoundSF, EarnByBiz, BILL_PROMOTION_PRINT_TIMES As String
                    Dim IS_VOID_EDCWIFI, CREATEDATE, MODDATE, MODBY As String


                    For i As Integer = 0 To dt.Rows.Count - 1
                        JOURNAL_ID = IIf(dt.Rows(i)("JOURNAL_ID").ToString = "", "null", "'" & dt.Rows(i)("JOURNAL_ID").ToString & "'")
                        POS_ID = IIf(dt.Rows(i)("POS_ID").ToString = "", "null", "'" & dt.Rows(i)("POS_ID").ToString & "'")
                        USERNAME = IIf(dt.Rows(i)("USERNAME").ToString = "", "null", "'" & dt.Rows(i)("USERNAME").ToString & "'")
                        TAX_NO = IIf(dt.Rows(i)("TAX_NO").ToString = "", "null", "'" & dt.Rows(i)("TAX_NO").ToString & "'")
                        DAY_ID = IIf(dt.Rows(i)("DAY_ID").ToString = "", "null", "'" & dt.Rows(i)("DAY_ID").ToString & "'")
                        SHIFT_ID = IIf(dt.Rows(i)("SHIFT_ID").ToString = "", "null", "'" & dt.Rows(i)("SHIFT_ID").ToString & "'")
                        SALE_TYPE = IIf(dt.Rows(i)("SALE_TYPE").ToString = "", "null", "'" & dt.Rows(i)("SALE_TYPE").ToString & "'")
                        CUS_ID = IIf(dt.Rows(i)("CUS_ID").ToString = "", "null", "'" & dt.Rows(i)("CUS_ID").ToString & "'")
                        CUR_VAT = IIf(dt.Rows(i)("CUR_VAT").ToString = "", "null", "'" & dt.Rows(i)("CUR_VAT").ToString & "'")
                        TOTAL = IIf(dt.Rows(i)("TOTAL").ToString = "", "null", "'" & dt.Rows(i)("TOTAL").ToString & "'")

                        VATTOTAL = IIf(dt.Rows(i)("VATTOTAL").ToString = "", "null", "'" & dt.Rows(i)("VATTOTAL").ToString & "'")
                        DC = IIf(dt.Rows(i)("DC").ToString = "", "null", "'" & dt.Rows(i)("DC").ToString & "'")
                        GRANDTOTAL = IIf(dt.Rows(i)("GRANDTOTAL").ToString = "", "null", "'" & dt.Rows(i)("GRANDTOTAL").ToString & "'")
                        REFUND = IIf(dt.Rows(i)("REFUND").ToString = "", "null", "'" & dt.Rows(i)("REFUND").ToString & "'")
                        VEHICLE_ID = IIf(dt.Rows(i)("VEHICLE_ID").ToString = "", "null", "'" & dt.Rows(i)("VEHICLE_ID").ToString & "'")
                        CAR_TYPE = IIf(dt.Rows(i)("CAR_TYPE").ToString = "", "null", "'" & dt.Rows(i)("CAR_TYPE").ToString & "'")
                        CARD_NO = IIf(dt.Rows(i)("CARD_NO").ToString = "", "null", "'" & dt.Rows(i)("CARD_NO").ToString & "'")
                        CARD_TYPE = IIf(dt.Rows(i)("CARD_TYPE").ToString = "", "null", "'" & dt.Rows(i)("CARD_TYPE").ToString & "'")
                        INVOICE_NO = IIf(dt.Rows(i)("INVOICE_NO").ToString = "", "null", "'" & dt.Rows(i)("INVOICE_NO").ToString & "'")
                        APPROVE_CODE = IIf(dt.Rows(i)("APPROVE_CODE").ToString = "", "null", "'" & dt.Rows(i)("APPROVE_CODE").ToString & "'")

                        Signature = IIf(dt.Rows(i)("Signature").ToString = "", "null", "'" & dt.Rows(i)("Signature").ToString & "'")
                        PRINT_TIMES = IIf(dt.Rows(i)("PRINT_TIMES").ToString = "", "null", "'" & dt.Rows(i)("PRINT_TIMES").ToString & "'")
                        REF_JOURNAL_ID = IIf(dt.Rows(i)("REF_JOURNAL_ID").ToString = "", "null", "'" & dt.Rows(i)("REF_JOURNAL_ID").ToString & "'")
                        PRICE_ID = IIf(dt.Rows(i)("PRICE_ID").ToString = "", "null", "'" & dt.Rows(i)("PRICE_ID").ToString & "'")
                        DEPT = IIf(dt.Rows(i)("DEPT").ToString = "", "null", "'" & dt.Rows(i)("DEPT").ToString & "'")
                        DEPT1 = IIf(dt.Rows(i)("DEPT1").ToString = "", "null", "'" & dt.Rows(i)("DEPT1").ToString & "'")
                        COUNTER = IIf(dt.Rows(i)("COUNTER").ToString = "", "null", "'" & dt.Rows(i)("COUNTER").ToString & "'")
                        EMP_NUMNER = IIf(dt.Rows(i)("EMP_NUMNER").ToString = "", "null", "'" & dt.Rows(i)("EMP_NUMNER").ToString & "'")
                        DISTANCE = IIf(dt.Rows(i)("DISTANCE").ToString = "", "null", "'" & dt.Rows(i)("DISTANCE").ToString & "'")
                        ACC_NUMBER = IIf(dt.Rows(i)("ACC_NUMBER").ToString = "", "null", "'" & dt.Rows(i)("ACC_NUMBER").ToString & "'")

                        CARD_EXPIRE = IIf(dt.Rows(i)("CARD_EXPIRE").ToString = "", "null", "'" & dt.Rows(i)("CARD_EXPIRE").ToString & "'")
                        TAX_CLASS = IIf(dt.Rows(i)("TAX_CLASS").ToString = "", "null", "'" & dt.Rows(i)("TAX_CLASS").ToString & "'")
                        DOCNO = IIf(dt.Rows(i)("DOCNO").ToString = "", "null", "'" & dt.Rows(i)("DOCNO").ToString & "'")
                        LICENCENO = IIf(dt.Rows(i)("LICENCENO").ToString = "", "null", "'" & dt.Rows(i)("LICENCENO").ToString & "'")
                        TAX_INVOICE = IIf(dt.Rows(i)("TAX_INVOICE").ToString = "", "null", "'" & dt.Rows(i)("TAX_INVOICE").ToString & "'")
                        MCARDNO = IIf(dt.Rows(i)("MCARDNO").ToString = "", "null", "'" & dt.Rows(i)("MCARDNO").ToString & "'")
                        LCARDNO = IIf(dt.Rows(i)("LCARDNO").ToString = "", "null", "'" & dt.Rows(i)("LCARDNO").ToString & "'")
                        LCARDDATA = IIf(dt.Rows(i)("LCARDDATA").ToString = "", "null", "'" & dt.Rows(i)("LCARDDATA").ToString & "'")
                        LREPOINT = IIf(dt.Rows(i)("LREPOINT").ToString = "", "null", "'" & dt.Rows(i)("LREPOINT").ToString & "'")
                        LTRANS_NO = IIf(dt.Rows(i)("LTRANS_NO").ToString = "", "null", "'" & dt.Rows(i)("LTRANS_NO").ToString & "'")

                        LBATCH_NO = IIf(dt.Rows(i)("LBATCH_NO").ToString = "", "null", "'" & dt.Rows(i)("LBATCH_NO").ToString & "'")
                        LCUSTOMER = IIf(dt.Rows(i)("LCUSTOMER").ToString = "", "null", "'" & dt.Rows(i)("LCUSTOMER").ToString & "'")
                        LBALANCE = IIf(dt.Rows(i)("LBALANCE").ToString = "", "null", "'" & dt.Rows(i)("LBALANCE").ToString & "'")
                        LPOINTTODAY = IIf(dt.Rows(i)("LPOINTTODAY").ToString = "", "null", "'" & dt.Rows(i)("LPOINTTODAY").ToString & "'")
                        LREMARK = IIf(dt.Rows(i)("LREMARK").ToString = "", "null", "'" & dt.Rows(i)("LREMARK").ToString & "'")
                        LPAY = IIf(dt.Rows(i)("LPAY").ToString = "", "null", "'" & dt.Rows(i)("LPAY").ToString & "'")
                        LSTAND_ID = IIf(dt.Rows(i)("LSTAND_ID").ToString = "", "null", "'" & dt.Rows(i)("LSTAND_ID").ToString & "'")
                        LREDEEM_TRAN_ID = IIf(dt.Rows(i)("LREDEEM_TRAN_ID").ToString = "", "null", "'" & dt.Rows(i)("LREDEEM_TRAN_ID").ToString & "'")
                        FLEET_HOST_ID = IIf(dt.Rows(i)("FLEET_HOST_ID").ToString = "", "null", "'" & dt.Rows(i)("FLEET_HOST_ID").ToString & "'")
                        FLEET_CUST_TAX_ID = IIf(dt.Rows(i)("FLEET_CUST_TAX_ID").ToString = "", "null", "'" & dt.Rows(i)("FLEET_CUST_TAX_ID").ToString & "'")


                        FLEET_CUST_BRANCH_NBR = IIf(dt.Rows(i)("FLEET_CUST_BRANCH_NBR").ToString = "", "null", "'" & dt.Rows(i)("FLEET_CUST_BRANCH_NBR").ToString & "'")
                        FLEET_CUST_NAME = IIf(dt.Rows(i)("FLEET_CUST_NAME").ToString = "", "null", "'" & dt.Rows(i)("FLEET_CUST_NAME").ToString & "'")
                        FLEET_CUST_ADDRESS = IIf(dt.Rows(i)("FLEET_CUST_ADDRESS").ToString = "", "null", "'" & dt.Rows(i)("FLEET_CUST_ADDRESS").ToString & "'")
                        FLEET_CAR_PLATE = IIf(dt.Rows(i)("FLEET_CAR_PLATE").ToString = "", "null", "'" & dt.Rows(i)("FLEET_CAR_PLATE").ToString & "'")
                        IS_VOID_EDC = IIf(dt.Rows(i)("IS_VOID_EDC").ToString = "", "null", "'" & dt.Rows(i)("IS_VOID_EDC").ToString & "'")
                        AVAILABLE_CREDIT = IIf(dt.Rows(i)("AVAILABLE_CREDIT").ToString = "", "null", "'" & dt.Rows(i)("AVAILABLE_CREDIT").ToString & "'")
                        FG_REF_NO = IIf(dt.Rows(i)("FG_REF_NO").ToString = "", "null", "'" & dt.Rows(i)("FG_REF_NO").ToString & "'")
                        FLEET_DOCNO = IIf(dt.Rows(i)("FLEET_DOCNO").ToString = "", "null", "'" & dt.Rows(i)("FLEET_DOCNO").ToString & "'")
                        FLEET_CUS_ID = IIf(dt.Rows(i)("FLEET_CUS_ID").ToString = "", "null", "'" & dt.Rows(i)("FLEET_CUS_ID").ToString & "'")
                        REASON_ID = IIf(dt.Rows(i)("REASON_ID").ToString = "", "null", "'" & dt.Rows(i)("REASON_ID").ToString & "'")

                        REASON_DESC = IIf(dt.Rows(i)("REASON_DESC").ToString = "", "null", "'" & dt.Rows(i)("REASON_DESC").ToString & "'")
                        ORIGINAL_TAX = IIf(dt.Rows(i)("ORIGINAL_TAX").ToString = "", "null", "'" & dt.Rows(i)("ORIGINAL_TAX").ToString & "'")
                        DOC_TYPE = IIf(dt.Rows(i)("DOC_TYPE").ToString = "", "null", "'" & dt.Rows(i)("DOC_TYPE").ToString & "'")
                        STATUS_FLAG = IIf(dt.Rows(i)("STATUS_FLAG").ToString = "", "null", "'" & dt.Rows(i)("STATUS_FLAG").ToString & "'")
                        TOTAL_BALANCE = IIf(dt.Rows(i)("TOTAL_BALANCE").ToString = "", "null", "'" & dt.Rows(i)("TOTAL_BALANCE").ToString & "'")
                        ORIGINAL_TAX_ID = IIf(dt.Rows(i)("ORIGINAL_TAX_ID").ToString = "", "null", "'" & dt.Rows(i)("ORIGINAL_TAX_ID").ToString & "'")
                        FULLTAX_BOOKING_NUM = IIf(dt.Rows(i)("FULLTAX_BOOKING_NUM").ToString = "", "null", "'" & dt.Rows(i)("FULLTAX_BOOKING_NUM").ToString & "'")
                        DC_BILL_VALUE = "null" ' IIf(dt.Rows(i)("DC_BILL_VALUE").ToString = "", "null", "'" & dt.Rows(i)("DC_BILL_VALUE").ToString & "'")
                        DC_BILL_AMOUNT = "null" 'IIf(dt.Rows(i)("DC_BILL_AMOUNT").ToString = "", "null", "'" & dt.Rows(i)("DC_BILL_AMOUNT").ToString & "'")
                        DC_BILL_TYPE = "null" 'IIf(dt.Rows(i)("DC_BILL_TYPE").ToString = "", "null", "'" & dt.Rows(i)("DC_BILL_TYPE").ToString & "'")

                        TRANSACTION_ID = "null" 'IIf(dt.Rows(i)("TRANSACTION_ID").ToString = "", "null", "'" & dt.Rows(i)("TRANSACTION_ID").ToString & "'")
                        LRESCODE = IIf(dt.Rows(i)("LRESCODE").ToString = "", "null", "'" & dt.Rows(i)("LRESCODE").ToString & "'")
                        MethodParam = IIf(dt.Rows(i)("MethodParam").ToString = "", "null", "'" & dt.Rows(i)("MethodParam").ToString & "'")
                        RoundSF = IIf(dt.Rows(i)("RoundSF").ToString = "", "null", "'" & dt.Rows(i)("RoundSF").ToString & "'")
                        EarnByBiz = IIf(dt.Rows(i)("EarnByBiz").ToString = "", "null", "'" & dt.Rows(i)("EarnByBiz").ToString & "'")
                        'BILL_PROMOTION_ID = "null" 'IIf(dt.Rows(i)("BILL_PROMOTION_ID").ToString = "", "null", "'" & dt.Rows(i)("BILL_PROMOTION_ID").ToString & "'")
                        BILL_PROMOTION_PRINT_TIMES = "null" 'IIf(dt.Rows(i)("BILL_PROMOTION_PRINT_TIMES").ToString = "", "null", "'" & dt.Rows(i)("BILL_PROMOTION_PRINT_TIMES").ToString & "'")

                        Dim columns As DataColumnCollection = dt.Columns
                        If columns.Contains("IS_VOID_EDCWIFI") Then
                            IS_VOID_EDCWIFI = IIf(dt.Rows(i)("IS_VOID_EDCWIFI").ToString = "", "null", "'" & dt.Rows(i)("IS_VOID_EDCWIFI").ToString & "'")
                        Else
                            IS_VOID_EDCWIFI = "null"
                        End If

                        CREATEDATE = IIf(dt.Rows(i)("CREATEDATE").ToString = "", "null", "" & dt.Rows(i)("CREATEDATE").ToString & "")
                        If CREATEDATE <> "null" Then
                            CREATEDATE = ClsClobalFunction.ConvertDateTime(CREATEDATE)
                        End If

                        MODDATE = IIf(dt.Rows(i)("MODDATE").ToString = "", "null", "" & dt.Rows(i)("MODDATE").ToString & "")
                        If MODDATE <> "null" Then
                            MODDATE = ClsClobalFunction.ConvertDateTime(MODDATE)
                        End If
                        MODBY = IIf(dt.Rows(i)("MODBY").ToString = "", "null", "'" & dt.Rows(i)("MODBY").ToString & "'")


                        sql = "INSERT INTO [dbo].[TSJOURNAL]"
                        sql &= "([JOURNAL_ID]"
                        sql &= " ,[POS_ID]"
                        sql &= " ,[USERNAME]"
                        sql &= " ,[TAX_NO]"
                        sql &= " ,[DAY_ID]"
                        sql &= " ,[SHIFT_ID]"
                        sql &= " ,[SALE_TYPE]"
                        sql &= " ,[CUS_ID]"
                        sql &= " ,[CUR_VAT]"
                        sql &= " ,[TOTAL]"
                        sql &= " ,[VATTOTAL]"
                        sql &= " ,[DC]"
                        sql &= " ,[GRANDTOTAL]"
                        sql &= " ,[REFUND]"
                        sql &= " ,[VEHICLE_ID]"
                        sql &= " ,[CAR_TYPE]"
                        sql &= " ,[CARD_NO]"
                        sql &= " ,[CARD_TYPE]"
                        sql &= " ,[INVOICE_NO]"
                        sql &= " ,[APPROVE_CODE]"
                        sql &= " ,[SIGNATURE]"
                        sql &= " ,[PRINT_TIMES]"
                        sql &= " ,[REF_JOURNAL_ID]"
                        sql &= " ,[PRICE_ID]"
                        sql &= " ,[CREATEDATE]"
                        sql &= " ,[MODDATE]"
                        sql &= " ,[MODBY]"
                        sql &= " ,[DEPT]"
                        sql &= " ,[DEPT1]"
                        sql &= " ,[COUNTER]"
                        sql &= " ,[EMP_NUMNER]"
                        sql &= " ,[DISTANCE]"
                        sql &= " ,[ACC_NUMBER]"
                        sql &= " ,[CARD_EXPIRE]"
                        sql &= " ,[TAX_CLASS]"
                        sql &= " ,[DOCNO]"
                        sql &= " ,[LICENCENO]"
                        sql &= " ,[TAX_INVOICE]"
                        sql &= " ,[MCARDNO]"
                        sql &= " ,[LCARDNO]"
                        sql &= " ,[LCARDDATA]"
                        sql &= " ,[LREPOINT]"
                        sql &= " ,[LTRANS_NO]"
                        sql &= " ,[LBATCH_NO]"
                        sql &= " ,[LCUSTOMER]"
                        sql &= " ,[LBALANCE]"
                        sql &= " ,[LPOINTTODAY]"
                        sql &= " ,[LREMARK]"
                        sql &= " ,[LPAY]"
                        sql &= " ,[LSTAND_ID]"
                        sql &= " ,[LREDEEM_TRAN_ID]"
                        sql &= " ,[FLEET_HOST_ID]"
                        sql &= " ,[FLEET_CUST_TAX_ID]"
                        sql &= " ,[FLEET_CUST_BRANCH_NBR]"
                        sql &= " ,[FLEET_CUST_NAME]"
                        sql &= " ,[FLEET_CUST_ADDRESS]"
                        sql &= " ,[FLEET_CAR_PLATE]"
                        sql &= " ,[IS_VOID_EDC]"
                        sql &= " ,[AVAILABLE_CREDIT]"
                        sql &= " ,[FG_REF_NO]"
                        sql &= " ,[FLEET_DOCNO]"
                        sql &= " ,[FLEET_CUS_ID]"
                        sql &= " ,[REASON_ID]"
                        sql &= " ,[REASON_DESC]"
                        sql &= " ,[ORIGINAL_TAX]"
                        sql &= " ,[DOC_TYPE]"
                        sql &= " ,[STATUS_FLAG]"
                        sql &= " ,[TOTAL_BALANCE]"
                        sql &= " ,[ORIGINAL_TAX_ID]"
                        sql &= " ,[FULLTAX_BOOKING_NUM]"
                        sql &= " ,[DC_BILL_VALUE]"
                        sql &= " ,[DC_BILL_AMOUNT]"
                        sql &= " ,[DC_BILL_TYPE]"
                        sql &= " ,[TRANSACTION_ID]"
                        sql &= " ,[LRESCODE]"
                        sql &= " ,[MethodParam]"
                        sql &= " ,[RoundSF]"
                        sql &= " ,[EarnByBiz]"
                        sql &= " ,[BILL_PROMOTION_PRINT_TIMES],[IS_VOID_EDCWIFI])"
                        sql &= " VALUES"
                        sql &= "  (" & JOURNAL_ID & ""
                        sql &= " ," & POS_ID & ""
                        sql &= " ," & USERNAME & ""
                        sql &= " ," & TAX_NO & ""
                        sql &= " ," & DAY_ID & ""
                        sql &= " ," & SHIFT_ID & ""
                        sql &= " ," & SALE_TYPE & ""
                        sql &= " ," & CUS_ID & ""
                        sql &= " ," & CUR_VAT & ""
                        sql &= " ," & TOTAL & ""
                        sql &= " ," & VATTOTAL & ""
                        sql &= " ," & DC & ""
                        sql &= " ," & GRANDTOTAL & ""
                        sql &= " ," & REFUND & ""
                        sql &= "  ," & VEHICLE_ID & ""
                        sql &= " ," & CAR_TYPE & ""
                        sql &= "  ," & CARD_NO & ""
                        sql &= " ," & CARD_TYPE & ""
                        sql &= " ," & INVOICE_NO & ""
                        sql &= " ," & APPROVE_CODE & ""
                        sql &= " ," & Signature & ""
                        sql &= " ," & PRINT_TIMES & ""
                        sql &= " ," & REF_JOURNAL_ID & ""
                        sql &= " ," & PRICE_ID & ""
                        sql &= " ," & CREATEDATE & ""
                        sql &= " ," & MODDATE & ""
                        sql &= " ," & MODBY & ""
                        sql &= " ," & DEPT & ""
                        sql &= " ," & DEPT1 & ""
                        sql &= " ," & COUNTER & ""
                        sql &= " ," & EMP_NUMNER & ""
                        sql &= " ," & DISTANCE & ""
                        sql &= " ," & ACC_NUMBER & ""
                        sql &= " ," & CARD_EXPIRE & ""
                        sql &= " ," & TAX_CLASS & ""
                        sql &= " ," & DOCNO & ""
                        sql &= " ," & LICENCENO & ""
                        sql &= " ," & TAX_INVOICE & ""
                        sql &= " ," & MCARDNO & ""
                        sql &= " ," & LCARDNO & ""
                        sql &= " ," & LCARDDATA & ""
                        sql &= " ," & LREPOINT & ""
                        sql &= " ," & LTRANS_NO & ""
                        sql &= " ," & LBATCH_NO & ""
                        sql &= " ," & LCUSTOMER & ""
                        sql &= " ," & LBALANCE & ""
                        sql &= " ," & LPOINTTODAY & ""
                        sql &= " ," & LREMARK & ""
                        sql &= " ," & LPAY & ""
                        sql &= " ," & LSTAND_ID & ""
                        sql &= " ," & LREDEEM_TRAN_ID & ""
                        sql &= " ," & FLEET_HOST_ID & ""
                        sql &= " ," & FLEET_CUST_TAX_ID & ""
                        sql &= " ," & FLEET_CUST_BRANCH_NBR & ""
                        sql &= " ," & FLEET_CUST_NAME & ""
                        sql &= " ," & FLEET_CUST_ADDRESS & ""
                        sql &= " ," & FLEET_CAR_PLATE & ""
                        sql &= " ," & IS_VOID_EDC & ""
                        sql &= " ," & AVAILABLE_CREDIT & ""
                        sql &= " ," & FG_REF_NO & ""
                        sql &= " ," & FLEET_DOCNO & ""
                        sql &= " ," & FLEET_CUS_ID & ""
                        sql &= " ," & REASON_ID & ""
                        sql &= " ," & REASON_DESC & ""
                        sql &= " ," & ORIGINAL_TAX & ""
                        sql &= " ," & DOC_TYPE & ""
                        sql &= " ," & STATUS_FLAG & ""
                        sql &= " ," & TOTAL_BALANCE & ""
                        sql &= " ," & ORIGINAL_TAX_ID & ""
                        sql &= " ," & FULLTAX_BOOKING_NUM & ""
                        sql &= " ," & DC_BILL_VALUE & ""
                        sql &= " ," & DC_BILL_AMOUNT & ""
                        sql &= " ," & DC_BILL_TYPE & ""
                        sql &= " ," & TRANSACTION_ID & ""
                        sql &= " ," & LRESCODE & ""
                        sql &= " ," & MethodParam & ""
                        sql &= " ," & RoundSF & ""
                        sql &= " ," & EarnByBiz & ""
                        sql &= " ," & BILL_PROMOTION_PRINT_TIMES & "," & IS_VOID_EDCWIFI & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()
                    Next
#End Region

                Case "TSJOURNAL_DETAIL"
#Region "TSJOURNAL_DETAIL"
                    Dim JOURNAL_ID, ITEM_NO, MAT_ID, VOLUME, QTY, PRICE, VALUE, DC_PRICE, DC_VALUE, TRANS_NO,
                    HOSE_ID, PUMP_ID, TANK_ID, IS_OFFLINE, JOURNAL_REF, DC_ITEM_VALUE, DC_ITEM_AMOUNT,
                    DC_ITEM_TYPE, DC_BILL_VALUE, DC_BILL_AMOUNT, DC_BILL_TYPE, VAT_TYPE, VATable,
                    VAT, Total, NetBeforeVat, NetTotal As String

                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        JOURNAL_ID = IIf(dt.Rows(i)("JOURNAL_ID").ToString = "", "null", "'" & dt.Rows(i)("JOURNAL_ID").ToString & "'")
                        ITEM_NO = IIf(dt.Rows(i)("ITEM_NO").ToString = "", "null", "'" & dt.Rows(i)("ITEM_NO").ToString & "'")
                        MAT_ID = IIf(dt.Rows(i)("MAT_ID").ToString = "", "null", "'" & dt.Rows(i)("MAT_ID").ToString & "'")
                        VOLUME = IIf(dt.Rows(i)("VOLUME").ToString = "", "null", "'" & dt.Rows(i)("VOLUME").ToString & "'")
                        QTY = IIf(dt.Rows(i)("QTY").ToString = "", "null", "'" & dt.Rows(i)("QTY").ToString & "'")
                        PRICE = IIf(dt.Rows(i)("PRICE").ToString = "", "null", "'" & dt.Rows(i)("PRICE").ToString & "'")
                        VALUE = IIf(dt.Rows(i)("VALUE").ToString = "", "null", "'" & dt.Rows(i)("VALUE").ToString & "'")
                        DC_PRICE = IIf(dt.Rows(i)("DC_PRICE").ToString = "", "null", "'" & dt.Rows(i)("DC_PRICE").ToString & "'")
                        DC_VALUE = IIf(dt.Rows(i)("DC_VALUE").ToString = "", "null", "'" & dt.Rows(i)("DC_VALUE").ToString & "'")
                        TRANS_NO = IIf(dt.Rows(i)("TRANS_NO").ToString = "", "null", "'" & dt.Rows(i)("TRANS_NO").ToString & "'")

                        HOSE_ID = IIf(dt.Rows(i)("HOSE_ID").ToString = "", "null", "'" & dt.Rows(i)("HOSE_ID").ToString & "'")
                        PUMP_ID = IIf(dt.Rows(i)("PUMP_ID").ToString = "", "null", "'" & dt.Rows(i)("PUMP_ID").ToString & "'")
                        TANK_ID = IIf(dt.Rows(i)("TANK_ID").ToString = "", "null", "'" & dt.Rows(i)("TANK_ID").ToString & "'")
                        IS_OFFLINE = IIf(dt.Rows(i)("IS_OFFLINE").ToString = "", "null", "'" & dt.Rows(i)("IS_OFFLINE").ToString & "'")
                        JOURNAL_REF = "null" 'IIf(dt.Rows(i)("JOURNAL_REF").ToString = "", "null", "'" & dt.Rows(i)("JOURNAL_REF").ToString & "'")
                        DC_ITEM_VALUE = "null" 'IIf(dt.Rows(i)("DC_ITEM_VALUE").ToString = "", "null", "'" & dt.Rows(i)("DC_ITEM_VALUE").ToString & "'")
                        DC_ITEM_AMOUNT = "null" 'IIf(dt.Rows(i)("DC_ITEM_AMOUNT").ToString = "", "null", "'" & dt.Rows(i)("DC_ITEM_AMOUNT").ToString & "'")
                        DC_ITEM_TYPE = "null" 'IIf(dt.Rows(i)("DC_ITEM_TYPE").ToString = "", "null", "'" & dt.Rows(i)("DC_ITEM_TYPE").ToString & "'")
                        DC_BILL_VALUE = "null" 'IIf(dt.Rows(i)("DC_BILL_VALUE").ToString = "", "null", "'" & dt.Rows(i)("DC_BILL_VALUE").ToString & "'")
                        DC_BILL_AMOUNT = "null" 'IIf(dt.Rows(i)("DC_BILL_AMOUNT").ToString = "", "null", "'" & dt.Rows(i)("DC_BILL_AMOUNT").ToString & "'")

                        DC_BILL_TYPE = "null" 'IIf(dt.Rows(i)("DC_BILL_TYPE").ToString = "", "null", "'" & dt.Rows(i)("DC_BILL_TYPE").ToString & "'")
                        VAT_TYPE = "null" 'IIf(dt.Rows(i)("VAT_TYPE").ToString = "", "null", "'" & dt.Rows(i)("VAT_TYPE").ToString & "'")
                        VATable = "null" 'IIf(dt.Rows(i)("VATable").ToString = "", "null", "'" & dt.Rows(i)("VATable").ToString & "'")
                        VAT = "null" 'IIf(dt.Rows(i)("VAT").ToString = "", "null", "'" & dt.Rows(i)("VAT").ToString & "'")
                        Total = "null" 'IIf(dt.Rows(i)("Total").ToString = "", "null", "'" & dt.Rows(i)("Total").ToString & "'")
                        NetBeforeVat = "null" ' IIf(dt.Rows(i)("NetBeforeVat").ToString = "", "null", "'" & dt.Rows(i)("NetBeforeVat").ToString & "'")
                        NetTotal = "null" ' IIf(dt.Rows(i)("NetTotal").ToString = "", "null", "'" & dt.Rows(i)("NetTotal").ToString & "'")

                        sql = "INSERT INTO [dbo].[TSJOURNAL_DETAIL]"
                        sql &= "([JOURNAL_ID]"
                        sql &= " ,[ITEM_NO]"
                        sql &= " ,[MAT_ID]"
                        sql &= " ,[VOLUME]"
                        sql &= " ,[QTY]"
                        sql &= " ,[PRICE]"
                        sql &= " ,[VALUE]"
                        sql &= " ,[DC_PRICE]"
                        sql &= " ,[DC_VALUE]"
                        sql &= " ,[TRANS_NO]"
                        sql &= " ,[HOSE_ID]"
                        sql &= " ,[PUMP_ID]"
                        sql &= " ,[TANK_ID]"
                        sql &= " ,[IS_OFFLINE]"
                        sql &= " ,[JOURNAL_REF]"
                        sql &= " ,[DC_ITEM_VALUE]"
                        sql &= " ,[DC_ITEM_AMOUNT]"
                        sql &= " ,[DC_ITEM_TYPE]"
                        sql &= " ,[DC_BILL_VALUE]"
                        sql &= " ,[DC_BILL_AMOUNT]"
                        sql &= " ,[DC_BILL_TYPE]"
                        sql &= " ,[VAT_TYPE]"
                        sql &= " ,[VATable]"
                        sql &= " ,[VAT]"
                        sql &= " ,[Total]"
                        sql &= " ,[NetBeforeVat]"
                        sql &= " ,[NetTotal])"
                        sql &= " VALUES"
                        sql &= "(" & JOURNAL_ID & ""
                        sql &= " ," & ITEM_NO & ""
                        sql &= " ," & MAT_ID & ""
                        sql &= " ," & VOLUME & ""
                        sql &= " ," & QTY & ""
                        sql &= " ," & PRICE & ""
                        sql &= " ," & VALUE & ""
                        sql &= " ," & DC_PRICE & ""
                        sql &= " ," & DC_VALUE & ""
                        sql &= " ," & TRANS_NO & ""
                        sql &= " ," & HOSE_ID & ""
                        sql &= " ," & PUMP_ID & ""
                        sql &= " ," & TANK_ID & ""
                        sql &= " ," & IS_OFFLINE & ""
                        sql &= " ," & JOURNAL_REF & ""
                        sql &= " ," & DC_ITEM_VALUE & ""
                        sql &= " ," & DC_ITEM_AMOUNT & ""
                        sql &= " ," & DC_ITEM_TYPE & ""
                        sql &= " ," & DC_BILL_VALUE & ""
                        sql &= " ," & DC_BILL_AMOUNT & ""
                        sql &= " ," & DC_BILL_TYPE & ""
                        sql &= " ," & VAT_TYPE & ""
                        sql &= " ," & VATable & ""
                        sql &= " ," & VAT & ""
                        sql &= " ," & Total & ""
                        sql &= " ," & NetBeforeVat & ""
                        sql &= " ," & NetTotal & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()
                    Next
#End Region

                Case "TSJOURNAL_PAYMENT"
#Region "TSJOURNAL_PAYMENT"
                    Dim JOURNAL_ID, ITEM_NO, GROUP_ID, PAYMENT_TYPE, VALUE, DC, CUS_ID, VEHICLE_ID, VOUCHER_NO, CARD_NO, CARD_TYPE,
                        INVOICE_NO, APPROVE_CODE, SIGNATURE, EDCNO, REDEEMVALUE, NII, LTRNPOINT, LBATCH_NO, LTRANS_NO, LSTAND_ID,
                        VOID_DATE, VOID_APPROVE_CODE, PO_NO, VEH_CODE, LCOUPON_QTY, LCOUPON_CODE As String

                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        JOURNAL_ID = IIf(dt.Rows(i)("JOURNAL_ID").ToString = "", "null", "'" & dt.Rows(i)("JOURNAL_ID").ToString & "'")
                        ITEM_NO = IIf(dt.Rows(i)("ITEM_NO").ToString = "", "null", "'" & dt.Rows(i)("ITEM_NO").ToString & "'")
                        GROUP_ID = IIf(dt.Rows(i)("GROUP_ID").ToString = "", "null", "'" & dt.Rows(i)("GROUP_ID").ToString & "'")
                        PAYMENT_TYPE = IIf(dt.Rows(i)("PAYMENT_TYPE").ToString = "", "null", "'" & dt.Rows(i)("PAYMENT_TYPE").ToString & "'")
                        VALUE = IIf(dt.Rows(i)("VALUE").ToString = "", "null", "'" & dt.Rows(i)("VALUE").ToString & "'")
                        DC = IIf(dt.Rows(i)("DC").ToString = "", "null", "'" & dt.Rows(i)("DC").ToString & "'")
                        CUS_ID = IIf(dt.Rows(i)("CUS_ID").ToString = "", "null", "'" & dt.Rows(i)("CUS_ID").ToString & "'")
                        VEHICLE_ID = IIf(dt.Rows(i)("VEHICLE_ID").ToString = "", "null", "'" & dt.Rows(i)("VEHICLE_ID").ToString & "'")
                        VOUCHER_NO = IIf(dt.Rows(i)("VOUCHER_NO").ToString = "", "null", "'" & dt.Rows(i)("VOUCHER_NO").ToString & "'")
                        CARD_NO = IIf(dt.Rows(i)("CARD_NO").ToString = "", "null", "'" & dt.Rows(i)("CARD_NO").ToString & "'")

                        CARD_TYPE = IIf(dt.Rows(i)("CARD_TYPE").ToString = "", "null", "'" & dt.Rows(i)("CARD_TYPE").ToString & "'")
                        INVOICE_NO = IIf(dt.Rows(i)("INVOICE_NO").ToString = "", "null", "'" & dt.Rows(i)("INVOICE_NO").ToString & "'")
                        APPROVE_CODE = IIf(dt.Rows(i)("APPROVE_CODE").ToString = "", "null", "'" & dt.Rows(i)("APPROVE_CODE").ToString & "'")
                        SIGNATURE = IIf(dt.Rows(i)("SIGNATURE").ToString = "", "null", "'" & dt.Rows(i)("SIGNATURE").ToString & "'")
                        EDCNO = IIf(dt.Rows(i)("EDCNO").ToString = "", "null", "'" & dt.Rows(i)("EDCNO").ToString & "'")
                        REDEEMVALUE = IIf(dt.Rows(i)("REDEEMVALUE").ToString = "", "null", "'" & dt.Rows(i)("REDEEMVALUE").ToString & "'")
                        NII = IIf(dt.Rows(i)("NII").ToString = "", "null", "'" & dt.Rows(i)("NII").ToString & "'")
                        LTRNPOINT = IIf(dt.Rows(i)("LTRNPOINT").ToString = "", "null", "'" & dt.Rows(i)("LTRNPOINT").ToString & "'")
                        LBATCH_NO = IIf(dt.Rows(i)("LBATCH_NO").ToString = "", "null", "'" & dt.Rows(i)("LBATCH_NO").ToString & "'")
                        LTRANS_NO = IIf(dt.Rows(i)("LTRANS_NO").ToString = "", "null", "'" & dt.Rows(i)("LTRANS_NO").ToString & "'")

                        LSTAND_ID = IIf(dt.Rows(i)("LSTAND_ID").ToString = "", "null", "'" & dt.Rows(i)("LSTAND_ID").ToString & "'")
                        VOID_DATE = IIf(dt.Rows(i)("VOID_DATE").ToString = "", "null", "" & dt.Rows(i)("VOID_DATE").ToString & "")
                        If VOID_DATE <> "null" Then
                            VOID_DATE = ClsClobalFunction.ConvertDateTime(VOID_DATE)
                        End If
                        VOID_APPROVE_CODE = IIf(dt.Rows(i)("VOID_APPROVE_CODE").ToString = "", "null", "'" & dt.Rows(i)("VOID_APPROVE_CODE").ToString & "'")
                        PO_NO = "null" 'IIf(dt.Rows(i)("PO_NO").ToString = "", "null", "'" & dt.Rows(i)("PO_NO").ToString & "'")
                        VEH_CODE = "null" 'IIf(dt.Rows(i)("VEH_CODE").ToString = "", "null", "'" & dt.Rows(i)("VEH_CODE").ToString & "'")
                        LCOUPON_QTY = IIf(dt.Rows(i)("LCOUPON_QTY").ToString = "", "null", "'" & dt.Rows(i)("LCOUPON_QTY").ToString & "'")
                        LCOUPON_CODE = IIf(dt.Rows(i)("LCOUPON_CODE").ToString = "", "null", "'" & dt.Rows(i)("LCOUPON_CODE").ToString & "'")

                        sql = "INSERT INTO [dbo].[TSJOURNAL_PAYMENT]"
                        sql &= "([JOURNAL_ID]"
                        sql &= " ,[ITEM_NO]"
                        sql &= " ,[GROUP_ID]"
                        sql &= " ,[PAYMENT_TYPE]"
                        sql &= " ,[VALUE]"
                        sql &= " ,[DC]"
                        sql &= " ,[CUS_ID]"
                        sql &= " ,[VEHICLE_ID]"
                        sql &= " ,[VOUCHER_NO]"
                        sql &= " ,[CARD_NO]"
                        sql &= " ,[CARD_TYPE]"
                        sql &= " ,[INVOICE_NO]"
                        sql &= " ,[APPROVE_CODE]"
                        sql &= " ,[SIGNATURE]"
                        sql &= " ,[EDCNO]"
                        sql &= " ,[REDEEMVALUE]"
                        sql &= " ,[NII]"
                        sql &= " ,[LTRNPOINT]"
                        sql &= " ,[LBATCH_NO]"
                        sql &= " ,[LTRANS_NO]"
                        sql &= " ,[LSTAND_ID]"
                        sql &= " ,[VOID_DATE]"
                        sql &= " ,[VOID_APPROVE_CODE]"
                        sql &= " ,[PO_NO]"
                        sql &= " ,[VEH_CODE]"
                        sql &= " ,[LCOUPON_QTY]"
                        sql &= " ,[LCOUPON_CODE])"
                        sql &= " VALUES"
                        sql &= "(" & JOURNAL_ID & ""
                        sql &= " ," & ITEM_NO & ""
                        sql &= " ," & GROUP_ID & ""
                        sql &= " ," & PAYMENT_TYPE & ""
                        sql &= " ," & VALUE & ""
                        sql &= " ," & DC & ""
                        sql &= " ," & CUS_ID & ""
                        sql &= " ," & VEHICLE_ID & ""
                        sql &= " ," & VOUCHER_NO & ""
                        sql &= " ," & CARD_NO & ""
                        sql &= "  ," & CARD_TYPE & ""
                        sql &= "  ," & INVOICE_NO & ""
                        sql &= " ," & APPROVE_CODE & ""
                        sql &= " ," & SIGNATURE & ""
                        sql &= "  ," & EDCNO & ""
                        sql &= " ," & REDEEMVALUE & ""
                        sql &= " ," & NII & ""
                        sql &= " ," & LTRNPOINT & ""
                        sql &= "  ," & LBATCH_NO & ""
                        sql &= "  ," & LTRANS_NO & ""
                        sql &= "  ," & LSTAND_ID & ""
                        sql &= "  ," & VOID_DATE & ""
                        sql &= "  ," & VOID_APPROVE_CODE & ""
                        sql &= "  ," & PO_NO & ""
                        sql &= "  ," & VEH_CODE & ""
                        sql &= "  ," & LCOUPON_QTY & ""
                        sql &= "  ," & LCOUPON_CODE & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()
                    Next
#End Region

                Case "TBPERIODS"
#Region "TBPERIODS"
                    Dim PERIOD_ID, POS_ID, USER_OPEN, USER_CLOSE, PERIOD_CREATE_TS, PERIOD_CLOSE_DT, PERIOD_TYPE, PERIOD_STATE,
                    DAY_ID, BUS_DATE, SHIFT_NO, TANK_DIPS_ENTERED, TANK_DROPS_ENTERED, PERIOD_METER_ENTERED, EXPORTED,
                    EXPORT_REQUIRED, WETSTOCK_OUT_OF_VARIANCE, WETSTOCK_APPRROVAL_ID, PRINT_TIMES, PUMP_ALLOW, LOGONMODE,
                    TRANSFER_STATUS, TRANSFER_DATE, MODBY As String

                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        PERIOD_ID = IIf(dt.Rows(i)("PERIOD_ID").ToString = "", "null", "'" & dt.Rows(i)("PERIOD_ID").ToString & "'")
                        POS_ID = IIf(dt.Rows(i)("POS_ID").ToString = "", "null", "'" & dt.Rows(i)("POS_ID").ToString & "'")
                        USER_OPEN = IIf(dt.Rows(i)("USER_OPEN").ToString = "", "null", "'" & dt.Rows(i)("USER_OPEN").ToString & "'")
                        USER_CLOSE = IIf(dt.Rows(i)("USER_CLOSE").ToString = "", "null", "'" & dt.Rows(i)("USER_CLOSE").ToString & "'")
                        PERIOD_CREATE_TS = IIf(dt.Rows(i)("PERIOD_CREATE_TS").ToString = "", "null", "" & dt.Rows(i)("PERIOD_CREATE_TS").ToString & "")
                        If PERIOD_CREATE_TS <> "null" Then
                            PERIOD_CREATE_TS = ClsClobalFunction.ConvertDateTime(PERIOD_CREATE_TS)
                        End If

                        PERIOD_CLOSE_DT = IIf(dt.Rows(i)("PERIOD_CLOSE_DT").ToString = "", "null", "" & dt.Rows(i)("PERIOD_CLOSE_DT").ToString & "")
                        If PERIOD_CLOSE_DT <> "null" Then
                            PERIOD_CLOSE_DT = ClsClobalFunction.ConvertDateTime(PERIOD_CLOSE_DT)
                        End If

                        PERIOD_TYPE = IIf(dt.Rows(i)("PERIOD_TYPE").ToString = "", "null", "'" & dt.Rows(i)("PERIOD_TYPE").ToString & "'")
                        PERIOD_STATE = IIf(dt.Rows(i)("PERIOD_STATE").ToString = "", "null", "'" & dt.Rows(i)("PERIOD_STATE").ToString & "'")
                        DAY_ID = IIf(dt.Rows(i)("DAY_ID").ToString = "", "null", "'" & dt.Rows(i)("DAY_ID").ToString & "'")
                        BUS_DATE = IIf(dt.Rows(i)("BUS_DATE").ToString = "", "null", "" & dt.Rows(i)("BUS_DATE").ToString & "")
                        If BUS_DATE <> "null" Then
                            BUS_DATE = ClsClobalFunction.ConvertDate(BUS_DATE)
                        End If

                        SHIFT_NO = IIf(dt.Rows(i)("SHIFT_NO").ToString = "", "null", "'" & dt.Rows(i)("SHIFT_NO").ToString & "'")
                        TANK_DIPS_ENTERED = IIf(dt.Rows(i)("TANK_DIPS_ENTERED").ToString = "", "null", "'" & dt.Rows(i)("TANK_DIPS_ENTERED").ToString & "'")
                        TANK_DROPS_ENTERED = IIf(dt.Rows(i)("TANK_DROPS_ENTERED").ToString = "", "null", "'" & dt.Rows(i)("TANK_DROPS_ENTERED").ToString & "'")
                        PERIOD_METER_ENTERED = IIf(dt.Rows(i)("PERIOD_METER_ENTERED").ToString = "", "null", "'" & dt.Rows(i)("PERIOD_METER_ENTERED").ToString & "'")
                        EXPORTED = IIf(dt.Rows(i)("EXPORTED").ToString = "", "null", "'" & dt.Rows(i)("EXPORTED").ToString & "'")
                        EXPORT_REQUIRED = IIf(dt.Rows(i)("EXPORT_REQUIRED").ToString = "", "null", "'" & dt.Rows(i)("EXPORT_REQUIRED").ToString & "'")
                        WETSTOCK_OUT_OF_VARIANCE = IIf(dt.Rows(i)("WETSTOCK_OUT_OF_VARIANCE").ToString = "", "null", "'" & dt.Rows(i)("WETSTOCK_OUT_OF_VARIANCE").ToString & "'")
                        WETSTOCK_APPRROVAL_ID = IIf(dt.Rows(i)("WETSTOCK_APPRROVAL_ID").ToString = "", "null", "'" & dt.Rows(i)("WETSTOCK_APPRROVAL_ID").ToString & "'")
                        PRINT_TIMES = IIf(dt.Rows(i)("PRINT_TIMES").ToString = "", "null", "'" & dt.Rows(i)("PRINT_TIMES").ToString & "'")
                        PUMP_ALLOW = IIf(dt.Rows(i)("PUMP_ALLOW").ToString = "", "null", "'" & dt.Rows(i)("PUMP_ALLOW").ToString & "'")

                        LOGONMODE = IIf(dt.Rows(i)("LOGONMODE").ToString = "", "null", "'" & dt.Rows(i)("LOGONMODE").ToString & "'")
                        TRANSFER_STATUS = IIf(dt.Rows(i)("TRANSFER_STATUS").ToString = "", "null", "'" & dt.Rows(i)("TRANSFER_STATUS").ToString & "'")
                        TRANSFER_DATE = IIf(dt.Rows(i)("TRANSFER_DATE").ToString = "", "null", "" & dt.Rows(i)("TRANSFER_DATE").ToString & "")
                        If TRANSFER_DATE <> "null" Then
                            TRANSFER_DATE = ClsClobalFunction.ConvertDate(TRANSFER_DATE)
                        End If

                        MODBY = IIf(dt.Rows(i)("MODBY").ToString = "", "null", "'" & dt.Rows(i)("MODBY").ToString & "'")

                        sql = "INSERT INTO [dbo].[TBPERIODS]"
                        sql &= "([PERIOD_ID]"
                        sql &= " ,[POS_ID]"
                        sql &= " ,[USER_OPEN]"
                        sql &= " ,[USER_CLOSE]"
                        sql &= " ,[PERIOD_CREATE_TS]"
                        sql &= " ,[PERIOD_CLOSE_DT]"
                        sql &= " ,[PERIOD_TYPE]"
                        sql &= " ,[PERIOD_STATE]"
                        sql &= " ,[DAY_ID]"
                        sql &= " ,[BUS_DATE]"
                        sql &= " ,[SHIFT_NO]"
                        sql &= " ,[TANK_DIPS_ENTERED]"
                        sql &= " ,[TANK_DROPS_ENTERED]"
                        sql &= " ,[PERIOD_METER_ENTERED]"
                        sql &= " ,[EXPORTED]"
                        sql &= " ,[EXPORT_REQUIRED]"
                        sql &= " ,[WETSTOCK_OUT_OF_VARIANCE]"
                        sql &= " ,[WETSTOCK_APPRROVAL_ID]"
                        sql &= " ,[PRINT_TIMES]"
                        sql &= " ,[PUMP_ALLOW]"
                        sql &= " ,[LOGONMODE]"
                        sql &= " ,[MODBY]"
                        sql &= " ,[TRANSFER_STATUS]"
                        sql &= " ,[TRANSFER_DATE])"
                        sql &= " VALUES"
                        sql &= " (" & PERIOD_ID & ""
                        sql &= " ," & POS_ID & ""
                        sql &= " ," & USER_OPEN & ""
                        sql &= " ," & USER_CLOSE & ""
                        sql &= " ," & PERIOD_CREATE_TS & ""
                        sql &= " ," & PERIOD_CLOSE_DT & ""
                        sql &= " ," & PERIOD_TYPE & ""
                        sql &= " ," & PERIOD_STATE & ""
                        sql &= " ," & DAY_ID & ""
                        sql &= " ," & BUS_DATE & ""
                        sql &= " ," & SHIFT_NO & ""
                        sql &= " ," & TANK_DIPS_ENTERED & ""
                        sql &= " ," & TANK_DROPS_ENTERED & ""
                        sql &= " ," & PERIOD_METER_ENTERED & ""
                        sql &= " ," & EXPORTED & ""
                        sql &= " ," & EXPORT_REQUIRED & ""
                        sql &= " ," & WETSTOCK_OUT_OF_VARIANCE & ""
                        sql &= " ," & WETSTOCK_APPRROVAL_ID & ""
                        sql &= " ," & PRINT_TIMES & ""
                        sql &= " ," & PUMP_ALLOW & ""
                        sql &= " ," & LOGONMODE & ""
                        sql &= " ," & MODBY & ""
                        sql &= " ," & TRANSFER_STATUS & ""
                        sql &= " ," & TRANSFER_DATE & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()
                    Next
#End Region

                Case "TBHOSE_HISTORY"
#Region "TBHOSE_HISTORY"
                    Dim HOSE_ID, PERIOD_ID, OPEN_METER_VALUE, CLOSE_METER_VALUE, OPEN_METER_VOLUME, CLOSE_METER_VOLUME,
                        POSTPAY_QUANTITY, POSTPAY_VALUE, POSTPAY_VOLUME, POSTPAY_COST, PREPAY_QUANTITY, PREPAY_VALUE,
                        PREPAY_VOLUME, PREPAY_COST, PREPAY_REFUND_QTY, PREPAY_REFUND_VAL, PREPAY_RFD_LST_QTY, PREPAY_RFD_LST_VAL,
                        PREAUTH_QUANTITY, PREAUTH_VALUE, PREAUTH_VOLUME, PREAUTH_COST, MONITOR_QUANTITY, MONITOR_VALUE,
                        MONITOR_VOLUME, MONITOR_COST, DRIVEOFFS_QUANTITY, DRIVEOFFS_VALUE, DRIVEOFFS_VOLUME, DRIVEOFFS_COST,
                        TEST_DEL_QUANTITY, TEST_DEL_VOLUME, OFFLINE_QUANTITY, OFFLINE_VOLUME, OFFLINE_VALUE, OFFLINE_COST,
                        OPEN_MECH_VOLUME, CLOSE_MECH_VOLUME, OPEN_VOLUME_TURNOVER_CORRECTION, OPEN_MONEY_TURNOVER_CORRECTION,
                        CLOSE_VOLUME_TURNOVER_CORRECTION, CLOSE_MONEY_TURNOVER_CORRECTION, OPEN_VOLUME_TURNOVER_CORRECTION2,
                        CLOSE_VOLUME_TURNOVER_CORRECTION2, PUMP_ID, MAT_ID, TANK_ID As String

                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        HOSE_ID = IIf(dt.Rows(i)("HOSE_ID").ToString = "", "null", "'" & dt.Rows(i)("HOSE_ID").ToString & "'")
                        PERIOD_ID = IIf(dt.Rows(i)("PERIOD_ID").ToString = "", "null", "'" & dt.Rows(i)("PERIOD_ID").ToString & "'")
                        OPEN_METER_VALUE = IIf(dt.Rows(i)("OPEN_METER_VALUE").ToString = "", "null", "'" & dt.Rows(i)("OPEN_METER_VALUE").ToString & "'")
                        CLOSE_METER_VALUE = IIf(dt.Rows(i)("CLOSE_METER_VALUE").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_METER_VALUE").ToString & "'")
                        OPEN_METER_VOLUME = IIf(dt.Rows(i)("OPEN_METER_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("OPEN_METER_VOLUME").ToString & "'")
                        CLOSE_METER_VOLUME = IIf(dt.Rows(i)("CLOSE_METER_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_METER_VOLUME").ToString & "'")
                        POSTPAY_QUANTITY = IIf(dt.Rows(i)("POSTPAY_QUANTITY").ToString = "", "null", "'" & dt.Rows(i)("POSTPAY_QUANTITY").ToString & "'")
                        POSTPAY_VALUE = IIf(dt.Rows(i)("POSTPAY_VALUE").ToString = "", "null", "'" & dt.Rows(i)("POSTPAY_VALUE").ToString & "'")
                        POSTPAY_VOLUME = IIf(dt.Rows(i)("POSTPAY_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("POSTPAY_VOLUME").ToString & "'")
                        POSTPAY_COST = IIf(dt.Rows(i)("POSTPAY_COST").ToString = "", "null", "'" & dt.Rows(i)("POSTPAY_COST").ToString & "'")

                        PREPAY_QUANTITY = IIf(dt.Rows(i)("PREPAY_QUANTITY").ToString = "", "null", "'" & dt.Rows(i)("PREPAY_QUANTITY").ToString & "'")
                        PREPAY_VALUE = IIf(dt.Rows(i)("PREPAY_VALUE").ToString = "", "null", "'" & dt.Rows(i)("PREPAY_VALUE").ToString & "'")
                        PREPAY_VOLUME = IIf(dt.Rows(i)("PREPAY_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("PREPAY_VOLUME").ToString & "'")
                        PREPAY_COST = IIf(dt.Rows(i)("PREPAY_COST").ToString = "", "null", "'" & dt.Rows(i)("PREPAY_COST").ToString & "'")
                        PREPAY_REFUND_QTY = IIf(dt.Rows(i)("PREPAY_REFUND_QTY").ToString = "", "null", "'" & dt.Rows(i)("PREPAY_REFUND_QTY").ToString & "'")
                        PREPAY_REFUND_VAL = IIf(dt.Rows(i)("PREPAY_REFUND_VAL").ToString = "", "null", "'" & dt.Rows(i)("PREPAY_REFUND_VAL").ToString & "'")
                        PREPAY_RFD_LST_QTY = IIf(dt.Rows(i)("PREPAY_RFD_LST_QTY").ToString = "", "null", "'" & dt.Rows(i)("PREPAY_RFD_LST_QTY").ToString & "'")
                        PREPAY_RFD_LST_VAL = IIf(dt.Rows(i)("PREPAY_RFD_LST_VAL").ToString = "", "null", "'" & dt.Rows(i)("PREPAY_RFD_LST_VAL").ToString & "'")
                        MONITOR_VOLUME = IIf(dt.Rows(i)("MONITOR_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("MONITOR_VOLUME").ToString & "'")
                        PREAUTH_QUANTITY = IIf(dt.Rows(i)("PREAUTH_QUANTITY").ToString = "", "null", "'" & dt.Rows(i)("PREAUTH_QUANTITY").ToString & "'")

                        PREAUTH_VALUE = IIf(dt.Rows(i)("PREAUTH_VALUE").ToString = "", "null", "'" & dt.Rows(i)("PREAUTH_VALUE").ToString & "'")
                        PREAUTH_VOLUME = IIf(dt.Rows(i)("PREAUTH_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("PREAUTH_VOLUME").ToString & "'")
                        PREAUTH_COST = IIf(dt.Rows(i)("PREAUTH_COST").ToString = "", "null", "'" & dt.Rows(i)("PREAUTH_COST").ToString & "'")
                        MONITOR_QUANTITY = IIf(dt.Rows(i)("MONITOR_QUANTITY").ToString = "", "null", "'" & dt.Rows(i)("MONITOR_QUANTITY").ToString & "'")
                        MONITOR_VALUE = IIf(dt.Rows(i)("MONITOR_VALUE").ToString = "", "null", "'" & dt.Rows(i)("MONITOR_VALUE").ToString & "'")
                        MONITOR_COST = IIf(dt.Rows(i)("MONITOR_COST").ToString = "", "null", "'" & dt.Rows(i)("MONITOR_COST").ToString & "'")
                        DRIVEOFFS_QUANTITY = IIf(dt.Rows(i)("DRIVEOFFS_QUANTITY").ToString = "", "null", "'" & dt.Rows(i)("DRIVEOFFS_QUANTITY").ToString & "'")
                        DRIVEOFFS_VALUE = IIf(dt.Rows(i)("DRIVEOFFS_VALUE").ToString = "", "null", "'" & dt.Rows(i)("DRIVEOFFS_VALUE").ToString & "'")
                        DRIVEOFFS_VOLUME = IIf(dt.Rows(i)("DRIVEOFFS_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("DRIVEOFFS_VOLUME").ToString & "'")
                        DRIVEOFFS_COST = IIf(dt.Rows(i)("DRIVEOFFS_COST").ToString = "", "null", "'" & dt.Rows(i)("DRIVEOFFS_COST").ToString & "'")

                        TEST_DEL_QUANTITY = IIf(dt.Rows(i)("TEST_DEL_QUANTITY").ToString = "", "null", "'" & dt.Rows(i)("TEST_DEL_QUANTITY").ToString & "'")
                        TEST_DEL_VOLUME = IIf(dt.Rows(i)("TEST_DEL_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("TEST_DEL_VOLUME").ToString & "'")
                        OFFLINE_QUANTITY = IIf(dt.Rows(i)("OFFLINE_QUANTITY").ToString = "", "null", "'" & dt.Rows(i)("OFFLINE_QUANTITY").ToString & "'")
                        OFFLINE_VOLUME = IIf(dt.Rows(i)("OFFLINE_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("OFFLINE_VOLUME").ToString & "'")
                        OFFLINE_VALUE = IIf(dt.Rows(i)("OFFLINE_VALUE").ToString = "", "null", "'" & dt.Rows(i)("OFFLINE_VALUE").ToString & "'")
                        OFFLINE_COST = IIf(dt.Rows(i)("OFFLINE_COST").ToString = "", "null", "'" & dt.Rows(i)("OFFLINE_COST").ToString & "'")
                        OPEN_MECH_VOLUME = IIf(dt.Rows(i)("OPEN_MECH_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("OPEN_MECH_VOLUME").ToString & "'")
                        CLOSE_MECH_VOLUME = IIf(dt.Rows(i)("CLOSE_MECH_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_MECH_VOLUME").ToString & "'")
                        OPEN_VOLUME_TURNOVER_CORRECTION = IIf(dt.Rows(i)("OPEN_VOLUME_TURNOVER_CORRECTION").ToString = "", "null", "'" & dt.Rows(i)("OPEN_VOLUME_TURNOVER_CORRECTION").ToString & "'")
                        OPEN_MONEY_TURNOVER_CORRECTION = IIf(dt.Rows(i)("OPEN_MONEY_TURNOVER_CORRECTION").ToString = "", "null", "'" & dt.Rows(i)("OPEN_MONEY_TURNOVER_CORRECTION").ToString & "'")

                        CLOSE_VOLUME_TURNOVER_CORRECTION = IIf(dt.Rows(i)("CLOSE_VOLUME_TURNOVER_CORRECTION").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_VOLUME_TURNOVER_CORRECTION").ToString & "'")
                        CLOSE_MONEY_TURNOVER_CORRECTION = IIf(dt.Rows(i)("CLOSE_MONEY_TURNOVER_CORRECTION").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_MONEY_TURNOVER_CORRECTION").ToString & "'")
                        OPEN_VOLUME_TURNOVER_CORRECTION2 = IIf(dt.Rows(i)("OPEN_VOLUME_TURNOVER_CORRECTION2").ToString = "", "null", "'" & dt.Rows(i)("OPEN_VOLUME_TURNOVER_CORRECTION2").ToString & "'")
                        CLOSE_VOLUME_TURNOVER_CORRECTION2 = IIf(dt.Rows(i)("CLOSE_VOLUME_TURNOVER_CORRECTION2").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_VOLUME_TURNOVER_CORRECTION2").ToString & "'")

                        Dim get_pump As String = ClsClobalFunction.GET_PUMP_ID(HOSE_ID, trans)
                        PUMP_ID = IIf(get_pump = "", "null", "'" & get_pump & "'")

                        Dim get_mat As String = ClsClobalFunction.GET_MAT_ID(HOSE_ID, trans)
                        MAT_ID = IIf(get_mat = "", "null", "'" & get_mat & "'")

                        Dim get_tank As String = ClsClobalFunction.GET_TANK_ID(HOSE_ID, trans)
                        TANK_ID = IIf(get_tank = "", "null", "'" & get_tank & "'")


                        sql = "INSERT INTO [dbo].[TBHOSE_HISTORY]"
                        sql &= "([HOSE_ID]"
                        sql &= " ,[PERIOD_ID]"
                        sql &= " ,[OPEN_METER_VALUE]"
                        sql &= " ,[CLOSE_METER_VALUE]"
                        sql &= " ,[OPEN_METER_VOLUME]"
                        sql &= " ,[CLOSE_METER_VOLUME]"
                        sql &= " ,[POSTPAY_QUANTITY]"
                        sql &= " ,[POSTPAY_VALUE]"
                        sql &= " ,[POSTPAY_VOLUME]"
                        sql &= " ,[POSTPAY_COST]"
                        sql &= " ,[PREPAY_QUANTITY]"
                        sql &= " ,[PREPAY_VALUE]"
                        sql &= " ,[PREPAY_VOLUME]"
                        sql &= " ,[PREPAY_COST]"
                        sql &= " ,[PREPAY_REFUND_QTY]"
                        sql &= " ,[PREPAY_REFUND_VAL]"
                        sql &= " ,[PREPAY_RFD_LST_QTY]"
                        sql &= " ,[PREPAY_RFD_LST_VAL]"
                        sql &= " ,[PREAUTH_QUANTITY]"
                        sql &= " ,[PREAUTH_VALUE]"
                        sql &= " ,[PREAUTH_VOLUME]"
                        sql &= " ,[PREAUTH_COST]"
                        sql &= " ,[MONITOR_QUANTITY]"
                        sql &= " ,[MONITOR_VALUE]"
                        sql &= " ,[MONITOR_VOLUME]"
                        sql &= " ,[MONITOR_COST]"
                        sql &= " ,[DRIVEOFFS_QUANTITY]"
                        sql &= " ,[DRIVEOFFS_VALUE]"
                        sql &= " ,[DRIVEOFFS_VOLUME]"
                        sql &= " ,[DRIVEOFFS_COST]"
                        sql &= " ,[TEST_DEL_QUANTITY]"
                        sql &= " ,[TEST_DEL_VOLUME]"
                        sql &= " ,[OFFLINE_QUANTITY]"
                        sql &= " ,[OFFLINE_VOLUME]"
                        sql &= " ,[OFFLINE_VALUE]"
                        sql &= " ,[OFFLINE_COST]"
                        sql &= " ,[OPEN_MECH_VOLUME]"
                        sql &= " ,[CLOSE_MECH_VOLUME]"
                        sql &= " ,[OPEN_VOLUME_TURNOVER_CORRECTION]"
                        sql &= " ,[OPEN_MONEY_TURNOVER_CORRECTION]"
                        sql &= " ,[CLOSE_VOLUME_TURNOVER_CORRECTION]"
                        sql &= " ,[CLOSE_MONEY_TURNOVER_CORRECTION]"
                        sql &= " ,[OPEN_VOLUME_TURNOVER_CORRECTION2]"
                        sql &= " ,[CLOSE_VOLUME_TURNOVER_CORRECTION2]"
                        sql &= " ,[PUMP_ID]"
                        sql &= " ,[MAT_ID]"
                        sql &= " ,[TANK_ID])"
                        sql &= "  VALUES"
                        sql &= "(" & HOSE_ID & ""
                        sql &= "  ," & PERIOD_ID & ""
                        sql &= "  ," & OPEN_METER_VALUE & ""
                        sql &= "  ," & CLOSE_METER_VALUE & ""
                        sql &= "  ," & OPEN_METER_VOLUME & ""
                        sql &= "  ," & CLOSE_METER_VOLUME & ""
                        sql &= "  ," & POSTPAY_QUANTITY & ""
                        sql &= "  ," & POSTPAY_VALUE & ""
                        sql &= "  ," & POSTPAY_VOLUME & ""
                        sql &= "  ," & POSTPAY_COST & ""
                        sql &= "  ," & PREPAY_QUANTITY & ""
                        sql &= " ," & PREPAY_VALUE & ""
                        sql &= "  ," & PREPAY_VOLUME & ""
                        sql &= "  ," & PREPAY_COST & ""
                        sql &= "  ," & PREPAY_REFUND_QTY & ""
                        sql &= "  ," & PREPAY_REFUND_VAL & ""
                        sql &= "  ," & PREPAY_RFD_LST_QTY & ""
                        sql &= "  ," & PREPAY_RFD_LST_VAL & ""
                        sql &= "  ," & PREAUTH_QUANTITY & ""
                        sql &= "  ," & PREAUTH_VALUE & ""
                        sql &= "  ," & PREAUTH_VOLUME & ""
                        sql &= "  ," & PREAUTH_COST & ""
                        sql &= "  ," & MONITOR_QUANTITY & ""
                        sql &= "  ," & MONITOR_VALUE & ""
                        sql &= "  ," & MONITOR_VOLUME & ""
                        sql &= "  ," & MONITOR_COST & ""
                        sql &= "  ," & DRIVEOFFS_QUANTITY & ""
                        sql &= "  ," & DRIVEOFFS_VALUE & ""
                        sql &= "  ," & DRIVEOFFS_VOLUME & ""
                        sql &= "  ," & DRIVEOFFS_COST & ""
                        sql &= "  ," & TEST_DEL_QUANTITY & ""
                        sql &= "  ," & TEST_DEL_VOLUME & ""
                        sql &= "  ," & OFFLINE_QUANTITY & ""
                        sql &= "  ," & OFFLINE_VOLUME & ""
                        sql &= "  ," & OFFLINE_VALUE & ""
                        sql &= "  ," & OFFLINE_COST & ""
                        sql &= "  ," & OPEN_MECH_VOLUME & ""
                        sql &= "  ," & CLOSE_MECH_VOLUME & ""
                        sql &= "   ," & OPEN_VOLUME_TURNOVER_CORRECTION & ""
                        sql &= "   ," & OPEN_MONEY_TURNOVER_CORRECTION & ""
                        sql &= "   ," & CLOSE_VOLUME_TURNOVER_CORRECTION & ""
                        sql &= "   ," & CLOSE_MONEY_TURNOVER_CORRECTION & ""
                        sql &= "   ," & OPEN_VOLUME_TURNOVER_CORRECTION2 & ""
                        sql &= "   ," & CLOSE_VOLUME_TURNOVER_CORRECTION2 & ""
                        sql &= "   ," & PUMP_ID & ""
                        sql &= "   ," & MAT_ID & ""
                        sql &= "   ," & TANK_ID & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()
                    Next
#End Region

                Case "TBMATERIAL_HISTORY"
#Region "TBMATERIAL_HISTORY"
                    Dim PERIOD_ID, MAT_ID, MAT_NAME, MAT_ID2, MAT_NAME2, MAT_NAME3, MAT_BARCODE, QTY, UOM, MOVING_AVG_PRICE, STOCK, STOCK_MIN, STOCK_MAX,
                        STOCK_LOCATION_ID, TAX_CLASS, MAT_GROUP, MAT_GROUP3, DIVISION_ID, PRICE0, PRICE1, PRICE2, PRICE3, PRICE4,
                        PRICE5, PRICE6, PRICE7, PRICE8, PRICE9, PRICE10, PRICE11, PRICE12, TIMEOFSALE, LAST_SALE, LAST_RECEIVE, BLOCK, PRICINGDATE, PRICINGMODBY,
                        LOCATION_ID, MATCOLOR, OBJ_ID, OBJ_ID_MAT_GROUP3, OBJ_ID_DIVISION_ID, CREATEDATE, MODDATE, MODBY As String

                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        PERIOD_ID = IIf(dt.Rows(i)("PERIOD_ID").ToString = "", "null", "'" & dt.Rows(i)("PERIOD_ID").ToString & "'")
                        MAT_ID = IIf(dt.Rows(i)("MAT_ID").ToString = "", "null", "'" & dt.Rows(i)("MAT_ID").ToString & "'")
                        MAT_NAME = IIf(dt.Rows(i)("MAT_NAME").ToString = "", "null", "'" & dt.Rows(i)("MAT_NAME").ToString & "'")
                        MAT_ID2 = IIf(dt.Rows(i)("MAT_ID2").ToString = "", "null", "'" & dt.Rows(i)("MAT_ID2").ToString & "'")
                        MAT_NAME2 = IIf(dt.Rows(i)("MAT_NAME2").ToString = "", "null", "'" & dt.Rows(i)("MAT_NAME2").ToString & "'")
                        MAT_NAME3 = IIf(dt.Rows(i)("MAT_NAME3").ToString = "", "null", "'" & dt.Rows(i)("MAT_NAME3").ToString & "'")
                        MAT_BARCODE = IIf(dt.Rows(i)("MAT_BARCODE").ToString = "", "null", "'" & dt.Rows(i)("MAT_BARCODE").ToString & "'")
                        QTY = IIf(dt.Rows(i)("QTY").ToString = "", "null", "'" & dt.Rows(i)("QTY").ToString & "'")
                        UOM = IIf(dt.Rows(i)("UOM").ToString = "", "null", "'" & dt.Rows(i)("UOM").ToString & "'")
                        MOVING_AVG_PRICE = IIf(dt.Rows(i)("MOVING_AVG_PRICE").ToString = "", "null", "'" & dt.Rows(i)("MOVING_AVG_PRICE").ToString & "'")

                        STOCK = IIf(dt.Rows(i)("STOCK").ToString = "", "null", "'" & dt.Rows(i)("STOCK").ToString & "'")
                        STOCK_MIN = IIf(dt.Rows(i)("STOCK_MIN").ToString = "", "null", "'" & dt.Rows(i)("STOCK_MIN").ToString & "'")
                        STOCK_MAX = IIf(dt.Rows(i)("STOCK_MAX").ToString = "", "null", "'" & dt.Rows(i)("STOCK_MAX").ToString & "'")
                        STOCK_LOCATION_ID = IIf(dt.Rows(i)("STOCK_LOCATION_ID").ToString = "", "null", "'" & dt.Rows(i)("STOCK_LOCATION_ID").ToString & "'")
                        TAX_CLASS = IIf(dt.Rows(i)("TAX_CLASS").ToString = "", "null", "'" & dt.Rows(i)("TAX_CLASS").ToString & "'")
                        MAT_GROUP = IIf(dt.Rows(i)("MAT_GROUP").ToString = "", "null", "'" & dt.Rows(i)("MAT_GROUP").ToString & "'")
                        MAT_GROUP3 = IIf(dt.Rows(i)("MAT_GROUP3").ToString = "", "null", "'" & dt.Rows(i)("MAT_GROUP3").ToString & "'")
                        DIVISION_ID = IIf(dt.Rows(i)("DIVISION_ID").ToString = "", "null", "'" & dt.Rows(i)("DIVISION_ID").ToString & "'")
                        PRICE0 = IIf(dt.Rows(i)("PRICE0").ToString = "", "null", "'" & dt.Rows(i)("PRICE0").ToString & "'")
                        PRICE1 = IIf(dt.Rows(i)("PRICE1").ToString = "", "null", "'" & dt.Rows(i)("PRICE1").ToString & "'")

                        PRICE2 = IIf(dt.Rows(i)("PRICE2").ToString = "", "null", "'" & dt.Rows(i)("PRICE2").ToString & "'")
                        PRICE3 = IIf(dt.Rows(i)("PRICE3").ToString = "", "null", "'" & dt.Rows(i)("PRICE3").ToString & "'")
                        PRICE4 = IIf(dt.Rows(i)("PRICE4").ToString = "", "null", "'" & dt.Rows(i)("PRICE4").ToString & "'")
                        PRICE5 = IIf(dt.Rows(i)("PRICE5").ToString = "", "null", "'" & dt.Rows(i)("PRICE5").ToString & "'")
                        PRICE6 = IIf(dt.Rows(i)("PRICE6").ToString = "", "null", "'" & dt.Rows(i)("PRICE6").ToString & "'")
                        PRICE7 = IIf(dt.Rows(i)("PRICE7").ToString = "", "null", "'" & dt.Rows(i)("PRICE7").ToString & "'")
                        PRICE8 = IIf(dt.Rows(i)("PRICE8").ToString = "", "null", "'" & dt.Rows(i)("PRICE8").ToString & "'")
                        PRICE9 = IIf(dt.Rows(i)("PRICE9").ToString = "", "null", "'" & dt.Rows(i)("PRICE9").ToString & "'")
                        PRICE10 = IIf(dt.Rows(i)("PRICE10").ToString = "", "null", "'" & dt.Rows(i)("PRICE10").ToString & "'")
                        PRICE11 = IIf(dt.Rows(i)("PRICE11").ToString = "", "null", "'" & dt.Rows(i)("PRICE11").ToString & "'")

                        TIMEOFSALE = IIf(dt.Rows(i)("TIMEOFSALE").ToString = "", "null", "'" & dt.Rows(i)("TIMEOFSALE").ToString & "'")
                        PRICE12 = IIf(dt.Rows(i)("PRICE12").ToString = "", "null", "'" & dt.Rows(i)("PRICE12").ToString & "'")
                        LAST_SALE = IIf(dt.Rows(i)("LAST_SALE").ToString = "", "null", "" & dt.Rows(i)("LAST_SALE").ToString & "")
                        If LAST_SALE <> "null" Then
                            LAST_SALE = ClsClobalFunction.ConvertDateTime(LAST_SALE)
                        End If

                        LAST_RECEIVE = IIf(dt.Rows(i)("LAST_RECEIVE").ToString = "", "null", "" & dt.Rows(i)("LAST_RECEIVE").ToString & "")
                        If LAST_RECEIVE <> "null" Then
                            LAST_RECEIVE = ClsClobalFunction.ConvertDate(LAST_RECEIVE)
                        End If

                        BLOCK = IIf(dt.Rows(i)("BLOCK").ToString = "", "null", "'" & dt.Rows(i)("BLOCK").ToString & "'")
                        PRICINGDATE = IIf(dt.Rows(i)("PRICINGDATE").ToString = "", "null", "" & dt.Rows(i)("PRICINGDATE").ToString & "")
                        If PRICINGDATE <> "null" Then
                            PRICINGDATE = ClsClobalFunction.ConvertDate(PRICINGDATE)
                        End If

                        PRICINGMODBY = IIf(dt.Rows(i)("PRICINGMODBY").ToString = "", "null", "'" & dt.Rows(i)("PRICINGMODBY").ToString & "'")
                        LOCATION_ID = IIf(dt.Rows(i)("LOCATION_ID").ToString = "", "null", "'" & dt.Rows(i)("LOCATION_ID").ToString & "'")
                        MATCOLOR = IIf(dt.Rows(i)("MATCOLOR").ToString = "", "null", "'" & dt.Rows(i)("MATCOLOR").ToString & "'")
                        OBJ_ID = "null" 'IIf(dt.Rows(i)("OBJ_ID").ToString = "", "null", "'" & dt.Rows(i)("OBJ_ID").ToString & "'")

                        OBJ_ID_MAT_GROUP3 = "null" 'IIf(dt.Rows(i)("OBJ_ID_MAT_GROUP3").ToString = "", "null", "'" & dt.Rows(i)("OBJ_ID_MAT_GROUP3").ToString & "'")
                        OBJ_ID_DIVISION_ID = "null" 'IIf(dt.Rows(i)("OBJ_ID_DIVISION_ID").ToString = "", "null", "'" & dt.Rows(i)("OBJ_ID_DIVISION_ID").ToString & "'")

                        CREATEDATE = IIf(dt.Rows(i)("CREATEDATE").ToString = "", "null", "" & dt.Rows(i)("CREATEDATE").ToString & "")
                        If CREATEDATE <> "null" Then
                            CREATEDATE = ClsClobalFunction.ConvertDateTime(CREATEDATE)
                        End If

                        MODDATE = IIf(dt.Rows(i)("MODDATE").ToString = "", "null", "" & dt.Rows(i)("MODDATE").ToString & "")
                        If MODDATE <> "null" Then
                            MODDATE = ClsClobalFunction.ConvertDateTime(MODDATE)
                        End If
                        MODBY = IIf(dt.Rows(i)("MODBY").ToString = "", "null", "'" & dt.Rows(i)("MODBY").ToString & "'")

                        sql = "INSERT INTO [dbo].[TBMATERIAL_HISTORY]"
                        sql &= "([PERIOD_ID]"
                        sql &= " ,[MAT_ID]"
                        sql &= " ,[MAT_NAME]"
                        sql &= " ,[MAT_ID2]"
                        sql &= " ,[MAT_NAME2]"
                        sql &= " ,[MAT_NAME3]"
                        sql &= " ,[MAT_BARCODE]"
                        sql &= " ,[QTY]"
                        sql &= " ,[UOM]"
                        sql &= " ,[MOVING_AVG_PRICE]"
                        sql &= " ,[STOCK]"
                        sql &= " ,[STOCK_MIN]"
                        sql &= " ,[STOCK_MAX]"
                        sql &= " ,[STOCK_LOCATION_ID]"
                        sql &= " ,[TAX_CLASS]"
                        sql &= " ,[MAT_GROUP]"
                        sql &= " ,[MAT_GROUP3]"
                        sql &= " ,[DIVISION_ID]"
                        sql &= " ,[PRICE0]"
                        sql &= " ,[PRICE1]"
                        sql &= " ,[PRICE2]"
                        sql &= " ,[PRICE3]"
                        sql &= " ,[PRICE4]"
                        sql &= " ,[PRICE5]"
                        sql &= " ,[PRICE6]"
                        sql &= " ,[PRICE7]"
                        sql &= " ,[PRICE8]"
                        sql &= " ,[PRICE9]"
                        sql &= " ,[PRICE10]"
                        sql &= " ,[PRICE11]"
                        sql &= " ,[PRICE12]"
                        sql &= " ,[TIMEOFSALE]"
                        sql &= " ,[LAST_SALE]"
                        sql &= " ,[LAST_RECEIVE]"
                        sql &= " ,[BLOCK]"
                        sql &= " ,[PRICINGDATE]"
                        sql &= " ,[PRICINGMODBY]"
                        sql &= " ,[LOCATION_ID]"
                        sql &= " ,[MATCOLOR]"
                        sql &= " ,[CREATEDATE]"
                        sql &= " ,[MODDATE]"
                        sql &= " ,[MODBY]"
                        sql &= " ,[OBJ_ID]"
                        sql &= " ,[OBJ_ID_MAT_GROUP3]"
                        sql &= " ,[OBJ_ID_DIVISION_ID])"
                        sql &= " VALUES"
                        sql &= "(" & PERIOD_ID & ""
                        sql &= " ," & MAT_ID & ""
                        sql &= " ," & MAT_NAME & ""
                        sql &= " ," & MAT_ID2 & ""
                        sql &= " ," & MAT_NAME2 & ""
                        sql &= " ," & MAT_NAME3 & ""
                        sql &= " ," & MAT_BARCODE & ""
                        sql &= " ," & QTY & ""
                        sql &= " ," & UOM & ""
                        sql &= " ," & MOVING_AVG_PRICE & ""
                        sql &= " ," & STOCK & ""
                        sql &= " ," & STOCK_MIN & ""
                        sql &= " ," & STOCK_MAX & ""
                        sql &= " ," & STOCK_LOCATION_ID & ""
                        sql &= " ," & TAX_CLASS & ""
                        sql &= " ," & MAT_GROUP & ""
                        sql &= " ," & MAT_GROUP3 & ""
                        sql &= " ," & DIVISION_ID & ""
                        sql &= " ," & PRICE0 & ""
                        sql &= " ," & PRICE1 & ""
                        sql &= " ," & PRICE2 & ""
                        sql &= " ," & PRICE3 & ""
                        sql &= " ," & PRICE4 & ""
                        sql &= " ," & PRICE5 & ""
                        sql &= " ," & PRICE6 & ""
                        sql &= " ," & PRICE7 & ""
                        sql &= " ," & PRICE8 & ""
                        sql &= " ," & PRICE9 & ""
                        sql &= " ," & PRICE10 & ""
                        sql &= " ," & PRICE11 & ""
                        sql &= " ," & PRICE12 & ""
                        sql &= " ," & TIMEOFSALE & ""
                        sql &= " ," & LAST_SALE & ""
                        sql &= " ," & LAST_RECEIVE & ""
                        sql &= " ," & BLOCK & ""
                        sql &= " ," & PRICINGDATE & ""
                        sql &= " ," & PRICINGMODBY & ""
                        sql &= " ," & LOCATION_ID & ""
                        sql &= " ," & MATCOLOR & ""
                        sql &= " ," & CREATEDATE & ""
                        sql &= " ," & MODDATE & ""
                        sql &= " ," & MODBY & ""
                        sql &= " ," & OBJ_ID & ""
                        sql &= " ," & OBJ_ID_MAT_GROUP3 & ""
                        sql &= " ," & OBJ_ID_DIVISION_ID & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()
                    Next
#End Region

                Case "TBTANK_HISTORY"
#Region "TBTANK_HISTORY"
                    Dim PERIOD_ID, TANK_ID, OPEN_GAUGE_VOLUME, CLOSE_GAUGE_VOLUME, OPEN_THEO_VOLUME, CLOSE_THEO_VOLUME, OPEN_DIP_VOLUME, CLOSE_DIP_VOLUME, HOSE_DEL_QUANTITY,
                        HOSE_DEL_VOLUME, HOSE_DEL_VALUE, HOSE_DEL_COST, TANK_DEL_QUANTITY, TANK_DEL_VOLUME, TANK_DEL_COST, TANK_LOSS_QUANTITY, TANK_LOSS_VOLUME,
                        TANK_TRANSFER_IN_QUANTITY, TANK_TRANSFER_IN_VOLUME, TANK_TRANSFER_OUT_QUANTITY, TANK_TRANSFER_OUT_VOLUME, DIP_FUEL_TEMP, DIP_FUEL_DENSITY,
                        OPEN_DIP_WATER_VOLUME, CLOSE_DIP_WATER_VOLUME, OPEN_GAUGE_TC_VOLUME, CLOSE_GAUGE_TC_VOLUME, OPEN_WATER_VOLUME, CLOSE_WATER_VOLUME,
                        OPEN_FUEL_DENSITY, CLOSE_FUEL_DENSITY, OPEN_FUEL_TEMP, CLOSE_FUEL_TEMP, OPEN_TANK_PROBE_STATUS_ID, CLOSE_TANK_PROBE_STATUS_ID, TANK_READINGS_DT,
                        OPEN_TANK_DELIVERY_STATE_ID, CLOSE_TANK_DELIVERY_STATE_ID, OPEN_PUMP_DELIVERY_STATE, CLOSE_PUMP_DELIVERY_STATE, OPEN_DIP_TYPE_ID,
                        CLOSE_DIP_TYPE_ID, TANK_VARIANCE_REASON_ID, QUOTED_VOLUME, MAT_ID, TANK_NAME, TANK_NUMBER As String

                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        PERIOD_ID = IIf(dt.Rows(i)("PERIOD_ID").ToString = "", "null", "'" & dt.Rows(i)("PERIOD_ID").ToString & "'")
                        TANK_ID = IIf(dt.Rows(i)("TANK_ID").ToString = "", "null", "'" & dt.Rows(i)("TANK_ID").ToString & "'")
                        OPEN_GAUGE_VOLUME = IIf(dt.Rows(i)("OPEN_GAUGE_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("OPEN_GAUGE_VOLUME").ToString & "'")
                        CLOSE_GAUGE_VOLUME = IIf(dt.Rows(i)("CLOSE_GAUGE_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_GAUGE_VOLUME").ToString & "'")
                        OPEN_THEO_VOLUME = IIf(dt.Rows(i)("OPEN_THEO_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("OPEN_THEO_VOLUME").ToString & "'")
                        CLOSE_THEO_VOLUME = IIf(dt.Rows(i)("CLOSE_THEO_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_THEO_VOLUME").ToString & "'")
                        OPEN_DIP_VOLUME = IIf(dt.Rows(i)("OPEN_DIP_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("OPEN_DIP_VOLUME").ToString & "'")
                        CLOSE_DIP_VOLUME = IIf(dt.Rows(i)("CLOSE_DIP_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_DIP_VOLUME").ToString & "'")
                        HOSE_DEL_QUANTITY = IIf(dt.Rows(i)("HOSE_DEL_QUANTITY").ToString = "", "null", "'" & dt.Rows(i)("HOSE_DEL_QUANTITY").ToString & "'")
                        HOSE_DEL_VOLUME = IIf(dt.Rows(i)("HOSE_DEL_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("HOSE_DEL_VOLUME").ToString & "'")

                        HOSE_DEL_VALUE = IIf(dt.Rows(i)("HOSE_DEL_VALUE").ToString = "", "null", "'" & dt.Rows(i)("HOSE_DEL_VALUE").ToString & "'")
                        HOSE_DEL_COST = IIf(dt.Rows(i)("HOSE_DEL_COST").ToString = "", "null", "'" & dt.Rows(i)("HOSE_DEL_COST").ToString & "'")
                        TANK_DEL_QUANTITY = IIf(dt.Rows(i)("TANK_DEL_QUANTITY").ToString = "", "null", "'" & dt.Rows(i)("TANK_DEL_QUANTITY").ToString & "'")
                        TANK_DEL_VOLUME = IIf(dt.Rows(i)("TANK_DEL_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("TANK_DEL_VOLUME").ToString & "'")
                        TANK_DEL_COST = IIf(dt.Rows(i)("TANK_DEL_COST").ToString = "", "null", "'" & dt.Rows(i)("TANK_DEL_COST").ToString & "'")
                        TANK_LOSS_QUANTITY = IIf(dt.Rows(i)("TANK_LOSS_QUANTITY").ToString = "", "null", "'" & dt.Rows(i)("TANK_LOSS_QUANTITY").ToString & "'")
                        TANK_LOSS_VOLUME = IIf(dt.Rows(i)("TANK_LOSS_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("TANK_LOSS_VOLUME").ToString & "'")
                        TANK_TRANSFER_IN_QUANTITY = IIf(dt.Rows(i)("TANK_TRANSFER_IN_QUANTITY").ToString = "", "null", "'" & dt.Rows(i)("TANK_TRANSFER_IN_QUANTITY").ToString & "'")
                        TANK_TRANSFER_IN_VOLUME = IIf(dt.Rows(i)("TANK_TRANSFER_IN_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("TANK_TRANSFER_IN_VOLUME").ToString & "'")
                        TANK_TRANSFER_OUT_QUANTITY = IIf(dt.Rows(i)("TANK_TRANSFER_OUT_QUANTITY").ToString = "", "null", "'" & dt.Rows(i)("TANK_TRANSFER_OUT_QUANTITY").ToString & "'")

                        TANK_TRANSFER_OUT_VOLUME = IIf(dt.Rows(i)("TANK_TRANSFER_OUT_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("TANK_TRANSFER_OUT_VOLUME").ToString & "'")
                        DIP_FUEL_TEMP = IIf(dt.Rows(i)("DIP_FUEL_TEMP").ToString = "", "null", "'" & dt.Rows(i)("DIP_FUEL_TEMP").ToString & "'")
                        DIP_FUEL_DENSITY = IIf(dt.Rows(i)("DIP_FUEL_DENSITY").ToString = "", "null", "'" & dt.Rows(i)("DIP_FUEL_DENSITY").ToString & "'")
                        OPEN_DIP_WATER_VOLUME = IIf(dt.Rows(i)("OPEN_DIP_WATER_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("OPEN_DIP_WATER_VOLUME").ToString & "'")
                        CLOSE_DIP_WATER_VOLUME = IIf(dt.Rows(i)("CLOSE_DIP_WATER_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_DIP_WATER_VOLUME").ToString & "'")
                        OPEN_GAUGE_TC_VOLUME = IIf(dt.Rows(i)("OPEN_GAUGE_TC_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("OPEN_GAUGE_TC_VOLUME").ToString & "'")
                        CLOSE_GAUGE_TC_VOLUME = IIf(dt.Rows(i)("CLOSE_GAUGE_TC_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_GAUGE_TC_VOLUME").ToString & "'")
                        OPEN_WATER_VOLUME = IIf(dt.Rows(i)("OPEN_WATER_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("OPEN_WATER_VOLUME").ToString & "'")
                        CLOSE_WATER_VOLUME = IIf(dt.Rows(i)("CLOSE_WATER_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_WATER_VOLUME").ToString & "'")
                        OPEN_FUEL_DENSITY = IIf(dt.Rows(i)("OPEN_FUEL_DENSITY").ToString = "", "null", "'" & dt.Rows(i)("OPEN_FUEL_DENSITY").ToString & "'")

                        CLOSE_FUEL_DENSITY = IIf(dt.Rows(i)("CLOSE_FUEL_DENSITY").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_FUEL_DENSITY").ToString & "'")
                        OPEN_FUEL_TEMP = IIf(dt.Rows(i)("OPEN_FUEL_TEMP").ToString = "", "null", "'" & dt.Rows(i)("OPEN_FUEL_TEMP").ToString & "'")
                        CLOSE_FUEL_TEMP = IIf(dt.Rows(i)("CLOSE_FUEL_TEMP").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_FUEL_TEMP").ToString & "'")
                        OPEN_TANK_PROBE_STATUS_ID = IIf(dt.Rows(i)("OPEN_TANK_PROBE_STATUS_ID").ToString = "", "null", "'" & dt.Rows(i)("OPEN_TANK_PROBE_STATUS_ID").ToString & "'")
                        CLOSE_TANK_PROBE_STATUS_ID = IIf(dt.Rows(i)("CLOSE_TANK_PROBE_STATUS_ID").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_TANK_PROBE_STATUS_ID").ToString & "'")
                        TANK_READINGS_DT = IIf(dt.Rows(i)("TANK_READINGS_DT").ToString = "", "null", "" & dt.Rows(i)("TANK_READINGS_DT").ToString & "")
                        If TANK_READINGS_DT <> "null" Then
                            TANK_READINGS_DT = ClsClobalFunction.ConvertDate(TANK_READINGS_DT)
                        End If

                        OPEN_TANK_DELIVERY_STATE_ID = IIf(dt.Rows(i)("OPEN_TANK_DELIVERY_STATE_ID").ToString = "", "null", "'" & dt.Rows(i)("OPEN_TANK_DELIVERY_STATE_ID").ToString & "'")
                        CLOSE_TANK_DELIVERY_STATE_ID = IIf(dt.Rows(i)("CLOSE_TANK_DELIVERY_STATE_ID").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_TANK_DELIVERY_STATE_ID").ToString & "'")
                        OPEN_PUMP_DELIVERY_STATE = IIf(dt.Rows(i)("OPEN_PUMP_DELIVERY_STATE").ToString = "", "null", "'" & dt.Rows(i)("OPEN_PUMP_DELIVERY_STATE").ToString & "'")
                        CLOSE_PUMP_DELIVERY_STATE = IIf(dt.Rows(i)("CLOSE_PUMP_DELIVERY_STATE").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_PUMP_DELIVERY_STATE").ToString & "'")

                        OPEN_DIP_TYPE_ID = IIf(dt.Rows(i)("OPEN_DIP_TYPE_ID").ToString = "", "null", "'" & dt.Rows(i)("OPEN_DIP_TYPE_ID").ToString & "'")
                        CLOSE_DIP_TYPE_ID = IIf(dt.Rows(i)("CLOSE_DIP_TYPE_ID").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_DIP_TYPE_ID").ToString & "'")
                        TANK_VARIANCE_REASON_ID = IIf(dt.Rows(i)("TANK_VARIANCE_REASON_ID").ToString = "", "null", "'" & dt.Rows(i)("TANK_VARIANCE_REASON_ID").ToString & "'")
                        QUOTED_VOLUME = "null" 'IIf(dt.Rows(i)("QUOTED_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("QUOTED_VOLUME").ToString & "'")

                        Dim get_mat As String = ClsClobalFunction.GET_MAT_ID_TANK(TANK_ID, trans)
                        MAT_ID = IIf(get_mat = "", "null", "'" & get_mat & "'")

                        Dim get_tank As String = ClsClobalFunction.GET_TANK_NAME(TANK_ID, trans)
                        TANK_NAME = IIf(get_tank = "", "null", "'" & get_tank & "'")

                        Dim get_tank_num As String = ClsClobalFunction.GET_TANK_NUMBER(TANK_ID, trans)
                        TANK_NUMBER = IIf(get_tank_num = "", "null", "'" & get_tank_num & "'")


                        sql = "INSERT INTO [dbo].[TBTANK_HISTORY]"
                        sql &= "([PERIOD_ID]"
                        sql &= " ,[TANK_ID]"
                        sql &= " ,[OPEN_GAUGE_VOLUME]"
                        sql &= " ,[CLOSE_GAUGE_VOLUME]"
                        sql &= " ,[OPEN_THEO_VOLUME]"
                        sql &= " ,[CLOSE_THEO_VOLUME]"
                        sql &= " ,[OPEN_DIP_VOLUME]"
                        sql &= " ,[CLOSE_DIP_VOLUME]"
                        sql &= " ,[HOSE_DEL_QUANTITY]"
                        sql &= " ,[HOSE_DEL_VOLUME]"
                        sql &= " ,[HOSE_DEL_VALUE]"
                        sql &= " ,[HOSE_DEL_COST]"
                        sql &= " ,[TANK_DEL_QUANTITY]"
                        sql &= " ,[TANK_DEL_VOLUME]"
                        sql &= " ,[TANK_DEL_COST]"
                        sql &= " ,[TANK_LOSS_QUANTITY]"
                        sql &= " ,[TANK_LOSS_VOLUME]"
                        sql &= " ,[TANK_TRANSFER_IN_QUANTITY]"
                        sql &= " ,[TANK_TRANSFER_IN_VOLUME]"
                        sql &= " ,[TANK_TRANSFER_OUT_QUANTITY]"
                        sql &= " ,[TANK_TRANSFER_OUT_VOLUME]"
                        sql &= " ,[DIP_FUEL_TEMP]"
                        sql &= " ,[DIP_FUEL_DENSITY]"
                        sql &= " ,[OPEN_DIP_WATER_VOLUME]"
                        sql &= " ,[CLOSE_DIP_WATER_VOLUME]"
                        sql &= " ,[OPEN_GAUGE_TC_VOLUME]"
                        sql &= " ,[CLOSE_GAUGE_TC_VOLUME]"
                        sql &= " ,[OPEN_WATER_VOLUME]"
                        sql &= " ,[CLOSE_WATER_VOLUME]"
                        sql &= " ,[OPEN_FUEL_DENSITY]"
                        sql &= " ,[CLOSE_FUEL_DENSITY]"
                        sql &= " ,[OPEN_FUEL_TEMP]"
                        sql &= " ,[CLOSE_FUEL_TEMP]"
                        sql &= " ,[OPEN_TANK_PROBE_STATUS_ID]"
                        sql &= " ,[CLOSE_TANK_PROBE_STATUS_ID]"
                        sql &= " ,[TANK_READINGS_DT]"
                        sql &= " ,[OPEN_TANK_DELIVERY_STATE_ID]"
                        sql &= " ,[CLOSE_TANK_DELIVERY_STATE_ID]"
                        sql &= " ,[OPEN_PUMP_DELIVERY_STATE]"
                        sql &= " ,[CLOSE_PUMP_DELIVERY_STATE]"
                        sql &= " ,[OPEN_DIP_TYPE_ID]"
                        sql &= " ,[CLOSE_DIP_TYPE_ID]"
                        sql &= " ,[TANK_VARIANCE_REASON_ID]"
                        sql &= " ,[QUOTED_VOLUME]"
                        sql &= " ,[MAT_ID]"
                        sql &= " ,[TANK_NAME]"
                        sql &= " ,[TANK_NUMBER])"
                        sql &= " VALUES"
                        sql &= "(" & PERIOD_ID & ""
                        sql &= " ," & TANK_ID & ""
                        sql &= " ," & OPEN_GAUGE_VOLUME & ""
                        sql &= " ," & CLOSE_GAUGE_VOLUME & ""
                        sql &= " ," & OPEN_THEO_VOLUME & ""
                        sql &= " ," & CLOSE_THEO_VOLUME & ""
                        sql &= " ," & OPEN_DIP_VOLUME & ""
                        sql &= " ," & CLOSE_DIP_VOLUME & ""
                        sql &= " ," & HOSE_DEL_QUANTITY & ""
                        sql &= " ," & HOSE_DEL_VOLUME & ""
                        sql &= " ," & HOSE_DEL_VALUE & ""
                        sql &= " ," & HOSE_DEL_COST & ""
                        sql &= " ," & TANK_DEL_QUANTITY & ""
                        sql &= " ," & TANK_DEL_VOLUME & ""
                        sql &= " ," & TANK_DEL_COST & ""
                        sql &= " ," & TANK_LOSS_QUANTITY & ""
                        sql &= " ," & TANK_LOSS_VOLUME & ""
                        sql &= " ," & TANK_TRANSFER_IN_QUANTITY & ""
                        sql &= " ," & TANK_TRANSFER_IN_VOLUME & ""
                        sql &= " ," & TANK_TRANSFER_OUT_QUANTITY & ""
                        sql &= " ," & TANK_TRANSFER_OUT_VOLUME & ""
                        sql &= " ," & DIP_FUEL_TEMP & ""
                        sql &= " ," & DIP_FUEL_DENSITY & ""
                        sql &= " ," & OPEN_DIP_WATER_VOLUME & ""
                        sql &= " ," & CLOSE_DIP_WATER_VOLUME & ""
                        sql &= " ," & OPEN_GAUGE_TC_VOLUME & ""
                        sql &= " ," & CLOSE_GAUGE_TC_VOLUME & ""
                        sql &= " ," & OPEN_WATER_VOLUME & ""
                        sql &= " ," & CLOSE_WATER_VOLUME & ""
                        sql &= " ," & OPEN_FUEL_DENSITY & ""
                        sql &= " ," & CLOSE_FUEL_DENSITY & ""
                        sql &= " ," & OPEN_FUEL_TEMP & ""
                        sql &= " ," & CLOSE_FUEL_TEMP & ""
                        sql &= " ," & OPEN_TANK_PROBE_STATUS_ID & ""
                        sql &= " ," & CLOSE_TANK_PROBE_STATUS_ID & ""
                        sql &= " ," & TANK_READINGS_DT & ""
                        sql &= " ," & OPEN_TANK_DELIVERY_STATE_ID & ""
                        sql &= " ," & CLOSE_TANK_DELIVERY_STATE_ID & ""
                        sql &= " ," & OPEN_PUMP_DELIVERY_STATE & ""
                        sql &= " ," & CLOSE_PUMP_DELIVERY_STATE & ""
                        sql &= " ," & OPEN_DIP_TYPE_ID & ""
                        sql &= " ," & CLOSE_DIP_TYPE_ID & ""
                        sql &= " ," & TANK_VARIANCE_REASON_ID & ""
                        sql &= " ," & QUOTED_VOLUME & ""
                        sql &= " ," & MAT_ID & ""
                        sql &= " ," & TANK_NAME & ""
                        sql &= "  ," & TANK_NUMBER & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()
                    Next
#End Region

                Case "TBPAY_IN"
#Region "TBPAY_IN"
                    Dim REFBILL_NO, TRANSFER_DATE, TRAN_DATE, BUS_DATE, SHIFT_DESCRIPTION, LAST_CLOSE_SHIFT_DT,
                    FILEPATH, FILENAME, TYPE, PAYMENT_TYPE, AMOUNTREC, AMOUNT, AMOUNT_DIFF, REMARK, STATUS_SAP,
                    STATUS, NO_SALE_STATUS, CREATEDATE, CREATEBY, UPDATEDATE, UPDATEBY As String

                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        REFBILL_NO = IIf(dt.Rows(i)("REFBILL_NO").ToString = "", "null", "'" & dt.Rows(i)("REFBILL_NO").ToString & "'")
                        TRANSFER_DATE = IIf(dt.Rows(i)("TRANSFER_DATE").ToString = "", "null", "" & dt.Rows(i)("TRANSFER_DATE").ToString & "")
                        If TRANSFER_DATE <> "null" Then
                            TRANSFER_DATE = ClsClobalFunction.ConvertDate(TRANSFER_DATE)
                        End If

                        TRAN_DATE = IIf(dt.Rows(i)("TRAN_DATE").ToString = "", "null", "" & dt.Rows(i)("TRAN_DATE").ToString & "")
                        If TRAN_DATE <> "null" Then
                            TRAN_DATE = ClsClobalFunction.ConvertDate(TRAN_DATE)
                        End If

                        BUS_DATE = IIf(dt.Rows(i)("BUS_DATE").ToString = "", "null", "" & dt.Rows(i)("BUS_DATE").ToString & "")
                        If BUS_DATE <> "null" Then
                            BUS_DATE = ClsClobalFunction.ConvertDate(BUS_DATE)
                        End If

                        Dim rs_shift_Desc As String = IIf(dt.Rows(i)("SHIFT_DESCRIPTION").ToString = "", "null", "'" & dt.Rows(i)("SHIFT_DESCRIPTION").ToString.Replace("$$", Chr(13)).Replace("&&", Chr(10)) & "'")
                        SHIFT_DESCRIPTION = rs_shift_Desc


                        LAST_CLOSE_SHIFT_DT = IIf(dt.Rows(i)("LAST_CLOSE_SHIFT_DT").ToString = "", "null", "" & dt.Rows(i)("LAST_CLOSE_SHIFT_DT").ToString & "")
                        If LAST_CLOSE_SHIFT_DT <> "null" Then
                            LAST_CLOSE_SHIFT_DT = ClsClobalFunction.ConvertDateTime(LAST_CLOSE_SHIFT_DT)
                        End If

                        FILEPATH = IIf(dt.Rows(i)("FILEPATH").ToString = "", "null", "'" & dt.Rows(i)("FILEPATH").ToString & "'")
                        FILENAME = IIf(dt.Rows(i)("FILENAME").ToString = "", "null", "'" & dt.Rows(i)("FILENAME").ToString & "'")
                        TYPE = IIf(dt.Rows(i)("TYPE").ToString = "", "null", "'" & dt.Rows(i)("TYPE").ToString & "'")
                        PAYMENT_TYPE = IIf(dt.Rows(i)("PAYMENT_TYPE").ToString = "", "null", "'" & dt.Rows(i)("PAYMENT_TYPE").ToString & "'")
                        AMOUNTREC = IIf(dt.Rows(i)("AMOUNTREC").ToString = "", "null", "'" & dt.Rows(i)("AMOUNTREC").ToString & "'")
                        AMOUNT = IIf(dt.Rows(i)("AMOUNT").ToString = "", "null", "'" & dt.Rows(i)("AMOUNT").ToString & "'")
                        AMOUNT_DIFF = IIf(dt.Rows(i)("AMOUNT_DIFF").ToString = "", "null", "'" & dt.Rows(i)("AMOUNT_DIFF").ToString & "'")

                        Dim rs_remark As String = IIf(dt.Rows(i)("REMARK").ToString = "", "null", "'" & dt.Rows(i)("REMARK").ToString.Replace("$$", Chr(13)).Replace("&&", Chr(10)) & "'")
                        REMARK = rs_remark
                        STATUS_SAP = IIf(dt.Rows(i)("STATUS_SAP").ToString = "", "null", "'" & dt.Rows(i)("STATUS_SAP").ToString & "'")
                        STATUS = IIf(dt.Rows(i)("STATUS").ToString = "", "null", "'" & dt.Rows(i)("STATUS").ToString & "'")
                        NO_SALE_STATUS = IIf(dt.Rows(i)("NO_SALE_STATUS").ToString = "", "null", "'" & dt.Rows(i)("NO_SALE_STATUS").ToString & "'")
                        CREATEDATE = IIf(dt.Rows(i)("CREATEDATE").ToString = "", "null", "" & dt.Rows(i)("CREATEDATE").ToString & "")
                        If CREATEDATE <> "null" Then
                            CREATEDATE = ClsClobalFunction.ConvertDateTime(CREATEDATE)
                        End If
                        CREATEBY = IIf(dt.Rows(i)("CREATEBY").ToString = "", "null", "'" & dt.Rows(i)("CREATEBY").ToString & "'")

                        UPDATEDATE = IIf(dt.Rows(i)("UPDATEDATE").ToString = "", "null", "" & dt.Rows(i)("UPDATEDATE").ToString & "")
                        If UPDATEDATE <> "null" Then
                            UPDATEDATE = ClsClobalFunction.ConvertDateTime(UPDATEDATE)
                        End If
                        UPDATEBY = IIf(dt.Rows(i)("UPDATEBY").ToString = "", "null", "'" & dt.Rows(i)("UPDATEBY").ToString & "'")

                        sql = "INSERT INTO [dbo].[TBPAY_IN]"
                        sql &= "([REFBILL_NO]"
                        sql &= " ,[TRANSFER_DATE]"
                        sql &= " ,[TRAN_DATE]"
                        sql &= " ,[BUS_DATE]"
                        sql &= " ,[SHIFT_DESCRIPTION]"
                        sql &= " ,[LAST_CLOSE_SHIFT_DT]"
                        sql &= " ,[FILEPATH]"
                        sql &= " ,[FILENAME]"
                        sql &= " ,[TYPE]"
                        sql &= " ,[PAYMENT_TYPE]"
                        sql &= " ,[AMOUNTREC]"
                        sql &= " ,[AMOUNT]"
                        sql &= " ,[AMOUNT_DIFF]"
                        sql &= " ,[REMARK]"
                        sql &= " ,[STATUS_SAP]"
                        sql &= " ,[STATUS]"
                        sql &= " ,[NO_SALE_STATUS]"
                        sql &= " ,[CREATEDATE]"
                        sql &= " ,[CREATEBY]"
                        sql &= " ,[UPDATEDATE]"
                        sql &= " ,[UPDATEBY])"
                        sql &= " VALUES"
                        sql &= "(" & REFBILL_NO & ""
                        sql &= " ," & TRANSFER_DATE & ""
                        sql &= " ," & TRAN_DATE & ""
                        sql &= " ," & BUS_DATE & ""
                        sql &= " ," & SHIFT_DESCRIPTION & ""
                        sql &= " ," & LAST_CLOSE_SHIFT_DT & ""
                        sql &= " ," & FILEPATH & ""
                        sql &= " ," & FILENAME & ""
                        sql &= " ," & TYPE & ""
                        sql &= " ," & PAYMENT_TYPE & ""
                        sql &= " ," & AMOUNTREC & ""
                        sql &= " ," & AMOUNT & ""
                        sql &= " ," & AMOUNT_DIFF & ""
                        sql &= " ," & REMARK & ""
                        sql &= " ," & STATUS_SAP & ""
                        sql &= " ," & STATUS & ""
                        sql &= " ," & NO_SALE_STATUS & ""
                        sql &= " ," & CREATEDATE & ""
                        sql &= " ," & CREATEBY & ""
                        sql &= " ," & UPDATEDATE & ""
                        sql &= " ," & UPDATEBY & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()
                    Next
#End Region

                Case "TBPAYIN_PERIOD_LOG"
#Region "TBPAYIN_PERIOD_LOG"
                    Dim PAYIN_ID, POS_ID, BUS_DATE, SHIFT_START, SHIFT_END, TYPE As String

                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        PAYIN_ID = IIf(dt.Rows(i)("PAYIN_ID").ToString = "", "null", "'" & dt.Rows(i)("PAYIN_ID").ToString & "'")
                        POS_ID = IIf(dt.Rows(i)("POS_ID").ToString = "", "null", "" & dt.Rows(i)("POS_ID").ToString & "")

                        BUS_DATE = IIf(dt.Rows(i)("BUS_DATE").ToString = "", "null", "" & dt.Rows(i)("BUS_DATE").ToString & "")
                        If BUS_DATE <> "null" Then
                            BUS_DATE = ClsClobalFunction.ConvertDate(BUS_DATE)
                        End If
                        SHIFT_START = IIf(dt.Rows(i)("SHIFT_START").ToString = "", "null", "" & dt.Rows(i)("SHIFT_START").ToString & "")
                        SHIFT_END = IIf(dt.Rows(i)("SHIFT_END").ToString = "", "null", "" & dt.Rows(i)("SHIFT_END").ToString & "")
                        TYPE = IIf(dt.Rows(i)("TYPE").ToString = "", "null", "" & dt.Rows(i)("TYPE").ToString & "")

                        sql = "INSERT INTO [dbo].[TBPAYIN_PERIOD_LOG]"
                        sql &= "([PAYIN_ID]"
                        sql &= " ,[POS_ID]"
                        sql &= " ,[BUS_DATE]"
                        sql &= " ,[SHIFT_START]"
                        sql &= " ,[SHIFT_END]"
                        sql &= " ,[TYPE])"
                        sql &= " VALUES"
                        sql &= " (" & PAYIN_ID & ""
                        sql &= " ," & POS_ID & ""
                        sql &= " ," & BUS_DATE & ""
                        sql &= " ," & SHIFT_START & ""
                        sql &= " ," & SHIFT_END & ""
                        sql &= " ," & TYPE & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()
                    Next
#End Region

                Case "TBCARD"
#Region "TBCARD"
                    Dim MAGNETIC_CARD, NAME, CARD_TYPE, SOLD_TO, SHIP_TO, COUNTER, VEHICLE_ID, DEPARTMENT, DEPARTMENT1, VALIDDATE, EXPIREDATE, REMARK, BLOCK,
                        EMPLOYEE_NUMBER, DESCRIPTION, HOLDER, TELEPHONE, ADDRESS, TAX, CAR_TYPE, BRAND, FILE_NAME, PERIOD_ID, CREATEDATE, MODDATE, MODBY As String

                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        MAGNETIC_CARD = IIf(dt.Rows(i)("MAGNETIC_CARD").ToString = "", "null", "'" & dt.Rows(i)("MAGNETIC_CARD").ToString & "'")
                        NAME = IIf(dt.Rows(i)("NAME").ToString = "", "null", "'" & dt.Rows(i)("NAME").ToString & "'")
                        CARD_TYPE = IIf(dt.Rows(i)("CARD_TYPE").ToString = "", "null", "'" & dt.Rows(i)("CARD_TYPE").ToString & "'")
                        SOLD_TO = IIf(dt.Rows(i)("SOLD_TO").ToString = "", "null", "'" & dt.Rows(i)("SOLD_TO").ToString & "'")
                        SHIP_TO = IIf(dt.Rows(i)("SHIP_TO").ToString = "", "null", "'" & dt.Rows(i)("SHIP_TO").ToString & "'")
                        COUNTER = IIf(dt.Rows(i)("COUNTER").ToString = "", "null", "'" & dt.Rows(i)("COUNTER").ToString & "'")
                        VEHICLE_ID = IIf(dt.Rows(i)("VEHICLE_ID").ToString = "", "null", "'" & dt.Rows(i)("VEHICLE_ID").ToString & "'")
                        DEPARTMENT = IIf(dt.Rows(i)("DEPARTMENT").ToString = "", "null", "'" & dt.Rows(i)("DEPARTMENT").ToString & "'")
                        DEPARTMENT1 = IIf(dt.Rows(i)("DEPARTMENT1").ToString = "", "null", "'" & dt.Rows(i)("DEPARTMENT1").ToString & "'")

                        VALIDDATE = IIf(dt.Rows(i)("VALIDDATE").ToString = "", "null", "" & dt.Rows(i)("VALIDDATE").ToString & "")
                        If VALIDDATE <> "null" Then
                            VALIDDATE = ClsClobalFunction.ConvertDate(VALIDDATE)
                        End If

                        EXPIREDATE = IIf(dt.Rows(i)("EXPIREDATE").ToString = "", "null", "'" & dt.Rows(i)("EXPIREDATE").ToString & "'")

                        'EXPIREDATE = IIf(dt.Rows(i)("EXPIREDATE").ToString = "", "null", "" & dt.Rows(i)("EXPIREDATE").ToString & "")
                        'If EXPIREDATE <> "null" Then
                        '    EXPIREDATE = ConvertDate(EXPIREDATE)
                        'End If

                        REMARK = IIf(dt.Rows(i)("REMARK").ToString = "", "null", "'" & dt.Rows(i)("REMARK").ToString & "'")
                        BLOCK = IIf(dt.Rows(i)("BLOCK").ToString = "", "null", "'" & dt.Rows(i)("BLOCK").ToString & "'")
                        EMPLOYEE_NUMBER = IIf(dt.Rows(i)("EMPLOYEE_NUMBER").ToString = "", "null", "'" & dt.Rows(i)("EMPLOYEE_NUMBER").ToString & "'")
                        DESCRIPTION = IIf(dt.Rows(i)("DESCRIPTION").ToString = "", "null", "'" & dt.Rows(i)("DESCRIPTION").ToString & "'")
                        HOLDER = IIf(dt.Rows(i)("HOLDER").ToString = "", "null", "'" & dt.Rows(i)("HOLDER").ToString & "'")
                        TELEPHONE = IIf(dt.Rows(i)("TELEPHONE").ToString = "", "null", "'" & dt.Rows(i)("TELEPHONE").ToString & "'")
                        ADDRESS = IIf(dt.Rows(i)("ADDRESS").ToString = "", "null", "'" & dt.Rows(i)("ADDRESS").ToString & "'")
                        TAX = IIf(dt.Rows(i)("TAX").ToString = "", "null", "'" & dt.Rows(i)("TAX").ToString & "'")
                        CAR_TYPE = IIf(dt.Rows(i)("CAR_TYPE").ToString = "", "null", "'" & dt.Rows(i)("CAR_TYPE").ToString & "'")
                        BRAND = IIf(dt.Rows(i)("BRAND").ToString = "", "null", "'" & dt.Rows(i)("BRAND").ToString & "'")

                        FILE_NAME = IIf(dt.Rows(i)("FILE_NAME").ToString = "", "null", "'" & dt.Rows(i)("FILE_NAME").ToString & "'")
                        PERIOD_ID = IIf(dt.Rows(i)("PERIOD_ID").ToString = "", "null", "'" & dt.Rows(i)("PERIOD_ID").ToString & "'")


                        CREATEDATE = IIf(dt.Rows(i)("CREATEDATE").ToString = "", "null", "" & dt.Rows(i)("CREATEDATE").ToString & "")
                        If CREATEDATE <> "null" Then
                            CREATEDATE = ClsClobalFunction.ConvertDateTime(CREATEDATE)
                        End If

                        MODDATE = IIf(dt.Rows(i)("MODDATE").ToString = "", "null", "" & dt.Rows(i)("MODDATE").ToString & "")
                        If MODDATE <> "null" Then
                            MODDATE = ClsClobalFunction.ConvertDateTime(MODDATE)
                        End If
                        MODBY = IIf(dt.Rows(i)("MODBY").ToString = "", "null", "'" & dt.Rows(i)("MODBY").ToString & "'")

                        sql = "INSERT INTO [dbo].[TBCARD]"
                        sql &= "([MAGNETIC_CARD]"
                        sql &= " ,[NAME]"
                        sql &= " ,[CARD_TYPE]"
                        sql &= " ,[SOLD_TO]"
                        sql &= " ,[SHIP_TO]"
                        sql &= " ,[COUNTER]"
                        sql &= " ,[VEHICLE_ID]"
                        sql &= " ,[DEPARTMENT]"
                        sql &= " ,[DEPARTMENT1]"
                        sql &= " ,[VALIDDATE]"
                        sql &= " ,[EXPIREDATE]"
                        sql &= " ,[REMARK]"
                        sql &= " ,[BLOCK]"
                        sql &= " ,[EMPLOYEE_NUMBER]"
                        sql &= " ,[DESCRIPTION]"
                        sql &= " ,[HOLDER]"
                        sql &= " ,[TELEPHONE]"
                        sql &= " ,[ADDRESS]"
                        sql &= " ,[TAX]"
                        sql &= " ,[CAR_TYPE]"
                        sql &= " ,[BRAND]"
                        sql &= " ,[FILE_NAME]"
                        sql &= " ,[PERIOD_ID]"
                        sql &= " ,[CREATEDATE]"
                        sql &= " ,[MODDATE]"
                        sql &= " ,[MODBY])"
                        sql &= " VALUES"
                        sql &= " (" & MAGNETIC_CARD & ""
                        sql &= " ," & NAME & ""
                        sql &= " ," & CARD_TYPE & ""
                        sql &= " ," & SOLD_TO & ""
                        sql &= " ," & SHIP_TO & ""
                        sql &= " ," & COUNTER & ""
                        sql &= " ," & VEHICLE_ID & ""
                        sql &= " ," & DEPARTMENT & ""
                        sql &= " ," & DEPARTMENT1 & ""
                        sql &= " ," & VALIDDATE & ""
                        sql &= " ," & EXPIREDATE & ""
                        sql &= " ," & REMARK & ""
                        sql &= " ," & BLOCK & ""
                        sql &= " ," & EMPLOYEE_NUMBER & ""
                        sql &= " ," & DESCRIPTION & ""
                        sql &= " ," & HOLDER & ""
                        sql &= " ," & TELEPHONE & ""
                        sql &= " ," & ADDRESS & ""
                        sql &= " ," & TAX & ""
                        sql &= " ," & CAR_TYPE & ""
                        sql &= " ," & BRAND & ""
                        sql &= " ," & FILE_NAME & ""
                        sql &= " ," & PERIOD_ID & ""
                        sql &= " ," & CREATEDATE & ""
                        sql &= " ," & MODDATE & ""
                        sql &= " ," & MODBY & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()
                    Next
#End Region

                Case "TBMATERIAL_HISTORY_DESC"
#Region "TBMATERIAL_HISTORY_DESC"
                    Dim MAT_ID, PERIOD_ID, CODE, QTY, REMARK, MODDATE As String

                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        MAT_ID = IIf(dt.Rows(i)("MAT_ID").ToString = "", "null", "'" & dt.Rows(i)("MAT_ID").ToString & "'")
                        PERIOD_ID = IIf(dt.Rows(i)("PERIOD_ID").ToString = "", "null", "'" & dt.Rows(i)("PERIOD_ID").ToString & "'")
                        CODE = IIf(dt.Rows(i)("CODE").ToString = "", "null", "'" & dt.Rows(i)("CODE").ToString & "'")
                        QTY = IIf(dt.Rows(i)("QTY").ToString = "", "null", "'" & dt.Rows(i)("QTY").ToString & "'")
                        REMARK = IIf(dt.Rows(i)("REMARK").ToString = "", "null", "'" & dt.Rows(i)("REMARK").ToString & "'")
                        MODDATE = IIf(dt.Rows(i)("MODDATE").ToString = "", "null", "" & dt.Rows(i)("MODDATE").ToString & "")
                        If MODDATE <> "null" Then
                            MODDATE = ClsClobalFunction.ConvertDateTime(MODDATE)
                        End If

                        sql = "INSERT INTO [dbo].[TBMATERIAL_HISTORY_DESC]"
                        sql &= "([MAT_ID]"
                        sql &= ",[PERIOD_ID]"
                        sql &= ",[CODE]"
                        sql &= ",[QTY]"
                        sql &= ",[REMARK]"
                        sql &= ",[MODDATE])"
                        sql &= " VALUES"
                        sql &= "(" & MAT_ID & ""
                        sql &= "," & PERIOD_ID & ""
                        sql &= "," & CODE & ""
                        sql &= "," & QTY & ""
                        sql &= "," & REMARK & ""
                        sql &= "," & MODDATE & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()
                    Next
#End Region

                Case "TSSAFTDROP"
#Region "TSSAFTDROP"
                    Dim POS_ID, DAY_ID, SHIFT_ID, CRMONEY, SDMONEY, SAFETYPE, REMARK, CREATEDATE, MODDATE, MODBY As String

                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        POS_ID = IIf(dt.Rows(i)("POS_ID").ToString = "", "null", "'" & dt.Rows(i)("POS_ID").ToString & "'")
                        DAY_ID = IIf(dt.Rows(i)("DAY_ID").ToString = "", "null", "'" & dt.Rows(i)("DAY_ID").ToString & "'")
                        SHIFT_ID = IIf(dt.Rows(i)("SHIFT_ID").ToString = "", "null", "'" & dt.Rows(i)("SHIFT_ID").ToString & "'")
                        CRMONEY = IIf(dt.Rows(i)("CRMONEY").ToString = "", "null", "'" & dt.Rows(i)("CRMONEY").ToString & "'")
                        SDMONEY = IIf(dt.Rows(i)("SDMONEY").ToString = "", "null", "'" & dt.Rows(i)("SDMONEY").ToString & "'")
                        SAFETYPE = IIf(dt.Rows(i)("SAFETYPE").ToString = "", "null", "'" & dt.Rows(i)("SAFETYPE").ToString & "'")
                        REMARK = IIf(dt.Rows(i)("REMARK").ToString = "", "null", "'" & dt.Rows(i)("REMARK").ToString & "'")
                        CREATEDATE = IIf(dt.Rows(i)("CREATEDATE").ToString = "", "null", "" & dt.Rows(i)("CREATEDATE").ToString & "")
                        If CREATEDATE <> "null" Then
                            CREATEDATE = ClsClobalFunction.ConvertDateTime(CREATEDATE)
                        End If

                        MODDATE = IIf(dt.Rows(i)("MODDATE").ToString = "", "null", "" & dt.Rows(i)("MODDATE").ToString & "")
                        If MODDATE <> "null" Then
                            MODDATE = ClsClobalFunction.ConvertDateTime(MODDATE)
                        End If
                        MODBY = IIf(dt.Rows(i)("MODBY").ToString = "", "null", "'" & dt.Rows(i)("MODBY").ToString & "'")

                        sql = "INSERT INTO [dbo].[TSSAFTDROP]"
                        sql &= "([POS_ID]"
                        sql &= " ,[DAY_ID]"
                        sql &= " ,[SHIFT_ID]"
                        sql &= " ,[CRMONEY]"
                        sql &= " ,[SDMONEY]"
                        sql &= " ,[SAFETYPE]"
                        sql &= " ,[REMARK]"
                        sql &= " ,[CREATEDATE]"
                        sql &= " ,[MODDATE]"
                        sql &= " ,[MODBY])"
                        sql &= " VALUES"
                        sql &= " (" & POS_ID & ""
                        sql &= " ," & DAY_ID & ""
                        sql &= " ," & SHIFT_ID & ""
                        sql &= " ," & CRMONEY & ""
                        sql &= " ," & SDMONEY & ""
                        sql &= " ," & SAFETYPE & ""
                        sql &= " ," & REMARK & ""
                        sql &= " ," & CREATEDATE & " "
                        sql &= " ," & MODDATE & ""
                        sql &= " ," & MODBY & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()
                    Next
#End Region

            End Select

            trans.Commit()
            conn.Close()
            Return ""
        Catch ex As Exception
            trans.Rollback()
            Return "พบปัญหาในการนำเข้าข้อมูล : Table " & TableName & ":" & ex.ToString & " sql: " & sql
        End Try
    End Function

End Class
