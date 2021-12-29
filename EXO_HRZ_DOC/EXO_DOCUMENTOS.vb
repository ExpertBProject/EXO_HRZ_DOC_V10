Imports SAPbouiCOM
Imports System.Xml
Public Class EXO_DOCUMENTOS
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)
        If actualizar Then
            CargarScripts()
        End If
    End Sub
    Private Sub CargarScripts()
        Dim sScript1 As String = ""

        If objGlobal.refDi.comunes.esAdministrador Then
            Try
                If objGlobal.compañia.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    sScript1 = objGlobal.funciones.leerEmbebido(Me.GetType(), "HANA_VALIDAR_NIF_CIF.sql")
                Else
                    sScript1 = objGlobal.funciones.leerEmbebido(Me.GetType(), "SQL_VALIDAR_NIF_CIF.sql")
                End If

                objGlobal.refDi.SQL.sqlUpdB1(sScript1)

            Catch exCOM As System.Runtime.InteropServices.COMException
                Throw exCOM
            Catch ex As Exception
                Throw ex
            End Try
        End If
    End Sub
    Public Overrides Function filtros() As EventFilters
        Dim filtrosXML As Xml.XmlDocument = New Xml.XmlDocument
        filtrosXML.LoadXml(objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_FILTROS.xml"))
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(filtrosXML.OuterXml)

        Return filtro
    End Function
    Public Overrides Function menus() As XmlDocument
        Return Nothing
    End Function
    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "139", "140", "133", "142", "143", "141"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                                    If EventHandler_Validate_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK


                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "139", "140", "133", "142", "143", "141"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_Before(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "139", "140", "133", "142", "143", "141"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "139", "140", "133", "142", "143", "141"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                End If
            End If

            Return MyBase.SBOApp_ItemEvent(infoEvento)
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        End Try
    End Function
    Private Function EventHandler_Validate_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sCardCode As String = ""
        Dim sMensaje As String = ""
        EventHandler_Validate_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "54" Then ' Or pVal.ItemUID = "4" Then
                sCardCode = CType(oForm.Items.Item("4").Specific, SAPbouiCOM.EditText).Value.ToString.Trim
                If sCardCode <> "" Then
                    sMensaje = objGlobal.refDi.SQL.sqlStringB1("SELECT ""Notes"" FROM ""OCRD"" Where ""CardCode""='" & sCardCode & "' ")
                    objGlobal.SBOApp.MessageBox(sMensaje)
                Else
                    ' objGlobal.SBOApp.StatusBar.SetText("(EXO) - No se encuentra el cod. de Ic.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                End If
            End If
            EventHandler_Validate_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_ItemPressed_Before(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim iTamMatrix As Integer = 0
        Dim bSelLinea As Boolean = False
        Dim sMensaje As String = ""
        Dim sCIF As String = "" : Dim sIndicator As String = ""
        EventHandler_ItemPressed_Before = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If pVal.ItemUID = "1" And pVal.FormTypeEx = "133" Then
                If oForm.Mode = BoFormMode.fm_ADD_MODE Or oForm.Mode = BoFormMode.fm_UPDATE_MODE Then
                    sCIF = CType(oForm.Items.Item("123").Specific, SAPbouiCOM.EditText).Value.ToString.Trim
                    If CType(oForm.Items.Item("120").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sIndicator = CType(oForm.Items.Item("120").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                    Else
                        sIndicator = ""
                    End If


                    If sCIF.Trim = "" Then
                        sMensaje = "El campo ""Número de identificación fiscal"" no puede estar vacío. Por favor, compruebe el dato."
                        objGlobal.SBOApp.MessageBox(sMensaje)
                        Exit Function
                    Else
                        If Left(sCIF, 2) = "ES" And sIndicator = "01" Then
                            EventHandler_ItemPressed_Before = True
                        Else
                            If Left(sCIF.Trim, 2) = "ES" Then
                                EventHandler_ItemPressed_Before = Comprobar_CIF_NIF(objGlobal, sCIF)
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If

            EventHandler_ItemPressed_Before = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Public Shared Function Comprobar_CIF_NIF(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByVal sValor As String) As Boolean
        Comprobar_CIF_NIF = False
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sSQL As String = ""
        Try
            oRs = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            'Validamos el CIF o NIF
            If oObjGlobal.compañia.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                sSQL = "SELECT ""EXO_VALIDAR_NIF_CIF""(RTRIM(LTRIM('" & sValor & "'))) ""Es_CIFNIF_OK"" FROM DUMMY;"
            Else
                sSQL = "SELECT [dbo].[EXO_VALIDAR_NIF_CIF](RTRIM(LTRIM('" & sValor & "'))) ""Es_CIFNIF_OK"" "
            End If
            oRs.DoQuery(sSQL)

            If oRs.RecordCount > 0 Then
                If CInt(oRs.Fields.Item("Es_CIFNIF_OK").Value.ToString) = 0 Then
                    oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - El CIF/NIF " & sValor & " no es válido.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oObjGlobal.SBOApp.MessageBox("El CIF/NIF " & sValor & " no es válido.")
                    Exit Function
                End If
            Else
                Throw New Exception("No se ha encontrado función EXO_VALIDAR_NIF_CIF")
                Exit Function
            End If
            Comprobar_CIF_NIF = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
End Class
