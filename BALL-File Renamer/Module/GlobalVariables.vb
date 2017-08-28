Imports System.Configuration
Imports Newtonsoft.Json.Linq
Imports System.IO

Module GlobalVariables

#Region "App.config"
    Dim appSet As New AppSettingsReader()
    'Public strFinsAddress As String = "1.12.0"

    Public strFinsAddress As String = appSet.GetValue("FinsAddress", GetType(String))

#End Region
    
#Region "config.json"

    Public strJsonPath As String = "config.json"
    Public Sub RefreshJson()
        varFromJson = JObject.Parse(File.ReadAllText(strJsonPath))

    End Sub

    Public strUserInput As String = ""
    Public strOutput As String = ""


    'Public varFromJson As JObject = JObject.Parse(File.ReadAllText("C:\AnEasyBrowseDir\Temp\config.json"))
    'Public varFromJson As JObject = JObject.Parse(File.ReadAllText("Resources\config.json"))
    Public varFromJson As JObject = JObject.Parse(File.ReadAllText(strJsonPath))

    'Public EXCEL_HEADER_datatype As String = "ddd"
    'Public EXCEL_HEADER_datatype As String = varFromJson.Item("excelDependent").Item("datatype").Item("heading")

    'Public excelHeaderDatatype As String
    ' excelHeaderDatatype = varFromJson.Item("excelDependent").Item("datatype").Item("heading")

    'Public EXCEL_HEADER_address As String = varFromJson.Item("excelDependent").Item("address").Item("heading")
    'Public EXCEL_HEADER_datatype As String = varFromJson.Item("excelDependent").Item("datatype").Item("heading")
    'Public EXCEL_HEADER_totalWords As String = varFromJson.Item("excelDependent").Item("total words").Item("heading")
    'Public EXCEL_HEADER_currentValue As String = varFromJson.Item("excelDependent").Item("current value").Item("heading")


    'Public EXCEL_DATATYPE_UINT As String = varFromJson.Item("strings").Item("datatype").Item("uint")
    'Public EXCEL_DATATYPE_REAL As String = varFromJson.Item("strings").Item("datatype").Item("real")
    'Public EXCEL_DATATYPE_ASCII As String = varFromJson.Item("strings").Item("datatype").Item("ascii")

    Public frmPleaseWait1 As New frmPleaseWait


#End Region

    


End Module
