Public Class clsBackgroundWorker1
    Inherits System.ComponentModel.BackgroundWorker

#Region "Variable"
    Private _DataGridView1 As DataGridView
    'Private strNetNodeUnit As String = "1.12.0"

    Private m_clsPlcWrapper As clsPlcWrapper


#End Region


    Public ReadOnly Property plcWrapper() As clsPlcWrapper
        Get
            Return m_clsPlcWrapper
        End Get
        'Set(ByVal value As clsPlcWrapper)
        '    m_clsPlcWrapper = value
        'End Set
    End Property


    Public Property DataGridView1() As DataGridView
        Get
            Return _DataGridView1
        End Get
        Set(ByVal value As DataGridView)
            _DataGridView1 = value
        End Set
    End Property


    Private Sub ReadPLC2Datagridview()

        'Dim m_clsPlcWrapper As clsPlcWrapper = New clsPlcWrapper
        'm_clsPlcWrapper = New clsPlcWrapper
        'Dim arrNetNodeUnit As String() = netNodeUnit.Split(".")
        'm_clsPlcWrapper.InitializeAndConnect(arrNetNodeUnit(0), arrNetNodeUnit(1), arrNetNodeUnit(2))
        'MsgBox(m_clsPlcWrapper.ReadMemoryWordIntegerDMBCD0(11001))

        For Each r As DataGridViewRow In _DataGridView1.Rows
            If Not (r.Cells(EXCEL_HEADER_address).Value Is Nothing) AndAlso Not String.IsNullOrEmpty(r.Cells(EXCEL_HEADER_address).Value.ToString) Then
                Dim strAddress As String = r.Cells(EXCEL_HEADER_address).Value.ToString
                strAddress = strAddress.Split("-").First

                If strAddress <> "" And Not (strAddress Like "*-*") Then
                    Dim strPLC As String = ""
                    'Dim dataType As String = UCase(r.Cells("Data Type").Value.ToString)
                    'Dim dataType As String = UCase(r.Cells(excelHeader.datatype).Value.ToString)
                    Dim dataType As String = UCase(r.Cells(EXCEL_HEADER_datatype).Value.ToString)
                    If dataType = EXCEL_DATATYPE_UINT Then
                        strPLC = m_clsPlcWrapper.ReadMemoryWordIntegerDMBCD0(CInt(strAddress)).ToString

                    ElseIf dataType = EXCEL_DATATYPE_REAL Then

                        'MsgBox(m_clsPlcWrapper.ReadMemory(OMRON.Compolet.SYSMAC.SysmacCSBase.MemoryTypes.DM, CInt(strAddress), 2, True))
                        'strPLC = m_clsPlcWrapper.ReadMemory(OMRON.Compolet.SYSMAC.SysmacCSBase.MemoryTypes.DM, CInt(strAddress), 2, True).ToString
                        strPLC = m_clsPlcWrapper.ReadMemoryDwordSingle(OMRON.Compolet.SYSMAC.SysmacCSBase.MemoryTypes.DM, CInt(strAddress), r.Cells(EXCEL_HEADER_totalWords).Value)(0).ToString
                        'm_clsPlcWrapper.ReadMemoryWordIntegerDMBCD0(CInt(strAddress)).ToString()

                    ElseIf dataType = EXCEL_DATATYPE_ASCII Then
                        'strPLC = m_clsPlcWrapper.ReadMemoryString(OMRON.Compolet.SYSMAC.SysmacCSBase.MemoryTypes.DM, CInt(strAddress), r.Cells(EXCEL_HEADER_totalWords).Value)(0).ToString
                        strPLC = m_clsPlcWrapper.ReadMemoryString(OMRON.Compolet.SYSMAC.SysmacCSBase.MemoryTypes.DM, CInt(strAddress), r.Cells(EXCEL_HEADER_totalWords).Value).ToString
                    Else
                        Continue For

                    End If

                    If r.Index = -1 Then

                        Continue For
                    End If

                    r.Cells(EXCEL_HEADER_currentValue).Value = strPLC
                    'r.SetValues()

                    ''r.Cells(7).Value = "5"
                    'r.Cells("Value_from_PLC").Value = "44"
                    'r.Cells("Value_from_PLC").Value = "aaa44"
                    ''r.Cells("Value_from_PLC").d = True
                    'r.Cells(6).ToolTipText = "aaa44"
                    ''r.Cells("Value From PLC 2").Value = strPLC

                End If
            End If

        Next r

        '_DataGridView1.Refresh()





        'Dim intRow = DataGridView1.Rows.Add()
        'Dim rowToAdd As DataRow = DataGridView1.Item(7, intRow).

        'DataGridView1.Item(7, intRow).Value = "dqqqq"






    End Sub

    Private Sub clsBackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles Me.DoWork

        While True
            ReadPLC2Datagridview()
            'System.Threading.Thread.Sleep(100000)
            System.Threading.Thread.Sleep(1000)
        End While

    End Sub

    Public Sub New()
        m_clsPlcWrapper = New clsPlcWrapper
        'Dim arrNetNodeUnit As String() = strNetNodeUnit.Split(".")
        Dim arrNetNodeUnit As String() = strFinsAddress.Split(".")
        m_clsPlcWrapper.InitializeAndConnect(arrNetNodeUnit(0), arrNetNodeUnit(1), arrNetNodeUnit(2))
    End Sub
End Class
