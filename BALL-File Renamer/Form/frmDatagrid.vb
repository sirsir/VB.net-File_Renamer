Public Class frmDatagrid
    'Protected m_clsPlcWrapper As clsPlcWrapper = New clsPlcWrapper()
    Protected m_clsPlcWrapper As clsPlcWrapper

    'm_clsPlcWrapper.InitializeAndConnect(1, 12, 0)

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        PrepareDatagridview()
        StartReadPLC()
    End Sub

    Private Sub PrepareDatagridview()
        'Dim filepath As String = "C:\Users\Administrator\Desktop\VB_PROJECTS\BALL\Ball\Ball\Resources\PLC1.xls"

        Dim filepath As String = My.Application.Info.DirectoryPath & "\Resources\PLC1.xls"
        clsExcel.LoadExcelDataToDataGrid(filepath, Me.DataGridView1)


        'Dim col As New DataGridViewTextBoxColumn
        'col.ValueType = Type.GetType("System.String")
        'col.Visible = True
        'col.ReadOnly = False
        'col.ToolTipText = "dddddqqqq"
        ''col.DataPropertyName = "PropertyName"
        'col.HeaderText = "Value From PLC2"
        'col.Name = "Value_from_PLC"
        ''Me.DataGridView1.Columns.Add(col)
        'Me.DataGridView1.Columns.Add(col)


        Dim btn As New DataGridViewButtonColumn()
        DataGridView1.Columns.Add(btn)
        btn.HeaderText = "Edit PLC value"
        btn.Text = "Edit PLC value"
        btn.Name = "btnEditPLCvalue"
        btn.UseColumnTextForButtonValue = True


        'Me.DataGridView1.Columns.Add("Value_from_PLC", EXCEL_HEADER_currentValue)


    End Sub

    Private Sub StartReadPLC()
        Dim bw1 As clsBackgroundWorker1 = New clsBackgroundWorker1()
        m_clsPlcWrapper = bw1.plcWrapper

        bw1.DataGridView1 = Me.DataGridView1

        bw1.RunWorkerAsync()


    End Sub

    'Private Sub StartReadPLC2()

    '    'Dim m_clsPlcWrapper As clsPlcWrapper = New clsPlcWrapper
    '    m_clsPlcWrapper.InitializeAndConnect(1, 12, 0)
    '    'MsgBox(m_clsPlcWrapper.ReadMemoryWordIntegerDMBCD0(11001))

    '    For Each r As DataGridViewRow In Me.DataGridView1.Rows
    '        If Not (r.Cells(2).Value Is Nothing) AndAlso Not String.IsNullOrEmpty(r.Cells(2).Value.ToString) Then
    '            Dim strAddress As String = r.Cells(2).Value.ToString
    '            strAddress = strAddress.Split("-").First

    '            If strAddress <> "" And Not (strAddress Like "*-*") Then
    '                Dim strPLC As String = ""
    '                If UCase(r.Cells(4).Value.ToString) = EXCEL_DATATYPE_UINT Then
    '                    strPLC = m_clsPlcWrapper.ReadMemoryWordIntegerDMBCD0(CInt(strAddress)).ToString

    '                ElseIf UCase(r.Cells(4).Value.ToString) = EXCEL_DATATYPE_REAL Then

    '                    'MsgBox(m_clsPlcWrapper.ReadMemory(OMRON.Compolet.SYSMAC.SysmacCSBase.MemoryTypes.DM, CInt(strAddress), 2, True))
    '                    'strPLC = m_clsPlcWrapper.ReadMemory(OMRON.Compolet.SYSMAC.SysmacCSBase.MemoryTypes.DM, CInt(strAddress), 2, True).ToString
    '                    strPLC = m_clsPlcWrapper.ReadMemoryDwordSingle(OMRON.Compolet.SYSMAC.SysmacCSBase.MemoryTypes.DM, CInt(strAddress), r.Cells("Total Words").Value)(0).ToString
    '                    'm_clsPlcWrapper.ReadMemoryWordIntegerDMBCD0(CInt(strAddress)).ToString()

    '                ElseIf UCase(r.Cells(4).Value.ToString) = EXCEL_DATATYPE_ASCII Then
    '                    strPLC = m_clsPlcWrapper.ReadMemoryString(OMRON.Compolet.SYSMAC.SysmacCSBase.MemoryTypes.DM, CInt(strAddress), r.Cells("Total Words").Value)(0).ToString

    '                End If
    '                r.Cells(EXCEL_HEADER_currentValue).Value = strPLC

    '                ''r.Cells(7).Value = "5"
    '                'r.Cells("Value_from_PLC").Value = "44"
    '                'r.Cells("Value_from_PLC").Value = "aaa44"
    '                ''r.Cells("Value_from_PLC").d = True
    '                'r.Cells(6).ToolTipText = "aaa44"
    '                ''r.Cells("Value From PLC 2").Value = strPLC

    '            End If
    '        End If

    '    Next r

    '    DataGridView1.Refresh()





    '    'Dim intRow = DataGridView1.Rows.Add()
    '    'Dim rowToAdd As DataRow = DataGridView1.Item(7, intRow).

    '    'DataGridView1.Item(7, intRow).Value = "dqqqq"






    'End Sub




    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Try
            If DataGridView1.Columns(e.ColumnIndex).Name = "btnEditPLCvalue" Then
                Dim strOldvalue As String = DataGridView1.Rows(e.RowIndex).Cells(EXCEL_HEADER_currentValue).Value
                Dim strValue As String = InputBox("Input value to write to this address.", "Enter value:", strOldvalue)
                If strValue <> strOldvalue Then
                    Dim strAddress As String = DataGridView1.Rows(e.RowIndex).Cells(EXCEL_HEADER_address).Value
                    strAddress = strAddress.Split("-").First


                    If UCase(DataGridView1.Rows(e.RowIndex).Cells(EXCEL_HEADER_datatype).Value.ToString) = EXCEL_DATATYPE_UINT Then
                        'DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).
                        'strPLC = m_clsPlcWrapper.ReadMemoryWordIntegerDMBCD0(CInt(strAddress)).ToString
                        m_clsPlcWrapper.WriteMemoryWordIntegerDMBCD0(CInt(strAddress), CInt(strValue))

                    ElseIf UCase(DataGridView1.Rows(e.RowIndex).Cells(EXCEL_HEADER_datatype).Value.ToString) = EXCEL_DATATYPE_REAL Then


                        ' strPLC = m_clsPlcWrapper.ReadMemoryDwordSingle(OMRON.Compolet.SYSMAC.SysmacCSBase.MemoryTypes.DM, CInt(strAddress), r.Cells("Total Words").Value)(0).ToString
                        m_clsPlcWrapper.WriteMemoryDwordSingle(OMRON.Compolet.SYSMAC.SysmacCSBase.MemoryTypes.DM, CInt(strAddress), CDbl(strValue))


                    ElseIf UCase(DataGridView1.Rows(e.RowIndex).Cells(EXCEL_HEADER_datatype).Value.ToString) = EXCEL_DATATYPE_ASCII Then
                        ' strPLC = m_clsPlcWrapper.ReadMemoryString(OMRON.Compolet.SYSMAC.SysmacCSBase.MemoryTypes.DM, CInt(strAddress), r.Cells("Total Words").Value)(0).ToString
                        If strValue = "" Then
                            m_clsPlcWrapper.WriteMemoryWordIntegerDMBCD0(CInt(strAddress), 0)

                        Else
                            m_clsPlcWrapper.WriteMemoryString(OMRON.Compolet.SYSMAC.SysmacCSBase.MemoryTypes.DM, CInt(strAddress), strValue)
                        End If




                    End If

                End If

            End If
        Catch ex As Exception

        End Try
        

    End Sub
End Class