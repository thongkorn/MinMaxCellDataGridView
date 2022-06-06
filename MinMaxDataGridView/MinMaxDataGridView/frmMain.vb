#Region "ABOUT"
' / --------------------------------------------------------------------
' / Developer : Mr.Surapon Yodsanga (Thongkorn Tubtimkrob)
' / eMail : thongkorn@hotmail.com
' / URL: http://www.g2gnet.com (Khon Kaen - Thailand)
' / Facebook: https://www.facebook.com/g2gnet (For Thailand)
' / Facebook: https://www.facebook.com/commonindy (Worldwide)
' / More Info: http://www.g2gsoft.com/
' /
' / Purpose: Set Min/Max Value in Column DataGridView and Validate Cell.
' / Microsoft Visual Basic .NET (2010)
' /
' / This is open source code under @Copyleft by Thongkorn Tubtimkrob.
' / You can modify and/or distribute without to inform the developer.
' / --------------------------------------------------------------------
#End Region

Public Class frmMain
    '// START HERE
    Private Sub frmMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Call InitializeGrid(dgvData)
        Call FillSampleData()
    End Sub

    ' / --------------------------------------------------------------------------------
    '// Default settings for Grids @Run Time
    Private Sub InitializeGrid(ByRef dgv As DataGridView)
        With dgv
            .RowHeadersVisible = False
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .AllowUserToResizeRows = False
            .MultiSelect = False
            .SelectionMode = DataGridViewSelectionMode.CellSelect
            .ReadOnly = False
            .Font = New Font("Tahoma", 10)
            .RowHeadersVisible = True
            .RowTemplate.MinimumHeight = 27
            .RowTemplate.Height = 27
            .AlternatingRowsDefaultCellStyle.BackColor = Color.SkyBlue
            .DefaultCellStyle.SelectionBackColor = Color.OrangeRed
            '/ Auto size column width of each main by sorting the field.
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            '/ Adjust Header Styles
            With .ColumnHeadersDefaultCellStyle
                .BackColor = Color.Navy
                .ForeColor = Color.White
                .Font = New Font("Tahoma", 10, FontStyle.Bold)
            End With
        End With
        dgvData.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        dgvData.ColumnHeadersHeight = 36
        '// กำหนดให้ EnableHeadersVisualStyles = False เพื่อให้ยอมรับการเปลี่ยนแปลงสีพื้นหลังของ Header
        dgvData.EnableHeadersVisualStyles = False
    End Sub

    ' / --------------------------------------------------------------------------------
    ' / SAMPLE DATA INTO DATAGRIDVIEW
    Private Sub FillSampleData()
        Dim DT As New DataTable
        DT.Columns.Add("Integer Column")
        DT.Columns.Add("Double Column")
        Dim RandomClass As New Random()
        For iRow As Long = 0 To 19
            Dim DR As DataRow = DT.NewRow()
            DR(0) = RandomClass.Next(1, 1000)   '// Random Integer Value 1 - 1000
            DR(1) = Format(RandomClass.NextDouble * 1000, "0.00")   '// Random Double Value.
            DT.Rows.Add(DR)
        Next
        dgvData.DataSource = DT
        DT.Dispose()
    End Sub

    ' / --------------------------------------------------------------------------------
    ' / เหตุการณ์ในการตรวจสอบเงื่อนไข หลังจากการป้อนค่าลงไปและกด Enter
    Private Sub dgvData_CellEndEdit(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvData.CellEndEdit
        Dim NewValue As Double = 0
        '// เช็ค Null Value หากไม่ป้อนค่าใดๆ ก็กำหนดค่าต่ำสุดเป็น 0
        If IsDBNull(dgvData.Rows.Item(e.RowIndex).Cells(e.ColumnIndex).Value) Then
            NewValue = 0    '// ค่าต่ำสุด
            '// มีการป้อนค่าตัวเลข ให้นำค่าไปเก็บไว้ในตัวแปร NewValue ก่อน
        Else
            NewValue = dgvData.Rows.Item(e.RowIndex).Cells(e.ColumnIndex).Value
        End If
        '// ตรวจสอบเงื่อนไขคือ 0 - 1000 (ที่จริง 0 ตัดทิ้งไปได้ เพราะเรา Validate Cell เป็นค่าต่ำสุดไว้อยู่แล้ว)
        If NewValue >= 0 AndAlso NewValue <= 1000 Then
            '// ค่าอยู่ในช่วง 0 - 1000 ให้นำค่าใหม่มาใส่แทน
            dgvData.Rows.Item(e.RowIndex).Cells(e.ColumnIndex).Value = NewValue
        Else
            '// ค่าไม่อยู่ในช่วง 0 - 1000 ให้นำค่าเดิมมาแทนที่ (ค่าเดิมเก็บไว้ที่ dgvData.Tag จากเหตุการณ์ dgvData_CellEnter)
            dgvData.Rows.Item(e.RowIndex).Cells(e.ColumnIndex).Value = dgvData.Tag
        End If
        '// เงื่อนไขจาก ColumnIndex เพื่อปรับชนิดข้อมูล (Convert) 
        Select Case e.ColumnIndex
            Case 0  '// Integer
                dgvData.Rows.Item(e.RowIndex).Cells(e.ColumnIndex).Value = Format(CInt(dgvData.Rows.Item(e.RowIndex).Cells(e.ColumnIndex).Value), "0")
            Case 1  '// Double
                dgvData.Rows.Item(e.RowIndex).Cells(e.ColumnIndex).Value = Format(CDbl(dgvData.Rows.Item(e.RowIndex).Cells(e.ColumnIndex).Value), "0.00")
        End Select
    End Sub

    Private Sub dgvData_CellEnter(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvData.CellEnter
        '// นำค่าเดิมไปเก็บไว้ในคุณสมบัติ Tag ของ DataGridView โดยไม่ต้องประกาศใช้งานตัวแปรใหม่
        '// ปกติเราจะป้อนข้อมูลได้ทีละเซลล์อยู่แล้ว ก็ให้โฟกัสไปที่แถวและหลักนั้นๆ
        dgvData.Tag = dgvData.Rows.Item(e.RowIndex).Cells(e.ColumnIndex).Value
    End Sub

    ' / --------------------------------------------------------------------------------
    ' / เหตุการณ์ในการกดคีย์ เลขจำนวนเต็มจะรับค่าได้ 0 - 9 ส่วนเลขทศนิยมจะเพิ่มเครื่องหมาย . มาได้เพียงตัวเดียวเท่านั้น
    Private Sub dgvData_EditingControlShowing(sender As Object, e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles dgvData.EditingControlShowing
        Select Case dgvData.Columns(dgvData.CurrentCell.ColumnIndex).Index
            '// ColumeIndex 0 is Integer and ColumnIndex 1 is Double.
            Case 0, 1
                '// Force to validate value at ValidKeyPress Event.
                RemoveHandler e.Control.KeyPress, AddressOf ValidKeyPress
                AddHandler e.Control.KeyPress, AddressOf ValidKeyPress
        End Select
    End Sub

    ' / --------------------------------------------------------------------------------
    Private Sub ValidKeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs)
        Dim tb As TextBox = sender
        Select Case dgvData.CurrentCell.ColumnIndex
            Case 0  '// Integer
                Select Case e.KeyChar
                    Case "0" To "9"   ' digits 0 - 9 allowed
                    Case ChrW(Keys.Back)    ' backspace allowed for deleting (Delete key automatically overrides)
                    Case Else ' everything else ....
                        '// True = CPU cancel the KeyPress event
                        e.Handled = True '// and it's just like you never pressed a key at all
                End Select

            Case 1  '// Double
                Select Case e.KeyChar
                    Case "0" To "9"
                        '// Allowed "."
                    Case "."
                        '// But it can present "." only one.
                        If InStr(tb.Text, ".") Then e.Handled = True

                    Case ChrW(Keys.Back)
                    Case Else
                        e.Handled = True
                End Select
        End Select
    End Sub

End Class
