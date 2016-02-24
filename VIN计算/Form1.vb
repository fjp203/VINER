
Imports System.Collections.Generic
Imports System.Diagnostics

Imports System.Data
Imports System.Text.RegularExpressions '引入正则表达式命名空间
Imports System.IO.IsolatedStorage


Public Class Form1

    '定义了一个到处函数
    Public Function daochu(ByVal x As DataGridView, ByVal filename As String, ByVal n As Integer) As Boolean '导出到Excel函数
        Try
            If x.Rows.Count <= 0 Then '判断记录数,如果没有记录就退出
                MessageBox.Show("没有记录可以导出", "没有可以导出的项目", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return False
            Else '如果有记录就导出到Excel
                Dim xx As Object : Dim yy As Object
                xx = CreateObject("Excel.Application") '创建Excel对象
                yy = xx.workbooks.add()
                Dim i As Integer, u As Integer = 0, v As Integer = 0 '定义循环变量,行列变量
                For i = 1 To x.Columns.Count '把表头写入Excel
                    yy.worksheets(n).cells(1, i) = x.Columns(i - 1).HeaderCell.Value
                Next
                Dim str(x.Rows.Count - 1, x.Columns.Count - 1) '定义一个二维数组
                For u = 1 To x.Rows.Count '行循环
                    For v = 1 To x.Columns.Count '列循环
                        If x.Item(v - 1, u - 1).Value.GetType.ToString <> "System.Guid" Then
                            str(u - 1, v - 1) = x.Item(v - 1, u - 1).Value
                        Else
                            str(u - 1, v - 1) = x.Item(v - 1, u - 1).Value.ToString
                        End If
                    Next
                Next
                yy.worksheets(n).range("A2").Resize(x.Rows.Count, x.Columns.Count).Value = str '把数组一起写入Excel
                yy.worksheets(n).Cells.EntireColumn.AutoFit() '自动调整Excel列
                'yy.worksheets(1).name = x.TopLeftHeaderCell.Value.ToString '表标题写入作为Excel工作表名称
                xx.visible = False '设置Excel可见


                Select Case xx.Workbooks(1).worksheets(1).UsedRange.Columns.Count
                    Case 4
                        xx.Workbooks(1).worksheets(1).UsedRange.font.size = 17
                        xx.Workbooks(1).worksheets(1).UsedRange.Borders(1).LineStyle = 1
                        xx.Workbooks(1).worksheets(1).UsedRange.Borders(2).LineStyle = 1
                        xx.Workbooks(1).worksheets(1).UsedRange.Borders(3).LineStyle = 1
                        xx.Workbooks(1).worksheets(1).UsedRange.Borders(4).LineStyle = 1
                        xx.Workbooks(1).worksheets(1).UsedRange.EntireColumn.AutoFit()


                        xx.Workbooks(1).worksheets(1).Columns("A:A").HorizontalAlignment = 3
                        xx.Workbooks(1).worksheets(1).Columns("A:A").NumberFormatLocal = "000000"
                        xx.Workbooks(1).worksheets(1).Columns("C:C").HorizontalAlignment = 3
                        xx.Workbooks(1).worksheets(1).Columns("C:C").NumberFormatLocal = "000000"
                    Case 2
                        xx.Workbooks(1).worksheets(1).UsedRange.font.size = 12
                        xx.Workbooks(1).worksheets(1).UsedRange.Borders(1).LineStyle = 1
                        xx.Workbooks(1).worksheets(1).UsedRange.Borders(2).LineStyle = 1
                        xx.Workbooks(1).worksheets(1).UsedRange.Borders(3).LineStyle = 1
                        xx.Workbooks(1).worksheets(1).UsedRange.Borders(4).LineStyle = 1
                        xx.Workbooks(1).worksheets(1).UsedRange.EntireColumn.AutoFit()
                End Select


              
                xx.Workbooks(1).SaveCopyAs(filename) '保存

                yy = Nothing '销毁组建释放资源
                xx = Nothing '销毁组建释放资源




            End If
                Return True
        Catch ex As Exception '错误处理
            MessageBox.Show(Err.Description.ToString, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error) '出错提示
            Return False
        End Try
    End Function










    '定义一个VIN码计算函数
    Public Function VIN(ByVal q8 As String, ByVal h8 As String)
        ' 定义a(17)为加权系数
        Dim a() As Integer = {8, 7, 6, 5, 4, 3, 2, 10, 0, 9, 8, 7, 6, 5, 4, 3, 2}
        Dim b(17) As String ' 定义b(17)为16各种字母加加权位
        Dim z As String
        z = ""
        If q8 <> "" And h8 <> "" Then
            Dim str17 As String
            str17 = q8 + "X" + h8

            For i = 0 To 16
                b(i) = Mid(str17, i + 1, 1)
            Next

            Dim c() As String = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F", "G", "H", "J", "K", "L", "M", "N", "P", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
            Dim d() As Integer = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 1, 2, 3, 4, 5, 6, 7, 8, 1, 2, 3, 4, 5, 7, 9, 2, 3, 4, 5, 6, 7, 8, 9}
            Dim f() As Integer = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}
            For p = 0 To 16
                For q = 0 To 32
                    If b(p) = c(q) Then
                        f(p) = d(q)
                    End If
                Next
            Next
            Dim sum, x As Integer
            sum = 0
            x = 0

            For i = 0 To 16
                sum = sum + a(i) * f(i)

            Next
            x = sum Mod 11


            If x = 10 Then
                z = q8 + "X" + h8

            Else

                z = q8 + x.ToString + h8
            End If

        End If
        Return z

        LabelState.Text = "行:0"
        TLabel2.Text = ""


    End Function
    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        '首先清空datagrid
        DataGridView1.Columns.Clear()


        'Dim myStream As System.IO.Stream

        OpenFileDialog1.InitialDirectory = My.Computer.FileSystem.CurrentDirectory
        OpenFileDialog1.Filter = "Excel 2003(*.xls)|*.xls|Excel 2007(*.xlsa)|*.xlsa|所有文件 (*.*)|*.*"
        OpenFileDialog1.FilterIndex = 1
        OpenFileDialog1.Title = "请选择VIN码导入文件"
        OpenFileDialog1.RestoreDirectory = True
        OpenFileDialog1.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory
        OpenFileDialog1.FileName = ""
        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Dim fileName As String
            fileName = Me.OpenFileDialog1.FileName
            '建立EXCEL连接，读入数据
            Dim myDataset As New DataSet

            Dim strConn As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & fileName & "'; Extended Properties='Excel 8.0;HDR=YES;IMEX=2'"

            Dim da As New OleDb.OleDbDataAdapter("SELECT * FROM [Sheet1$A:D]  where len(trim(物料编码))>0", strConn)
            Try
                da.Fill(myDataset)
                Me.DataGridView1.DataSource = myDataset.Tables(0)

                ToolStripButton4.Enabled = True  '可以计算
                '******载入后对列宽度进行重新规定*****'
                DataGridView1.Columns(0).Width = 140
                DataGridView1.Columns(1).Width = 103
                DataGridView1.Columns(2).Width = 165
                DataGridView1.Columns(3).Width = 165
            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try


        End If


    End Sub

    Private Sub DT1_CtMStrip1_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs)

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick

    End Sub

    Private Sub DataGridView1_Resize(sender As Object, e As EventArgs) Handles DataGridView1.Resize


    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LabelState.Text = "欢迎使用VIN码编制软件!技术支持:工艺科范建鹏,电话:83388352。"
        ToolStripTextBox2.Width = WebBrowser1.Width - ToolStripButton20.Width - 20

        WebBrowser1.Url = New Uri("file:///" + AppDomain.CurrentDomain.BaseDirectory.ToString.Replace("\", "/") + "res/shouce/首页.html")



    End Sub

    Private Sub 复制CToolStripMenuItem_Click(sender As Object, e As EventArgs)
        System.Windows.Forms.Clipboard.Clear()

        System.Windows.Forms.Clipboard.SetDataObject(DataGridView1.CurrentCell.Value)
    End Sub

    Private Sub ToolStripButton9_Click(sender As Object, e As EventArgs)
        Dim i, j As Integer
        Dim pRow, pCol As Integer
        Dim selectedCellCount As Integer
        Dim startRow, startCol, endRow, endCol As Integer
        Dim pasteText, strline, strVal As String
        Dim strlines, vals As String()
        Dim pasteData(,) As String
        Dim flag As Boolean = False
        ' 当前单元格是否选择的判断
        If DataGridView1.CurrentCell Is Nothing Then
            Return
        End If
        Dim insertRowIndex As Integer = DataGridView1.CurrentCell.RowIndex
        ' 获取DataGridView选择区域，并计算要复制的行列开始、结束位置
        startRow = 9999
        startCol = 9999
        endRow = 0
        endCol = 0
        selectedCellCount = DataGridView1.GetCellCount(DataGridViewElementStates.Selected)
        For i = 0 To selectedCellCount - 1
            startRow = Math.Min(DataGridView1.SelectedCells(i).RowIndex, startRow)
            startCol = Math.Min(DataGridView1.SelectedCells(i).ColumnIndex, startCol)
            endRow = Math.Max(DataGridView1.SelectedCells(i).RowIndex, endRow)
            endCol = Math.Max(DataGridView1.SelectedCells(i).ColumnIndex, endCol)
        Next
        ' 获取剪切板的内容，并按行分割
        pasteText = Clipboard.GetText()
        If String.IsNullOrEmpty(pasteText) Then
            Return
        End If
        pasteText = pasteText.Replace(vbCrLf, vbLf)
        ReDim strlines(0)
        strlines = pasteText.Split(vbLf)
        pRow = strlines.Length        '行数
        pCol = 0
        For Each strline In strlines
            ReDim vals(0)
            vals = strline.Split(New Char() {vbTab, vbCr, vbNullChar, vbNullString}, 256, StringSplitOptions.RemoveEmptyEntries) ' 按Tab分割数据
            pCol = Math.Max(vals.Length, pCol) '列数
        Next
        ReDim pasteData(pRow, pCol)
        pasteText = Clipboard.GetText()
        pasteText = pasteText.Replace(vbCrLf, vbLf)
        ReDim strlines(0)
        strlines = pasteText.Split(vbLf)
        i = 1
        For Each strline In strlines
            j = 1
            ReDim vals(0)
            strline.TrimEnd(New Char() {vbLf})
            vals = strline.Split(New Char() {vbTab, vbCr, vbNullChar, vbNullString}, 256, StringSplitOptions.RemoveEmptyEntries)
            For Each strVal In vals
                pasteData(i, j) = strVal
                j = j + 1
            Next
            i = i + 1
        Next
        flag = False
        For j = 1 To pCol
            If pasteData(pRow, j) <> "" Then
                flag = True
                Exit For
            End If
        Next
        If flag = False Then
            pRow = Math.Max(pRow - 1, 0)
        End If

        For i = 1 To endRow - startRow + 1
            Dim row As DataGridViewRow = DataGridView1.Rows(i + startRow - 1)
            If i <= pRow Then
                For j = 1 To endCol - startCol + 1
                    If j <= pCol Then
                        row.Cells(j + startCol - 1).Value = pasteData(i, j)
                    Else
                        Exit For
                    End If
                Next
            Else
                Exit For
            End If
        Next

        '清除剪切板原有内容，将表格数据复制到剪切板
        System.Windows.Forms.Clipboard.Clear()
        System.Windows.Forms.Clipboard.SetDataObject(DataGridView1.GetClipboardContent())
    End Sub

    Private Sub ToolStripButton6_Click(sender As Object, e As EventArgs) Handles ToolStripButton6.Click

        DataGridView1.Rows.Remove(DataGridView1.CurrentRow)

        Try
            TLabel2.Text = "（第" + (DataGridView1.CurrentCell.RowIndex + 1).ToString + "行,第" + (DataGridView1.CurrentCell.ColumnIndex + 1).ToString + "列)" + DataGridView1.CurrentCell.ErrorText
            TLabel2.Text = TLabel2.Text.Replace(Microsoft.VisualBasic.Constants.vbCrLf, "@")
        Catch ex As Exception

        End Try

    End Sub

    Private Sub 删除DToolStripMenuItem_Click(sender As Object, e As EventArgs)
        DataGridView1.Rows.Remove(DataGridView1.CurrentRow)

        Try
            TLabel2.Text = "（第" + (DataGridView1.CurrentCell.RowIndex + 1).ToString + "行,第" + (DataGridView1.CurrentCell.ColumnIndex + 1).ToString + "列)" + DataGridView1.CurrentCell.ErrorText
            TLabel2.Text = TLabel2.Text.Replace(Microsoft.VisualBasic.Constants.vbCrLf, "@")
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click

        DataGridView2.Rows.Clear() '每次计算前先清空datagridview2 普通
        DataGridView3.Rows.Clear() '每次计算前先清空datagridview3 MES
        DataGridView4.Rows.Clear() '每次计算前先清空datagridview4 打印
        For i = 0 To DataGridView1.RowCount - 1

            Dim cx As String = ""  '车型
            Dim VIN8 As String = "" 'ＶＩＮ码前８位
            Dim nfbh As String '年份编号
            Dim lsh_s As Integer '流水号开始
            Dim lsh_o As Integer '流水号结束
            Dim rwh As String = "" '任务号
            '首先判断改行是否为空
            If DataGridView1.Rows(i).Cells(0).Value.ToString <> "" And DataGridView1.Rows(i).Cells(1).Value.ToString <> "" And DataGridView1.Rows(i).Cells(2).Value.ToString <> "" And DataGridView1.Rows(i).Cells(3).Value.ToString <> "" Then
                nfbh = Mid(DataGridView1.Rows(i).Cells(2).Value.ToString, 1, 2) '取得年份编号“DG”等
                rwh = Trim(DataGridView1.Rows(i).Cells(3).Value.ToString)
                If Len(Trim(DataGridView1.Rows(i).Cells(2).Value.ToString)) = 17 Then

                    lsh_s = Val(Mid(DataGridView1.Rows(i).Cells(2).Value.ToString, 3, 6)) '取得流水号串中第一个流水号

                    lsh_o = Val(Mid(DataGridView1.Rows(i).Cells(2).Value.ToString, 12, 6)) '取得流水号串中最后一个流水号
                Else
                    lsh_s = Val(Mid(DataGridView1.Rows(i).Cells(2).Value.ToString, 3, 6)) '取得流水号串中第一个流水号
                    lsh_o = Val(Mid(DataGridView1.Rows(i).Cells(2).Value.ToString, 3, 6)) '取得流水号串中最后一个流水号
                End If





                cx = DataGridView1.Rows(i).Cells(0).Value.ToString & "/" & (lsh_o - lsh_s + 1).ToString & "辆" '取得车型
                VIN8 = DataGridView1.Rows(i).Cells(1).Value.ToString '取得ＶＩＮ码前８位
                Dim lsh(lsh_o - lsh_s + 1), lsh3(lsh_o - lsh_s + 1) As String
                Dim VIN18(lsh_o - lsh_s + 1), VIN183(lsh_o - lsh_s + 1) As String




                Dim cxx As New DataGridViewRow
                '定义车型、数量行
                cxx.CreateCells(DataGridView2)
                cxx.Cells(0).Value = ""
                cxx.Cells(1).Value = cx
                cxx.Cells(2).Value = ""
                cxx.Cells(3).Value = rwh
                DataGridView2.Rows.Add(cxx)




                For q = lsh_s To lsh_o
                    '定义流水号、VIN码列
                    Dim sj1 As New DataGridViewRow
                    Dim sj2 As New DataGridViewRow
                    sj1.CreateCells(DataGridView3)
                    sj2.CreateCells(DataGridView4)
                    lsh3(q - lsh_s) = nfbh + StrDup(8 - 2 - Len(q.ToString), "0") + q.ToString '算出每一个流水号
                    If VIN8 <> "" And lsh(q - lsh_s) <> "0" Then
                        VIN183(q - lsh_s) = VIN(VIN8, lsh3(q - lsh_s))  '算出每一个VIN码

                    End If
                    sj1.Cells(0).Value = lsh3(q - lsh_s)
                    sj1.Cells(1).Value = VIN183(q - lsh_s)
                    sj2.Cells(0).Value = VIN183(q - lsh_s)
                    sj2.Cells(1).Value = rwh
                    DataGridView3.Rows.Add(sj1)
                    DataGridView4.Rows.Add(sj2)
                Next

               
                '实验()
           



                '实验()

                    For p = lsh_s To lsh_o Step 2
                        '定义流水号、VIN码列
                        Dim sj As New DataGridViewRow
                        sj.CreateCells(DataGridView2)


                        lsh(p - lsh_s) = nfbh + StrDup(8 - 2 - Len(p.ToString), "0") + p.ToString '算出每一个流水号
                        If VIN8 <> "" And lsh(p - lsh_s) <> "0" Then
                            VIN18(p - lsh_s) = VIN(VIN8, lsh(p - lsh_s))  '算出每一个VIN码

                        End If

                        sj.Cells(0).Value = StrDup(8 - 2 - Len(p.ToString), "0") + p.ToString
                        sj.Cells(1).Value = VIN18(p - lsh_s)
                        If p <= lsh_o - 1 Then
                            lsh(p + 1 - lsh_s) = nfbh + StrDup(8 - 2 - Len((p + 1).ToString), "0") + (p + 1).ToString '算出每一个流水号
                            If VIN8 <> "" And lsh(p + 1 - lsh_s) <> "0" Then
                                VIN18(p + 1 - lsh_s) = VIN(VIN8, lsh(p + 1 - lsh_s)) '算出每一个VIN码
                            End If

                            sj.Cells(2).Value = StrDup(8 - 2 - Len((p + 1).ToString), "0") + (p + 1).ToString
                            sj.Cells(3).Value = VIN18(p + 1 - lsh_s)

                        Else
                            sj.Cells(2).Value = ""
                            sj.Cells(3).Value = ""

                        End If



                        DataGridView2.Rows.Add(sj)


                    Next


                
            End If

        Next
    End Sub

    Private Sub ToolStripButton8_Click(sender As Object, e As EventArgs) Handles ToolStripButton8.Click
        '导出普通上传格式
        If DataGridView2.Rows.Count <= 0 Then '判断记录数,如果没有记录就退出
            MessageBox.Show("没有记录可以导出", "没有可以导出的项目", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Else
            Dim saveExcel As SaveFileDialog
            saveExcel = New SaveFileDialog
            saveExcel.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            saveExcel.Filter = "Excel文件(.xlsx)|*.xlsx"
            saveExcel.FileName = "VIN导出"
            Dim filename As String
            If saveExcel.ShowDialog = System.Windows.Forms.DialogResult.Cancel Then Exit Sub

            filename = saveExcel.FileName
            Try
                daochu(DataGridView2, filename, 1)


                TLabel2.Text = "保存成功！位置：" + filename.ToString

                openbt.Visible = True

            Catch ex As Exception

            End Try
        End If
   

        'Dim i As Integer
        'Dim proc As Process()

        ''判断excel进程是否存在
        'If System.Diagnostics.Process.GetProcessesByName("excel").Length > 0 Then
        '    proc = Process.GetProcessesByName("excel")
        '    '得到名为excel进程个数，全部关闭
        '    For i = 0 To proc.Length - 1
        '        proc(i).Kill()
        '    Next
        'End If
        'proc = Nothing



    End Sub

    Private Sub ToolStripButton10_Click(sender As Object, e As EventArgs)
        AboutBox1.Show()

    End Sub

    Private Sub ToolStripButton11_Click(sender As Object, e As EventArgs)

    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs)







    End Sub

    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged









    End Sub

    Private Sub DataGridView1_CellStateChanged(sender As Object, e As DataGridViewCellStateChangedEventArgs) Handles DataGridView1.CellStateChanged

        Dim cx, qd, Vqd As String
        cx = ""

        qd = ""


        Try
            For i = 0 To DataGridView1.RowCount - 1
                DataGridView1.Rows(i).Cells(1).ErrorText = ""
                cx = Trim(DataGridView1.Rows(i).Cells(0).Value.ToString)
                Dim num As String = Regex.Replace(cx, "[\D]", "") '取出数字
                If Len(num) > 5 Then '是否大于5
                    If IsNumeric(Microsoft.VisualBasic.Mid(cx, 7, 1)) Then '如果时数字
                        qd = Microsoft.VisualBasic.Mid(num, 8, 1)
                    Else
                        qd = Microsoft.VisualBasic.Mid(num, 7, 1)

                    End If

                Else

                    '进行特殊车型判断
                    Select Case Microsoft.VisualBasic.Mid(cx, 1, 6)
                        Case "SX2110"
                            qd = "2"
                        Case "SX2180"
                            qd = "2"
                        Case "SX2150"
                            qd = "5"
                        Case "SX2160"
                            qd = "2"
                        Case "SX2151"
                            qd = "2"
                        Case "SX2153"
                            qd = "5"
                        Case "SX2180"
                            qd = "2"
                        Case "SX2190"
                            qd = "5"
                        Case "SX2300"
                            qd = "7"
                        Case "SX4260"
                            qd = "5"
                        Case "SX4323"
                            qd = "5"
                        Case "SX4400"
                            qd = "7"
                        Case "SX1380"
                            qd = "6"

                        Case "1291.2"
                            qd = "2"
                        Case "1491.2"
                            qd = "4"
                      

                    End Select

                End If
                If Len(Trim(DataGridView1.Rows(i).Cells(1).Value.ToString)) = 8 Then
                    Vqd = ""
                    Vqd = Microsoft.VisualBasic.Right(Trim(DataGridView1.Rows(i).Cells(1).Value.ToString), 1)
                    If qd <> Vqd Then
                        DataGridView1.Rows(i).Cells(1).ErrorText = Microsoft.VisualBasic.Constants.vbCrLf + "车型驱动位为" + qd + ",VIN码驱动位为" + Vqd
                    End If

                End If
            Next

        Catch ex As Exception

        End Try



    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs)



    End Sub

    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs)

        Dim fileName As String
        fileName = Me.OpenFileDialog1.FileName
        '建立EXCEL连接，读入数据
        Dim myDataset As New DataSet

        Dim strConn As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & fileName & "'; Extended Properties='Excel 8.0;HDR=YES;IMEX=2'"

        Dim da As New OleDb.OleDbDataAdapter("SELECT * FROM [Sheet1$A:C]  where len(trim(物料编码))>0", strConn)
        Try
            da.Fill(myDataset)

            Dim dr As DataRow = myDataset.Tables(0).NewRow
            For i = 0 To DataGridView1.RowCount - 1
                dr(i) = ""
            Next
            myDataset.Tables(0).Rows.Add(dr)

            Me.DataGridView1.DataSource = Nothing


            Me.DataGridView1.DataSource = myDataset.Tables(0)


            ToolStripButton4.Enabled = True  '可以计算
            '******载入后对列宽度进行重新规定*****'
            DataGridView1.Columns(0).Width = 140
            DataGridView1.Columns(1).Width = 103
            DataGridView1.Columns(2).Width = 165
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try





    End Sub

    Private Sub 插入IToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub datagridview1_cellerrortextneeded(sender As Object, e As DataGridViewCellErrorTextNeededEventArgs) Handles DataGridView1.CellErrorTextNeeded
        '错误判断()
        '1、车型前两位为SX
        '2、VIN码前3为必须为lzg；
        '3、第4位必须为c或者j
        '4、长度为8
        '5、VIN中不会包含 I、O、Q 三个英文字母*******重要
        '************************高级判断**************************************
        '6、如果第4位为C，第5位位2或3

        '8、驱动判断
        '  8.1、如果车型第三位不为2或5，取第11位与第8位比较、
        '  8.2、如果车型第三位为2，拿SX2190、2150、2151、2110、2153、2300、4323、4260、等进行判断
        '  8.3  如果车型第三位为5，第2个字幕组后3比较


        Dim dgv As DataGridView = CType(sender, DataGridView)
        Dim cellVal As Object = dgv(e.ColumnIndex, e.RowIndex).Value
        '1、车型前两位为SX
        If e.ColumnIndex = 0 And Microsoft.VisualBasic.Left(Trim(cellVal.ToString), 2) <> "SX" Then
            e.ErrorText += Microsoft.VisualBasic.Constants.vbCrLf + "车型前应该有" & """" & "SX" & """"
        End If
        '2、VIN码前3为必须为lzg；
        If e.ColumnIndex = 1 And Microsoft.VisualBasic.Left(Trim(cellVal.ToString), 3) <> "LZG" Then
            e.ErrorText += Microsoft.VisualBasic.Constants.vbCrLf + "VIN码前3位应为" & """" & "LZG" & """"
        End If
        '3、第4位必须为c或者j




        '4、长度为8
        If e.ColumnIndex = 1 And Microsoft.VisualBasic.Len(Trim(cellVal.ToString)) > 8 Then
            e.ErrorText += Microsoft.VisualBasic.Constants.vbCrLf + "VIN码大于8位，应为8位"
        End If
        If e.ColumnIndex = 1 And Microsoft.VisualBasic.Len(Trim(cellVal.ToString)) < 8 Then
            e.ErrorText += Microsoft.VisualBasic.Constants.vbCrLf + "VIN码小于8位，应为8位"
        End If
        '5、VIN中不会包含 I、O、Q 三个英文字母*******重要
        If e.ColumnIndex = 1 And InStr(Trim(cellVal.ToString), "I") Then
            e.ErrorText += Microsoft.VisualBasic.Constants.vbCrLf + "VIN中不会包含 I、O、Q 三个英文字母"
        End If
        If e.ColumnIndex = 1 And InStr(Trim(cellVal.ToString), "O") Then
            e.ErrorText += Microsoft.VisualBasic.Constants.vbCrLf + "VIN中不会包含 I、O、Q 三个英文字母"
        End If
        If e.ColumnIndex = 1 And InStr(Trim(cellVal.ToString), "Q") Then
            e.ErrorText += Microsoft.VisualBasic.Constants.vbCrLf + "VIN中不会包含 I、O、Q 三个英文字母"
        End If

        '6、如果第4位为C，第6位位2或3
        If e.ColumnIndex = 1 And Microsoft.VisualBasic.Mid(Trim(cellVal.ToString), 4, 1) = "C" Then
            If e.ColumnIndex = 1 And InStr("23", Microsoft.VisualBasic.Mid(Trim(cellVal.ToString), 6, 1)) Then

            Else
                e.ErrorText += Microsoft.VisualBasic.Constants.vbCrLf + "非完整车辆VIN码第6位应为" & """" & "2或3" & """"
            End If


        End If
        '7、任务号前3位为Z2F
        If e.ColumnIndex = 3 And Microsoft.VisualBasic.Mid(Trim(cellVal.ToString), 1, 3) <> "Z2F" Then
            
            e.ErrorText += Microsoft.VisualBasic.Constants.vbCrLf + cellVal.ToString + "开头非事业部任务号"



        End If
        '8、驱动判断 驱动的判读是一个粗略的判断，因为车型编码规则在陕汽执行不是很严格。——范建鹏 2013年8月29日
        '8.1先从车型中将数字全部提取出来，然后判断位数（小于5，则为特殊车型，不然）车型第7位是否为数字，如果是，取8，不然，取7


    End Sub

    Private Sub DataGridView1_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles DataGridView1.CellValidating
        '单元格错误显示
        Try


            'With Me.DataGridView1
            '    If e.ColumnIndex = 0 Then
            '        If Microsoft.VisualBasic.Left(Trim(e.FormattedValue), 2) <> "SX" Then
            '            Dim myerrot As String = "车型不正确"
            '                .Rows(e.RowIndex).ErrorText = myerrot

            '            e.Cancel = True


            '        End If


            '        End If






            'End With





        Catch ex As Exception
            MessageBox.Show(ex.Message, "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End Try


    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        With Me.DataGridView1
            If e.ColumnIndex = 0 Then
                .Rows(e.RowIndex).ErrorText = ""

            End If
        End With
    End Sub

    Private Sub ToolStripButton12_Click(sender As Object, e As EventArgs)
        Me.DataGridView1.Rows.Clear()



        Me.DataGridView1.Rows.Add(10)





    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        Try
            e.Value = e.Value.ToString.Trim
        Catch ex As Exception

        End Try



    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs)


    End Sub




    Private Sub 粘贴ToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub DataGridView1_CellMouseUp(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseUp

    End Sub

    Private Sub DataGridView1_RowStateChanged(sender As Object, e As DataGridViewRowStateChangedEventArgs) Handles DataGridView1.RowStateChanged
        LabelState.Text = "共" + DataGridView1.RowCount.ToString + "行"
    End Sub

    Private Sub DataGridView1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        Try
            TLabel2.Text = "（第" + (DataGridView1.CurrentCell.RowIndex + 1).ToString + "行,第" + (DataGridView1.CurrentCell.ColumnIndex + 1).ToString + "列)" + DataGridView1.CurrentCell.ErrorText
            TLabel2.Text = TLabel2.Text.Replace(Microsoft.VisualBasic.Constants.vbCrLf, "@")
        Catch ex As Exception

        End Try

    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub DataGridView1_CellContextMenuStripNeeded(sender As Object, e As DataGridViewCellContextMenuStripNeededEventArgs) Handles DataGridView1.CellContextMenuStripNeeded

    End Sub

    Private Sub ToolStripButton7_Click(sender As Object, e As EventArgs) Handles ToolStripButton7.Click
        '导出MES格式
        If DataGridView3.Rows.Count <= 0 Then '判断记录数,如果没有记录就退出
            MessageBox.Show("没有记录可以导出", "没有可以导出的项目", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            Dim saveExcel As SaveFileDialog
            saveExcel = New SaveFileDialog
            saveExcel.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            saveExcel.Filter = "Excel文件(.xlsx)|*.xlsx"
            saveExcel.FileName = "VIN(MES)导出"
            Dim filename As String
            If saveExcel.ShowDialog = System.Windows.Forms.DialogResult.Cancel Then Exit Sub

            filename = saveExcel.FileName
            Try
                daochu(DataGridView3, filename, 1)
                TLabel2.Text = "保存成功！位置：" + filename.ToString
                openbt.Visible = True
            Catch ex As Exception

            End Try

        End If
      

        'Dim i As Integer
        'Dim proc As Process()

        ''判断excel进程是否存在
        'If System.Diagnostics.Process.GetProcessesByName("excel").Length > 0 Then
        '    proc = Process.GetProcessesByName("excel")
        '    '得到名为excel进程个数，全部关闭
        '    For i = 0 To proc.Length - 1
        '        proc(i).Kill()
        '    Next
        'End If
        'proc = Nothing

      

    End Sub

    Private Sub Label_LZG_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label_LZG_MouseHover(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button_C_Click(sender As Object, e As EventArgs)

        ' Create an instance of the ListBox.
        Dim listBox1 As New System.Windows.Forms.ListBox
        ' Set the size and location of the ListBox.
        listBox1.Size = New System.Drawing.Size(200, 100)
        listBox1.Location = New System.Drawing.Point(10, 10)
        ' Add the ListBox to the form.
        Me.TabControl1.TabPages(1).Controls.Add(listBox1)
        Me.TabControl1.TabPages(0).Controls.Add(listBox1)
        Me.Controls.Add(listBox1)
        ' Set the ListBox to display items in multiple columns.
        listBox1.MultiColumn = True
        ' Set the selection mode to multiple and extended.
        listBox1.SelectionMode = SelectionMode.MultiExtended

        ' Shutdown the painting of the ListBox as items are added.
        listBox1.BeginUpdate()
        ' Loop through and add 50 items to the ListBox.
        Dim x As Integer
        For x = 1 To 50
            listBox1.Items.Add("Item " & x.ToString())
        Next x
        ' Allow the ListBox to repaint and display the new items.
        listBox1.EndUpdate()

        ' Select three items from the ListBox.
        listBox1.SetSelected(1, True)
        listBox1.SetSelected(3, True)
        listBox1.SetSelected(5, True)

        ' Display the second selected item in the ListBox to the console.
        System.Diagnostics.Debug.WriteLine(listBox1.SelectedItems(1).ToString())
        ' Display the index of the first selected item in the ListBox.
        System.Diagnostics.Debug.WriteLine(listBox1.SelectedIndices(0).ToString())
    End Sub

    Private Sub ListBox_4_SelectedIndexChanged(sender As Object, e As EventArgs)


    End Sub

    Private Sub Button_5_Click(sender As Object, e As EventArgs)


    End Sub


    Private Sub ToolStripComboBox1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ToolStripComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs)



    End Sub

    Private Sub ToolStripComboBox1_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub ToolStripButton5_Click_1(sender As Object, e As EventArgs)

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged_1(sender As Object, e As EventArgs)


    End Sub

    Private Sub ToolStripDropDownButton1_Click(sender As Object, e As EventArgs) Handles ToolStripDropDownButton3.Click

    End Sub

    Private Sub ToolStripDropDownButton1_DropDownItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles ToolStripDropDownButton3.DropDownItemClicked


        ToolStripDropDownButton3.Text = Microsoft.VisualBasic.Left(Trim(ToolStripDropDownButton3.DropDownItems(ToolStripDropDownButton3.DropDownItems.IndexOf(e.ClickedItem)).Text), 1)






    End Sub

    Private Sub ToolStripDropDownButton2_Click(sender As Object, e As EventArgs) Handles ToolStripDropDownButton2.Click

    End Sub

    Private Sub ToolStripDropDownButton2_DropDownItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles ToolStripDropDownButton2.DropDownItemClicked
        ToolStripDropDownButton2.Text = Microsoft.VisualBasic.Left(Trim(ToolStripDropDownButton2.DropDownItems(ToolStripDropDownButton2.DropDownItems.IndexOf(e.ClickedItem)).Text), 1)
    End Sub

    Private Sub ToolStripDropDownButton4_Click(sender As Object, e As EventArgs) Handles ToolStripDropDownButton4.Click

    End Sub

    Private Sub ToolStripDropDownButton4_DropDownItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles ToolStripDropDownButton4.DropDownItemClicked
        ToolStripDropDownButton4.Text = Microsoft.VisualBasic.Left(Trim(ToolStripDropDownButton4.DropDownItems(ToolStripDropDownButton4.DropDownItems.IndexOf(e.ClickedItem)).Text), 1)
    End Sub

    Private Sub ToolStripDropDownButton5_Click(sender As Object, e As EventArgs) Handles ToolStripDropDownButton5.Click

    End Sub

    Private Sub ToolStripDropDownButton5_DropDownItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles ToolStripDropDownButton5.DropDownItemClicked
        ToolStripDropDownButton5.Text = Microsoft.VisualBasic.Left(Trim(ToolStripDropDownButton5.DropDownItems(ToolStripDropDownButton5.DropDownItems.IndexOf(e.ClickedItem)).Text), 1)
    End Sub

    Private Sub ToolStripDropDownButton6_Click(sender As Object, e As EventArgs) Handles ToolStripDropDownButton6.Click

    End Sub

    Private Sub ToolStripDropDownButton6_DropDownItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles ToolStripDropDownButton6.DropDownItemClicked
        ToolStripDropDownButton6.Text = Microsoft.VisualBasic.Left(Trim(ToolStripDropDownButton6.DropDownItems(ToolStripDropDownButton6.DropDownItems.IndexOf(e.ClickedItem)).Text), 1)
    End Sub

    Private Sub ToolStripButton14_Click(sender As Object, e As EventArgs) Handles ToolStripButton14.Click
        Dim q8, h8, x9 As String
        h8 = ""

        If ToolStripTextBox1.Text = "" Then
            MsgBox("请输入流水号")

        Else
            If Len(Trim(ToolStripTextBox1.Text)) <> 8 Then
                MsgBox("流水号应该为8位")

            Else
                h8 = Trim(ToolStripTextBox1.Text)
                q8 = ToolStripButton5.Text + ToolStripDropDownButton2.Text + ToolStripDropDownButton3.Text + ToolStripDropDownButton4.Text + ToolStripDropDownButton5.Text + ToolStripDropDownButton6.Text
                x9 = Microsoft.VisualBasic.Mid(Trim(VIN(q8, h8)), 9, 1)
                ToolStripButton12.Text = x9
                ToolStripLabel3.Text = "*" + Trim(VIN(q8, h8)) + "*"

                ToolStripLabel2.Text = "计算结果："
                ToolStripButton5.Visible = False
                ToolStripDropDownButton2.Visible = False
                ToolStripDropDownButton3.Visible = False
                ToolStripDropDownButton4.Visible = False
                ToolStripDropDownButton5.Visible = False
                ToolStripDropDownButton6.Visible = False
                ToolStripButton12.Visible = False
                ToolStripTextBox1.Visible = False

                ToolStripLabel3.Visible = True
                ToolStripButton14.Visible = False
                ToolStripButton18.Visible = True '返回按钮
            End If
        End If




    End Sub

    Private Sub TabPage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Click



    End Sub

    Private Sub ToolStripTextBox1_Click(sender As Object, e As EventArgs) Handles ToolStripTextBox1.Click

    End Sub

    Private Sub ToolStripTextBox1_TextChanged(sender As Object, e As EventArgs) Handles ToolStripTextBox1.TextChanged
        ToolStripTextBox1.Text = ToolStripTextBox1.Text.ToUpper

    End Sub

    Private Sub ToolStripButton17_Click(sender As Object, e As EventArgs) Handles ToolStripButton17.Click
        AboutBox1.Show()
    End Sub

    Private Sub ToolStripLabel3_Click(sender As Object, e As EventArgs) Handles ToolStripLabel3.Click

    End Sub

    Private Sub ToolStripButton18_Click(sender As Object, e As EventArgs) Handles ToolStripButton18.Click
        ToolStripLabel2.Text = "代码选择："
        ToolStripButton5.Visible = True
        ToolStripDropDownButton2.Visible = True
        ToolStripDropDownButton3.Visible = True
        ToolStripDropDownButton4.Visible = True
        ToolStripDropDownButton5.Visible = True
        ToolStripDropDownButton6.Visible = True
        ToolStripButton12.Visible = True
        ToolStripTextBox1.Visible = True

        ToolStripLabel3.Visible = False
        ToolStripButton14.Visible = True
        ToolStripButton18.Visible = False

        ToolStripButton12.Text = "?" '检验位
        ToolStripTextBox1.Text = "" '流水号归0



    End Sub

    Private Sub ToolStripButton20_Click(sender As Object, e As EventArgs) Handles ToolStripButton20.Click
        Try
            If ToolStripTextBox2.Text <> "" Then
                If Microsoft.VisualBasic.Left(Trim(ToolStripTextBox2.Text), 7) <> "http://" And Microsoft.VisualBasic.Left(Trim(ToolStripTextBox2.Text), 7) <> "file://" Then
                    WebBrowser1.Url = New Uri("http://" + ToolStripTextBox2.Text)
                Else
                    WebBrowser1.Url = New Uri(ToolStripTextBox2.Text)

                End If
            End If
        Catch ex As Exception

        End Try



    End Sub

    Private Sub Form1_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        ToolStripTextBox2.Width = WebBrowser1.Width - ToolStripButton20.Width - 20

    End Sub

    Private Sub ListView2_ItemChecked(sender As Object, e As ItemCheckedEventArgs)

    End Sub

    Private Sub ListView2_SelectedIndexChanged(sender As Object, e As EventArgs)



    End Sub

    Private Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView1.AfterSelect

    End Sub

    Private Sub TreeView1_MouseClick(sender As Object, e As MouseEventArgs) Handles TreeView1.MouseClick

    End Sub

    Private Sub TreeView1_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseClick
        Try

            Dim startPath As String
            startPath = "file:///" + AppDomain.CurrentDomain.BaseDirectory.ToString.Replace("\", "/")

            Select Case e.Node.Name.ToString
                Case "节点0"
                    ToolStripTextBox2.Text = "首页.html"

                Case "节点1"
                    ToolStripTextBox2.Text = "1.装配计划的查询，导出.htm"
                Case "节点2"
                    ToolStripTextBox2.Text = "2、VIN码及技术参数的编制.html"
                Case "节点3"
                    ToolStripTextBox2.Text = "3、VIN及技术参数发放流程.htm"

                Case "节点4"
                    ToolStripTextBox2.Text = "4、更改单发放流程.htm"
                Case "节点5"
                    ToolStripTextBox2.Text = "5、注意事项.htm"

            End Select
            Dim urll As String
            urll = startPath + "res/shouce/" + ToolStripTextBox2.Text


            Try
                If ToolStripTextBox2.Text <> "" Then
                    If Microsoft.VisualBasic.Left(Trim(urll), 7) <> "http://" And Microsoft.VisualBasic.Left(Trim(urll), 7) <> "file://" Then
                        WebBrowser1.Url = New Uri("http://" + urll)
                    Else
                        WebBrowser1.Url = New Uri(urll)

                    End If
                End If
            Catch ex As Exception

            End Try
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ToolStripTextBox2_Click(sender As Object, e As EventArgs) Handles ToolStripTextBox2.Click

    End Sub

    Private Sub ToolStripTextBox2_KeyUp(sender As Object, e As KeyEventArgs) Handles ToolStripTextBox2.KeyUp
        If e.KeyCode = Keys.Enter Then
            Try
                If ToolStripTextBox2.Text <> "" Then
                    If Microsoft.VisualBasic.Left(Trim(ToolStripTextBox2.Text), 7) <> "http://" And Microsoft.VisualBasic.Left(Trim(ToolStripTextBox2.Text), 7) <> "file://" Then
                        WebBrowser1.Url = New Uri("http://" + ToolStripTextBox2.Text)
                    Else
                        WebBrowser1.Url = New Uri(ToolStripTextBox2.Text)

                    End If
                End If
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub ToolStripButton15_Click(sender As Object, e As EventArgs) Handles ToolStripButton15.Click
        If ToolStripButton15.Text = "隐藏" Then
            ToolStripButton15.Text = "显示"
            SplitContainer3.SplitterDistance = "50"

            Panel2.Dock = DockStyle.Fill
            ToolStrip2.Dock = DockStyle.Left


            '隐藏按钮和treeview
            'ToolStripButton13.Visible = False
            'ToolStripButton11.Visible = False
            'ToolStripButton10.Visible = False
            'ToolStripButton3.Visible = False
            TreeView1.Visible = False

        Else
            ToolStripButton15.Text = "隐藏"
            SplitContainer3.SplitterDistance = "262"
            'ToolStripButton13.Visible = True
            'ToolStripButton11.Visible = True
            'ToolStripButton10.Visible = True
            'ToolStripButton3.Visible = True
            Panel2.Dock = DockStyle.Top
            Panel2.Height = 41

            ToolStrip2.Dock = DockStyle.Top
            TreeView1.Visible = True
        End If
    End Sub

    Private Sub ToolStripButton13_Click(sender As Object, e As EventArgs) Handles ToolStripButton13.Click
        WebBrowser1.GoBack()



    End Sub

    Private Sub ToolStripButton11_Click_1(sender As Object, e As EventArgs) Handles ToolStripButton11.Click
        WebBrowser1.GoForward()
    End Sub

    Private Sub ToolStripButton10_Click_1(sender As Object, e As EventArgs) Handles ToolStripButton10.Click
        WebBrowser1.Url = New Uri("file:///" + AppDomain.CurrentDomain.BaseDirectory.ToString.Replace("\", "/") + "res/shouce/首页.html")
    End Sub

    Private Sub 清除内容NToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub openbt_Click(sender As Object, e As EventArgs) Handles openbt.Click
        Dim lujing As String
        lujing = TLabel2.Text.Replace("保存成功！位置：", "")

        Shell("explorer.exe /select," & Chr(34) & lujing & Chr(34), 1)
    End Sub

    Private Sub ToolStripButton2_Click_1(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        '导出MES格式
        If DataGridView4.Rows.Count <= 0 Then '判断记录数,如果没有记录就退出
            MessageBox.Show("没有记录可以导出", "没有可以导出的项目", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            Dim saveExcel As SaveFileDialog
            saveExcel = New SaveFileDialog
            saveExcel.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            saveExcel.Filter = "Excel文件(.xlsx)|*.xlsx"
            saveExcel.FileName = "VIN(打印)导出"
            Dim filename As String
            If saveExcel.ShowDialog = System.Windows.Forms.DialogResult.Cancel Then Exit Sub

            filename = saveExcel.FileName
            Try
                daochu(DataGridView4, filename, 1)
                TLabel2.Text = "保存成功！位置：" + filename.ToString
                openbt.Visible = True
            Catch ex As Exception

            End Try

        End If
    End Sub
End Class
