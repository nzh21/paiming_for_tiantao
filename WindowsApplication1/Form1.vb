Public Class Form1
    Dim xlApp As Excel.Application '定义EXCEL类
    Dim xlBook As Excel.Workbook '定义工件簿类
    Dim xlsheet As Excel.Worksheet '定义工作表类
    Private Property Fault As Boolean
    Private Property ture As Boolean


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click '开始按钮
        Button1.Enabled = Fault
        Me.Cursor = Cursors.AppStarting
        BackgroundWorker1.RunWorkerAsync()  '起始背景執行的呼叫

    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click '结束按钮
        If BackgroundWorker1.IsBusy <> True Then
            Me.Close()
        ElseIf MsgBox("是否停止？", MsgBoxStyle.YesNo, "正在运行") = MsgBoxResult.Yes Then
            BackgroundWorker1.CancelAsync()
        End If

    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click '浏览...按钮
        With OpenFileDialog1
            .Filter = "Microsoft Office Excel 工作簿 (*.xls)|*.xls|所有文件 (*.*)|*.*"
            .FilterIndex = 1
            .Title = "打开文件"
        End With
        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            TextBox1.Text = OpenFileDialog1.FileName
        End If
    End Sub
    Private Sub Form_Initialize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.MaximumSize = Me.Size
        Me.MinimumSize = Me.Size
        ComboBox1.Text = "8"
        Tdj9.Text = "A+"
        Tdj8.Text = "A"
        Tdj7.Text = "B+"
        Tdj6.Text = "B"
        Tdj5.Text = "C+"
        Tdj4.Text = "C"
        Tdj3.Text = "D"
        Tdj2.Text = "E"
        CheckBox0.Checked = True
        CheckBox1.Checked = True
        CheckBox2.Checked = True
        CheckBox3.Checked = True
        CheckBox4.Checked = True
        CheckBox5.Checked = True
        CheckBox6.Checked = True
    End Sub
    Private Sub CheckBox0_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox0.CheckedChanged
        If CheckBox0.Checked = False Then Tkm0.Text = "" : Tkm0.Enabled = False '清空并禁用tkm0
        If CheckBox0.Checked = True Then Tkm0.Enabled = True '启用tkm0
    End Sub
    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = False Then Tkm1.Text = "" : Tkm1.Enabled = False
        If CheckBox1.Checked = True Then Tkm1.Enabled = True
    End Sub
    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = False Then Tkm2.Text = "" : Tkm2.Enabled = False
        If CheckBox2.Checked = True Then Tkm2.Enabled = True
    End Sub
    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = False Then Tkm3.Text = "" : Tkm3.Enabled = False
        If CheckBox3.Checked = True Then Tkm3.Enabled = True
    End Sub
    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = False Then Tkm4.Text = "" : Tkm4.Enabled = False
        If CheckBox4.Checked = True Then Tkm4.Enabled = True
    End Sub
    Private Sub CheckBox5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = False Then Tkm5.Text = "" : Tkm5.Enabled = False
        If CheckBox5.Checked = True Then Tkm5.Enabled = True
    End Sub
    Private Sub CheckBox6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox6.CheckedChanged
        If CheckBox6.Checked = False Then Tkm6.Text = "" : Tkm6.Enabled = False
        If CheckBox6.Checked = True Then Tkm6.Enabled = True
    End Sub
    Private Sub CheckBox7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox7.CheckedChanged
        If CheckBox7.Checked = False Then Tkm7.Text = "" : Tkm7.Enabled = False
        If CheckBox7.Checked = True Then Tkm7.Enabled = True
    End Sub
    Private Sub CheckBox8_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox8.CheckedChanged
        If CheckBox8.Checked = False Then Tkm8.Text = "" : Tkm8.Enabled = False
        If CheckBox8.Checked = True Then Tkm8.Enabled = True
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Tdj0.Text = ""
        Tdj1.Text = ""
        Tdj2.Text = ""
        Tdj3.Text = ""
        Tdj4.Text = ""
        Tdj5.Text = ""
        Tdj6.Text = ""
        Tdj7.Text = ""
        Tdj8.Text = ""
        Tdj9.Text = ""
        If ComboBox1.Text = "1" Then
            Tdj9.Enabled = True
            Tdj8.Enabled = False
            Tdj7.Enabled = False
            Tdj6.Enabled = False
            Tdj5.Enabled = False
            Tdj4.Enabled = False
            Tdj3.Enabled = False
            Tdj2.Enabled = False
            Tdj1.Enabled = False
            Tdj0.Enabled = False
        End If
        If ComboBox1.Text = "2" Then
            Tdj9.Enabled = True
            Tdj8.Enabled = True
            Tdj7.Enabled = False
            Tdj6.Enabled = False
            Tdj5.Enabled = False
            Tdj4.Enabled = False
            Tdj3.Enabled = False
            Tdj2.Enabled = False
            Tdj1.Enabled = False
            Tdj0.Enabled = False
        End If
        If ComboBox1.Text = "3" Then
            Tdj9.Enabled = True
            Tdj8.Enabled = True
            Tdj7.Enabled = True
            Tdj6.Enabled = False
            Tdj5.Enabled = False
            Tdj4.Enabled = False
            Tdj3.Enabled = False
            Tdj2.Enabled = False
            Tdj1.Enabled = False
            Tdj0.Enabled = False
        End If
        If ComboBox1.Text = "4" Then
            Tdj9.Enabled = True
            Tdj8.Enabled = True
            Tdj7.Enabled = True
            Tdj6.Enabled = True
            Tdj5.Enabled = False
            Tdj4.Enabled = False
            Tdj3.Enabled = False
            Tdj2.Enabled = False
            Tdj1.Enabled = False
            Tdj0.Enabled = False
        End If
        If ComboBox1.Text = "5" Then
            Tdj9.Enabled = True
            Tdj8.Enabled = True
            Tdj7.Enabled = True
            Tdj6.Enabled = True
            Tdj5.Enabled = True
            Tdj4.Enabled = False
            Tdj3.Enabled = False
            Tdj2.Enabled = False
            Tdj1.Enabled = False
            Tdj0.Enabled = False
        End If
        If ComboBox1.Text = "6" Then
            Tdj9.Enabled = True
            Tdj8.Enabled = True
            Tdj7.Enabled = True
            Tdj6.Enabled = True
            Tdj5.Enabled = True
            Tdj4.Enabled = True
            Tdj3.Enabled = False
            Tdj2.Enabled = False
            Tdj1.Enabled = False
            Tdj0.Enabled = False
        End If
        If ComboBox1.Text = "7" Then
            Tdj9.Enabled = True
            Tdj8.Enabled = True
            Tdj7.Enabled = True
            Tdj6.Enabled = True
            Tdj5.Enabled = True
            Tdj4.Enabled = True
            Tdj3.Enabled = True
            Tdj2.Enabled = False
            Tdj1.Enabled = False
            Tdj0.Enabled = False
        End If
        If ComboBox1.Text = "8" Then
            Tdj9.Enabled = True
            Tdj8.Enabled = True
            Tdj7.Enabled = True
            Tdj6.Enabled = True
            Tdj5.Enabled = True
            Tdj4.Enabled = True
            Tdj3.Enabled = True
            Tdj2.Enabled = True
            Tdj1.Enabled = False
            Tdj0.Enabled = False
        End If
        If ComboBox1.Text = "9" Then
            Tdj9.Enabled = True
            Tdj8.Enabled = True
            Tdj7.Enabled = True
            Tdj6.Enabled = True
            Tdj5.Enabled = True
            Tdj4.Enabled = True
            Tdj3.Enabled = True
            Tdj2.Enabled = True
            Tdj1.Enabled = True
            Tdj0.Enabled = False
        End If
        If ComboBox1.Text = "10" Then
            Tdj9.Enabled = True
            Tdj8.Enabled = True
            Tdj7.Enabled = True
            Tdj6.Enabled = True
            Tdj5.Enabled = True
            Tdj4.Enabled = True
            Tdj3.Enabled = True
            Tdj2.Enabled = True
            Tdj1.Enabled = True
            Tdj0.Enabled = True
        End If
    End Sub
    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
5:
        BackgroundWorker1.ReportProgress(0)
        Dim g As Integer 'g为科目总数
        If Tkm0.Text <> "" Then
            g = g + 1
        End If
        BackgroundWorker1.ReportProgress(50)
        If Tkm1.Text <> "" Then
            g = g + 1
            If g < 2 Then CheckBox0.Checked = True : Tkm0.Text = Tkm1.Text : CheckBox1.Checked = False : Tkm1.Text = "" : Tkm1.Enabled = False : Tkm0.Enabled = True
            '上一行代码意思是：若Tkm0无字符，则将Tkm1中的文字剪切指Tkm0
        End If
        BackgroundWorker1.ReportProgress(100) '汇报进度
        If Tkm2.Text <> "" Then
            g = g + 1
            If g < 3 Then CheckBox1.Checked = True : Tkm1.Text = Tkm2.Text : CheckBox2.Checked = False : Tkm2.Text = "" : Tkm2.Enabled = False : Tkm1.Enabled = True
        End If
        BackgroundWorker1.ReportProgress(150) '汇报进度
        If Tkm3.Text <> "" Then
            g = g + 1
            If g < 4 Then CheckBox2.Checked = True : Tkm2.Text = Tkm3.Text : CheckBox3.Checked = False : Tkm3.Text = "" : Tkm3.Enabled = False : Tkm2.Enabled = True
        End If
        BackgroundWorker1.ReportProgress(200)
        If Tkm4.Text <> "" Then
            g = g + 1
            If g < 5 Then CheckBox3.Checked = True : Tkm3.Text = Tkm4.Text : CheckBox4.Checked = False : Tkm4.Text = "" : Tkm4.Enabled = False : Tkm3.Enabled = True
        End If
        BackgroundWorker1.ReportProgress(250)
        If Tkm5.Text <> "" Then
            g = g + 1
            If g < 6 Then CheckBox4.Checked = True : Tkm4.Text = Tkm5.Text : CheckBox5.Checked = False : Tkm5.Text = "" : Tkm5.Enabled = False : Tkm4.Enabled = True
        End If
        BackgroundWorker1.ReportProgress(300)
        If Tkm6.Text <> "" Then
            g = g + 1
            If g < 7 Then CheckBox5.Checked = True : Tkm5.Text = Tkm6.Text : CheckBox6.Checked = False : Tkm6.Text = "" : Tkm6.Enabled = False : Tkm5.Enabled = True
        End If
        BackgroundWorker1.ReportProgress(350)
        If Tkm7.Text <> "" Then
            g = g + 1
            If g < 8 Then CheckBox6.Checked = True : Tkm6.Text = Tkm7.Text : CheckBox7.Checked = False : Tkm7.Text = "" : Tkm7.Enabled = False : Tkm6.Enabled = True
        End If
        BackgroundWorker1.ReportProgress(400)
        If Tkm8.Text <> "" Then
            g = g + 1
            If g < 9 Then CheckBox7.Checked = True : Tkm7.Text = Tkm8.Text : CheckBox8.Checked = False : Tkm8.Text = "" : Tkm8.Enabled = False : Tkm7.Enabled = True
        End If
        BackgroundWorker1.ReportProgress(450)
        Dim p As Integer, n As Integer, i As Integer, q As Integer, f As Integer
        f = Val(TextBox2.Text)
        xlApp = CreateObject("Excel.Application") '创建EXCEL应用类
        xlApp.Visible = True '设置EXCEL可见
        xlBook = xlApp.Workbooks.Open(TextBox1.Text) '打开EXCEL工作簿
        xlsheet = xlBook.Worksheets(f) '打开EXCEL工作表
        '打开文件
        BackgroundWorker1.ReportProgress(600)
        p = 0
        Dim kmy(8) As Integer
        Do
            p = p + 1
            If xlsheet.Cells(1, p).Value = Tkm1.Text Then kmy(1) = p
            If xlsheet.Cells(1, p).Value = Tkm2.Text Then kmy(2) = p
            If xlsheet.Cells(1, p).Value = Tkm3.Text Then kmy(3) = p
            If xlsheet.Cells(1, p).Value = Tkm4.Text Then kmy(4) = p
            If xlsheet.Cells(1, p).Value = Tkm5.Text Then kmy(5) = p
            If xlsheet.Cells(1, p).Value = Tkm6.Text Then kmy(6) = p
            If xlsheet.Cells(1, p).Value = Tkm7.Text Then kmy(7) = p
            If xlsheet.Cells(1, p).Value = Tkm8.Text Then kmy(8) = p
            If xlsheet.Cells(1, p).Value = Tkm0.Text Then kmy(0) = p
            '获取各科所在列，kmy（0）为总分
        Loop While xlsheet.Cells(1, p).Value <> Nothing
        BackgroundWorker1.ReportProgress(700)
        Dim h As Integer
        If kmy(1) = 0 Then
            If CheckBox1.Checked = True Then h = MsgBox(("未找到" & Tkm1.Text + Chr(13) & Chr(10) & ("如点击""忽略""，将忽略同类错误")), 2 + 16 + 65536, ("错误")) : GoTo 80
        End If

        If kmy(2) = 0 Then
            If CheckBox2.Checked = True Then h = MsgBox(("未找到" & Tkm2.Text + Chr(13) & Chr(10) & ("如点击""忽略""，将忽略同类错误")), 2 + 16 + 65536, ("错误")) : GoTo 80
        End If
        If kmy(3) = 0 Then
            If CheckBox3.Checked = True Then h = MsgBox(("未找到" & Tkm3.Text + Chr(13) & Chr(10) & ("如点击""忽略""，将忽略同类错误")), 2 + 16 + 65536, ("错误")) : GoTo 80
        End If

        If kmy(4) = 0 Then
            If CheckBox4.Checked = True Then h = MsgBox(("未找到" & Tkm4.Text + Chr(13) & Chr(10) & ("如点击""忽略""，将忽略同类错误")), 2 + 16 + 65536, ("错误")) : GoTo 80
        End If

        If kmy(5) = 0 Then
            If CheckBox5.Checked = True Then h = MsgBox(("未找到" & Tkm5.Text + Chr(13) & Chr(10) & ("如点击""忽略""，将忽略同类错误")), 2 + 16 + 65536, ("错误")) : GoTo 80
        End If

        If kmy(6) = 0 Then
            If CheckBox6.Checked = True Then h = MsgBox(("未找到" & Tkm6.Text + Chr(13) & Chr(10) & ("如点击""忽略""，将忽略同类错误")), 2 + 16 + 65536, ("错误")) : GoTo 80
        End If

        If kmy(7) = 0 Then
            If CheckBox7.Checked = True Then h = MsgBox(("未找到" & Tkm7.Text + Chr(13) & Chr(10) & ("如点击""忽略""，将忽略同类错误")), 2 + 16 + 65536, ("错误")) : GoTo 80
        End If

        If kmy(8) = 0 Then
            If CheckBox8.Checked = True Then h = MsgBox(("未找到" & Tkm8.Text + Chr(13) & Chr(10) & ("如点击""忽略""，将忽略同类错误")), 2 + 16 + 65536, ("错误")) : GoTo 80
        End If

        If kmy(0) = 0 Then
            If CheckBox0.Checked = True Then h = MsgBox(("未找到" & Tkm0.Text + Chr(13) & Chr(10) & ("如点击""忽略""，将忽略同类错误")), 1 + 16 + 65536, ("错误")) : GoTo 80
        End If
        '检查是否有科目未找到
70:
        BackgroundWorker1.ReportProgress(850)

        q = 0
        Do
            q = q + 1
        Loop While xlsheet.Cells(q, 1).value <> Nothing '人数为q-2 最后一人所在行为q-1
        BackgroundWorker1.ReportProgress(950)

        Dim km(8, q) As String
        Dim kmd(8, q) As String

        For n = 2 To q - 1 Step 1
            For p = 0 To g
                km(p, n - 1) = xlsheet.Cells(n, kmy(p)).Value
                '获取各科成绩(包括分数与等级）
            Next
            BackgroundWorker1.ReportProgress((10 * n / (20 * q)) * 9000 + 1000)
        Next
        For n = 2 To q - 1 Step 1
            For p = 0 To g
                km(p, n - 1) = Replace(km(p, n - 1), "(", ",(")
                '用，（代替（ 方便下面分割
            Next
            BackgroundWorker1.ReportProgress(((10 * q + 0.5 * n) / (20 * q)) * 9000 + 1000)
        Next
        Dim b() As String
        For n = 1 To q Step 1
            For p = 0 To g
                b = Split(km(p, n - 1), ",", 2)
                km(p, n) = b(UBound(b, 1)) '分割分数与等级，等级在b(UBound(b, 1)) km(p, n)表示第n个人第p科成绩(等级）
                BackgroundWorker1.ReportProgress(((10.5 * q + 0.5 * n) / (20 * q)) * 9000 + 1000)
            Next
        Next

        For n = 2 To q - 1
            For p = 0 To g
                If km(p, n - 1) = "(" & Tdj9.Text & ")" Or km(p, n - 1) = Tdj9.Text Then
                    kmd(p, n - 1) = 9
                ElseIf km(p, n - 1) = "(" & Tdj8.Text & ")" Or km(p, n - 1) = Tdj8.Text Then
                    kmd(p, n - 1) = 8
                ElseIf km(p, n - 1) = "(" & Tdj7.Text & ")" Or km(p, n - 1) = Tdj7.Text Then
                    kmd(p, n - 1) = 7
                ElseIf km(p, n - 1) = "(" & Tdj6.Text & ")" Or km(p, n - 1) = Tdj6.Text Then
                    kmd(p, n - 1) = 6
                ElseIf km(p, n - 1) = "(" & Tdj5.Text & ")" Or km(p, n - 1) = Tdj5.Text Then
                    kmd(p, n - 1) = 5
                ElseIf km(p, n - 1) = "(" & Tdj4.Text & ")" Or km(p, n - 1) = Tdj4.Text Then
                    kmd(p, n - 1) = 4
                ElseIf km(p, n - 1) = "(" & Tdj3.Text & ")" Or km(p, n - 1) = Tdj3.Text Then
                    kmd(p, n - 1) = 3
                ElseIf km(p, n - 1) = "(" & Tdj2.Text & ")" Or km(p, n - 1) = Tdj2.Text Then
                    kmd(p, n - 1) = 2
                ElseIf km(p, n - 1) = "(" & Tdj1.Text & ")" Or km(p, n - 1) = Tdj1.Text Then
                    kmd(p, n - 1) = 1
                ElseIf km(p, n - 1) = "(" & Tdj0.Text & ")" Or km(p, n - 1) = Tdj0.Text Then
                    kmd(p, n - 1) = 0
                End If
                '将等级转化为数值，数值用kmd（）储存
            Next
            BackgroundWorker1.ReportProgress(((12 * q + n) / (20 * q)) * 9000 + 1000)
        Next

        i = 1
        Do While xlsheet.Cells(1, i).Value <> Nothing
            i = i + 1
        Loop '检查表格共有几列


        Dim d(9, q - 1) As Integer
        Dim zfdd(q - 1) As Integer

        For n = 1 To q - 1
            For p = 1 To g
                d(kmd(p, n), n) = d(kmd(p, n), n) + 1 '统计第n个人的kmd(p, n)的数量
            Next
            zfdd(n) = kmd(0, n) '统计第n个人的总分等级
            BackgroundWorker1.ReportProgress(((13 * q + n) / (20 * q)) * 9000 + 1000)
        Next

        Dim w(q - 1) As String
        Dim o(q - 1) As String
        Dim r(q - 1) As String
        For n = 1 To q - 1
            w(n) = Format(zfdd(n), 0) & Format(d(9, n), 0) & Format(d(8, n), 0) & Format(d(7, n), 0) & Format(d(6, n), 0) & Format(d(5, n), 0) & Format(d(4, n), 0) & Format(d(3, n), 0) & Format(d(2, n), 0) & Format(d(1, n), 0) & Format(d(0, n), 0) & Format(2 ^ (8 * kmd(1, n) - 1) + 2 ^ (8 * kmd(2, n) - 2) + 2 ^ (8 * kmd(3, n) - 3) + 2 ^ (8 * kmd(4, n) - 4) + 2 ^ (8 * kmd(5, n) - 5) + 2 ^ (8 * kmd(6, n) - 6) + 2 ^ (8 * kmd(7, n) - 7) + 2 ^ (8 * kmd(8, n) - 8), "000000000000000000000000000000000000000000000000000000000000000000")
            BackgroundWorker1.ReportProgress(((14 * q + n) / (20 * q)) * 9000 + 1000)
            '成绩越好w（）的制约高（虽然w（）为String）
        Next
        Dim x(q - 1) As Integer

        Dim y(q - 1) As String
        '开始排序
        For n = 1 To q - 1
            x(n) = n
            y(n) = w(n)
            BackgroundWorker1.ReportProgress(((15 * q + n) / (20 * q)) * 9000 + 1000)
        Next

        Dim m As Integer, t As String
        For n = 1 To q - 1
            For m = n To q - 1
                If y(n) < y(m) Then
                    t = y(m)
                    y(m) = y(n)
                    y(n) = t
                End If
            Next
            BackgroundWorker1.ReportProgress(((16 * q + n) / (20 * q)) * 9000 + 1000)
        Next

        For n = 1 To q - 1
            For m = q - 1 To 1 Step -1
                If w(n) = y(m) Then x(n) = m
            Next
            '排序结束
            BackgroundWorker1.ReportProgress(((17 * q + n) / (20 * q)) * 9000 + 1000)
        Next
        '开始输出结果等级排名
        xlsheet.Cells(1, i).Value = "等级排名"
        For n = 1 To q - 2
            xlsheet.Cells(n + 1, i) = x(n)
            BackgroundWorker1.ReportProgress(((18 * q + n) / (20 * q)) * 9000 + 1000)
        Next

        '整理各人等级分布情况
        Dim z(9, q - 1) As String
        For n = 1 To q - 2
            If d(9, n) = 0 Then
                z(9, n) = "" '若没有Tkm9的等级，则不显示（即不会显示0A+）
            Else : z(9, n) = d(9, n) & Tdj9.Text '若有Tkm9的等级，则显示
            End If
            If d(8, n) = 0 Then
                z(8, n) = ""
            Else : z(8, n) = d(8, n) & Tdj8.Text
            End If

            If d(7, n) = 0 Then
                z(7, n) = ""
            Else : z(7, n) = d(7, n) & Tdj7.Text
            End If

            If d(6, n) = 0 Then
                z(6, n) = ""
            Else : z(6, n) = d(6, n) & Tdj6.Text
            End If

            If d(5, n) = 0 Then
                z(5, n) = ""
            Else : z(5, n) = d(5, n) & Tdj5.Text
            End If

            If d(4, n) = 0 Then
                z(4, n) = ""
            Else : z(4, n) = d(4, n) & Tdj4.Text
            End If

            If d(3, n) = 0 Then
                z(3, n) = ""
            Else : z(3, n) = d(3, n) & Tdj3.Text
            End If

            If d(2, n) = 0 Then
                z(2, n) = ""
            Else : z(2, n) = d(2, n) & Tdj2.Text
            End If

            If d(1, n) = 0 Then
                z(1, n) = ""
            Else : z(1, n) = d(1, n) & Tdj0.Text
            End If

            If d(0, n) = 0 Then
                z(0, n) = ""
            Else : z(0, n) = d(0, n) & Tdj0.Text
            End If

            z(0, n) = Replace(Replace(km(0, n), "(", ""), ")", "") & "(" & z(9, n) & z(8, n) & z(7, n) & z(6, n) & z(5, n) & z(4, n) & z(3, n) & z(2, n) & z(1, n) & z(0, n) & ")" '合并
            xlsheet.Cells(n + 1, i + 1) = z(0, n) '输出等级分布
            BackgroundWorker1.ReportProgress(((19 * q + n) / (20 * q)) * 9000 + 1000)
        Next
        xlsheet.Cells(1, i + 1).Value = "等级分布"
90:
        xlBook.Close(True) '关闭工作簿
        xlApp.Quit() '结束EXCEL对象
        xlApp = Nothing '释放xlApp对象

        BackgroundWorker1.ReportProgress(10000)
        BackgroundWorker1.ReportProgress(0)
        Exit Sub
80:     '处理行295至行328传来的错误
        If h = 3 Then GoTo 90
        If h = 4 Then GoTo 5
        If h = 5 Then GoTo 70
    End Sub
    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged '报告进度
        ProgressBar1.Value = e.ProgressPercentage.ToString()
        If ProgressBar1.Value = 10000 Then
            Button1.Enabled = True
            Me.Cursor = Cursors.Default
        End If

    End Sub

End Class
