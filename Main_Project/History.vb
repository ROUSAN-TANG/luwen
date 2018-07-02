Imports System.Data.OleDb
Imports PIS炉温采集系统.cn
Imports PIS炉温采集系统.Ex
Public Class History
    Private Sub History_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Test.Show()
    End Sub

    Private Sub History_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Newcon()
        'NewHistory()
        'Readtable(historycom, ComboBox1)       '将数据库表格加入到Combox1
        'History_Log("", DataGridView1)
        Dim i As String = Dir(Application.StartupPath & "\TestLog\", vbDirectory)
        Do While i <> ""
            If i <> "." And i <> ".." And (i.ToUpper.LastIndexOf(".XL") > 0 Or i.ToUpper.LastIndexOf(".CSV") > 0) Then
                ComboBox1.Items.Add(i.Replace(".xlsx", ""))
            End If
            i = Dir()
        Loop
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        MsgBox("必须选择一个机种" & vbCrLf & "采集时间间隔为0.5,如(1.5),若为空则查询所有采集时间的数据" & vbCrLf & "若采集设备项为空则查询5台机器的炉温值")
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim tx As String = ""
        If ComboBox1.SelectedItem <> "Test" Then
            historycom.Open()
            If TextBox2.Text = "" Then              '未输入采集时间
                If ComboBox2.SelectedItem <> "" Then    '输入了采集设备
                    tx = "select 记录时间,时间, [" & ComboBox2.SelectedItem & "] from [" & ComboBox1.SelectedItem & "]  where 记录时间='" & DateTimePicker1.Value.ToString("yyyy_MM_dd") & "'"
                End If
                If ComboBox2.SelectedItem = "" Then      '未输入采集设备
                    tx = "select * from [" & ComboBox1.SelectedItem & "]  where 记录时间='" & DateTimePicker1.Value.ToString("yyyy_MM_dd") & "'"
                End If
            Else                                    '输入了采集时间
                If ComboBox2.SelectedItem <> "" Then    '输入了采集设备
                    tx = "select 记录时间,时间,[" & ComboBox2.SelectedItem & "] from [" & ComboBox1.SelectedItem & "]  where 记录时间='" & DateTimePicker1.Value.ToString("yyyy_MM_dd") & "' and 时间= " & TextBox2.Text
                End If
                If ComboBox2.SelectedItem = "" Then      '未输采集设备
                    tx = "select * from [" & ComboBox1.SelectedItem & "]  where 记录时间='" & DateTimePicker1.Value.ToString("yyyy_MM_dd") & "' and 时间= " & TextBox2.Text
                End If
            End If
            Dim dataadapt As New OleDbDataAdapter
            Dim dst As New DataSet
            Dim dbl As OleDbCommand = New OleDbCommand(tx, historycom)
            dataadapt.SelectCommand = dbl
            dataadapt.Fill(dst, "info")
            dt = dst.Tables("info")
            DataGridView1.AutoGenerateColumns = True
            DataGridView1.DataSource = dt
            dbl = Nothing

        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        History_Log(ComboBox1.SelectedItem, DataGridView1)
    End Sub
End Class