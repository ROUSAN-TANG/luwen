Public Class MultiplePCB
    Dim i As Integer
    Private Sub MultiplePCB_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        FormBorderStyle = Windows.Forms.FormBorderStyle.FixedSingle
        i = 1
        ListBox1.Items.Clear()
        TextBox1.Focus()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If Test.ComboBox1.SelectedItem <> "" And Test.ComboBox5.SelectedItem <> "" Then
            Test.InSenser = True
            Test.ChartInitialization() '曲線圖像初始化
            Test.t0(Test.InNum) = Now.ToString("yyyy/MM/dd HH:mm:ss")
            For j = 0 To i - 2
                '記錄當前條碼
                Test.TestSN(Test.InNum) += ListBox1.Items.Item(j).ToString.Remove(0, 2) + "|"
            Next
        Else
            MessageBox.Show("請選擇機種與基準文件")
        End If
        i = 1
        ListBox1.Items.Clear()
        TextBox1.Focus()
    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyValue = 13 And TextBox1.Text.Length > 3 Then
            ListBox1.Items.Add(i.ToString + ":" + TextBox1.Text)
            TextBox1.Text = ""
            i += 1
            TextBox1.Focus()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ListBox1.Items.Clear()
        i = 1
        TextBox1.Focus()
    End Sub
End Class