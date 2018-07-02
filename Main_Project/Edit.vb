Public Class Edit

    Private Sub Edit_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        FormBorderStyle = Windows.Forms.FormBorderStyle.FixedSingle
        TextBox1.Text = Test.ListBox1.Items(5).ToString.Substring(4) '板邊
        TextBox2.Text = Test.ListBox1.Items(6).ToString.Substring(8) '板長
        TextBox3.Text = Test.ListBox1.Items(4).ToString.Substring(6) '鍊速
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Test.PCB_Edges = Val(TextBox1.Text)
        Test.PCB_Length = Val(TextBox2.Text)
        Test.Velocity = Val(TextBox3.Text)
        Test.ListBox1.Items(5) = "板邊： " + Test.PCB_Edges.ToString + "mm"
        Test.ListBox1.Items(4) = "當前链速： " + Test.Velocity.ToString + "cm/s"
        Test.DataGridView3.Rows(0).Cells(0).Value = Test.Velocity.ToString + "(cm/s)"
        Test.OledbConnection = cn.New_Datacon()
        With Test.OledbCommand
            .Connection = Test.OledbConnection
            .CommandType = CommandType.Text
            .CommandText = "update  [Modle] set [PCBLength] ='" + Test.PCB_Length.ToString + "' " + "where Model= '" + Test.ComboBox1.SelectedItem + "'"
            .ExecuteNonQuery()
        End With
        Test.OledbCommand.Dispose()
        Test.OledbConnection.Close()
        Test.DataInitialization()
        Me.Close()
    End Sub

    Private Sub Group_Enter(sender As Object, e As EventArgs) Handles Group.Enter

    End Sub
End Class