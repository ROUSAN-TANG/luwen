Public Class NewModle
    Private Sub NewModle_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
        TextBox1.Focus()
        FormBorderStyle = Windows.Forms.FormBorderStyle.FixedSingle
    End Sub
    ''' <summary>
    ''' 取消
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
    ''' <summary>
    ''' 確認
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text <> "" And TextBox2.Text <> "" And TextBox6.Text <> "" Then
            Test.PCB_Length = Val(TextBox2.Text)
            Test.PCB_Edges = Val(TextBox6.Text)
            Test.OledbConnection = cn.New_Datacon()
            With Test.OledbCommand
                .Connection = Test.OledbConnection
                .CommandType = CommandType.Text
                .CommandText = "INSERT INTO Modle (Model, PCBLength, [PCB-Edges], BaseLog, Base ,CreatTime,MultiplePCB) VALUES  ('" +
                            TextBox1.Text + "'," + Test.PCB_Length.ToString + "," + Test.PCB_Edges.ToString + ", 'null', 'null','" + Now.ToString("yyyyMMddHHmmss") + "'," + CheckBox1.Checked.ToString + ")"
                .ExecuteNonQuery()
            End With
            Test.OledbCommand.Dispose()
            Test.OledbConnection.Close()
            Test.LoadStart()
            Test.DataInitialization()
            Me.Close()
        Else
            MessageBox.Show("任意一項不能爲空！")
        End If
    End Sub
    ''' <summary>
    ''' 確認
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub TextBox6_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox6.KeyDown
        If e.KeyValue = 13 Then
            If TextBox1.Text <> "" And TextBox2.Text <> "" And TextBox6.Text <> "" Then
                Test.OledbConnection = cn.New_Datacon()
                With Test.OledbCommand
                    .Connection = Test.OledbConnection
                    .CommandType = CommandType.Text
                    .CommandText = "INSERT INTO Modle (Model, PCBLength, [PCB-Edges], BaseLog, Base, Test,CreatTime) VALUES  ('" +
                                TextBox1.Text + "'," + TextBox2.Text + "," + TextBox6.Text + ", 'null', 'null', 'null'," + Now.ToString("yyyyMMddHHmmss") + ")"
                    .ExecuteNonQuery()
                End With
                Test.OledbCommand.Dispose()
                Test.OledbConnection.Close()
                Test.LoadStart()
                Test.DataInitialization()
                Me.Close()
            Else
                MessageBox.Show("任意一項不能爲空！")
            End If
        End If
    End Sub
End Class