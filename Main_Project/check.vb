Imports System.IO
Imports System.Text
Imports PIS炉温采集系统.Ex
Imports PIS炉温采集系统.cn
Public Class check
    Public str, i, pathstr As String
    Dim j As Integer
    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim iniwriter As New StreamWriter(".\floder.ini", False, Encoding.Unicode) '写入合并文件路径以便下次使用
        iniwriter.WriteLine("floder1:" & TextBox1.Text)
        iniwriter.WriteLine("floder2:" & TextBox2.Text)
        iniwriter.WriteLine("floder3:" & TextBox3.Text)
        iniwriter.Close()
        'Readtable(datacom, Test.ComboBox1)
        'Readtable(datacom, History.ComboBox1)
        Test.Show()
        Me.Hide()
    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim IniReader As StreamReader = File.OpenText(".\floder.ini")   '读取合并文件路径
        str = IniReader.ReadLine
        TextBox1.Text = str.Substring(str.IndexOf("floder1") + 8)
        str = IniReader.ReadLine
        TextBox2.Text = str.Substring(str.IndexOf("floder2") + 8)
        str = IniReader.ReadLine
        TextBox3.Text = str.Substring(str.IndexOf("floder3") + 8)
        IniReader.Close()
        '将路径文件名读取在表格中
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListBox3.Items.Clear()
        pathstr = TextBox1.Text & "\"
        i = Dir(pathstr, vbDirectory)
        Do While i <> ""
            If i <> "." And i <> ".." And (i.ToUpper.LastIndexOf(".XL") > 0 Or i.ToUpper.LastIndexOf(".CSV") > 0) Then
                ListBox1.Items.Add(i)
            End If
            i = Dir()
        Loop
        pathstr = TextBox2.Text & "\"
        i = Dir(pathstr, vbDirectory)
        Do While i <> ""
            If i <> "." And i <> ".." And (i.ToUpper.LastIndexOf(".XL") > 0 Or i.ToUpper.LastIndexOf(".CSV") > 0) Then
                ListBox2.Items.Add(i)
            End If
            i = Dir()
        Loop
        pathstr = TextBox3.Text & "\"
        i = Dir(pathstr, vbDirectory)
        Do While i <> ""
            If i <> "." And i <> ".." And (i.ToUpper.LastIndexOf(".XL") > 0 Or i.ToUpper.LastIndexOf(".CSV") > 0) Then
                ListBox3.Items.Add(i)
            End If
            i = Dir()
        Loop
    End Sub
    Private Sub TextBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.Click
        '选择文件夹路径
        FolderBrowserDialog1.SelectedPath = TextBox1.Text
        Dim fl As Integer = FolderBrowserDialog1.ShowDialog()
        If fl = 1 Then
            TextBox1.Text = FolderBrowserDialog1.SelectedPath
            '将文件名读取在表格1中
            ListBox1.Items.Clear()
            pathstr = TextBox1.Text & "\"
            i = Dir(pathstr, vbDirectory)
            Do While i <> ""
                If i <> "." And i <> ".." And (i.ToUpper.LastIndexOf(".XL") > 0 Or i.ToUpper.LastIndexOf(".CSV") > 0) Then
                    ListBox1.Items.Add(i)
                End If
                i = Dir()
            Loop
        End If
    End Sub
    Private Sub TextBox2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.Click
        '选择文件夹路径
        FolderBrowserDialog1.SelectedPath = TextBox2.Text
        Dim fl As Integer = FolderBrowserDialog1.ShowDialog()
        If fl = 1 Then
            TextBox2.Text = FolderBrowserDialog1.SelectedPath
            '将文件名读取在表格2中
            ListBox2.Items.Clear()
            pathstr = TextBox2.Text & "\"
            i = Dir(pathstr, vbDirectory)
            Do While i <> ""
                If i <> "." And i <> ".." And (i.ToUpper.LastIndexOf(".XL") > 0 Or i.ToUpper.LastIndexOf(".CSV") > 0) Then
                    ListBox2.Items.Add(i)
                End If
                i = Dir()
            Loop
        End If
    End Sub
    Private Sub TextBox3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox3.Click
        '选择文件夹路径
        FolderBrowserDialog1.SelectedPath = TextBox3.Text
        Dim fl As Integer = FolderBrowserDialog1.ShowDialog()
        If fl = 1 Then
            TextBox3.Text = FolderBrowserDialog1.SelectedPath
            '将文件名读取在表格3中
            ListBox3.Items.Clear()
            pathstr = TextBox3.Text & "\"
            i = Dir(pathstr, vbDirectory)
            Do While i <> ""
                If i <> "." And i <> ".." And (i.ToUpper.LastIndexOf(".XL") > 0 Or i.ToUpper.LastIndexOf(".CSV") > 0) Then
                    ListBox3.Items.Add(i)
                End If
                i = Dir()
            Loop
        End If
    End Sub
    'Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
    '    Dim iniwriter As New StreamWriter(".\floder.ini", False, Encoding.Unicode) '写入合并文件路径以便下次使用
    '    iniwriter.WriteLine("floder1:" & TextBox1.Text)
    '    iniwriter.WriteLine("floder2:" & TextBox2.Text)
    '    iniwriter.WriteLine("floder3:" & TextBox3.Text)
    '    iniwriter.Close()
    '    'Readtable(datacom, Test.ComboBox1)
    '    'Readtable(datacom, History.ComboBox1)
    '    Test.Show()
    '    Me.Hide()
    'End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If (ListBox1.SelectedItem = Nothing) Or (ListBox2.SelectedItem = Nothing) Then
            MsgBox("请选择需要合并的两个文件")
        Else
            'cn.NewBaseprofileTab(ListBox1.SelectedItem, Readtable(datacom, Test.ComboBox1))
            'Ex.PopulateExcel(ListBox1.SelectedItem, TextBox1.Text & "\" & ListBox1.SelectedItem, TextBox2.Text & "\" & ListBox2.SelectedItem, ProgressBar1, Test.data)
            j = Base_Log(ListBox1.SelectedItem, ListBox2.SelectedItem)
            ListBox3.Items.Clear()
            pathstr = TextBox3.Text & "\"
            i = Dir(pathstr, vbDirectory)
            Do While i <> ""
                If i <> "." And i <> ".." And (i.ToUpper.LastIndexOf(".XL") > 0 Or i.ToUpper.LastIndexOf(".CSV") > 0) Then
                    ListBox3.Items.Add(i)
                End If
                i = Dir()
            Loop
        End If
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If ListBox3.SelectedItem = Nothing Then
            MsgBox("请选择需要删除的文件")
        Else
            pathstr = TextBox3.Text & "\"
            i = Dir(pathstr, vbDirectory)
            Try
                My.Computer.FileSystem.DeleteFile(pathstr & ListBox3.SelectedItem, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            ListBox3.Items.Clear()
            pathstr = TextBox3.Text
            i = Dir(pathstr & "\", vbDirectory)
            Do While i <> ""
                If i <> "." And i <> ".." And (i.ToUpper.LastIndexOf(".XL") > 0 Or i.ToUpper.LastIndexOf(".CSV") > 0) Then
                    ListBox3.Items.Add(i)
                End If
                i = Dir()
            Loop
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If ListBox3.SelectedItem <> Nothing Then
            Test.data = Base(2000, ListBox3.SelectedItem)
        End If
    End Sub
End Class
