Imports System.Threading.Thread
Imports System.Data.OleDb
Imports System.Math
Public Class Test
#Region "Definition"
    '以下爲溫度採集數據定義
    Public Length_Machine, Senser_Distant,
    Senser_Num, Velocity,
    PCB_Edges, PCB_Length As Single     '機器總長度/傳感器距離/傳感器數量/链速/板邊/PCB板長度
    Public DataA_Value(15), DataB_Value(15) As Single  '16通道數據,A,B仪器
    Dim T(16), DifferenceValue(2),
        tt(4, 15) As Single  '每个传感器间距的时间Tn/t1t2t3時間間隔/各區域t1t2t3t4t5到達時間
    Dim BaseData(4, 1500), Offset(2000, 4), TestData(20, 4, 1500) As Single
    '每個探頭的所有時間段的參考值，補償值，測試值
    Public t0(20), TestSN(20), EnterString As String '記錄採集時間，記錄條碼
    Public TestNum, InNum, OutNum, TestLength(20) As Integer '當前爐子含有的測試序號,記錄進板序號/記錄出板序號/記錄當前測試數據長度用於採集生成Excel
    Public InSenser, OutSenser As Boolean '進板傳感器&出板傳感器
    Public ModleName, BaseModle, OffSetModle As String '機種名稱/参考文件/基准文件

    '以下爲CPK計算定義參數
    Public TestallNum As Integer '計算總的測試數量
    Public MaxTestValue(4), MaxUpValue(4), MaxDownValue(4) As Single '最大溫度,最大上升下降溫度
    Public SoakTime(4), RefluxingTime(4) As Single '恆溫時間，迴流時間
    Public MaxValueTime As Boolean '是否到達峯值時間
    '以下定義爲報警使用
    Public Err As Boolean = False

    '以下爲數據庫處理定義
    Public OledbConnection As OleDbConnection
    Public OledbCommand As OleDbCommand = New OleDbCommand

    '以下爲多線程任務定義
    Public task1, task2, task3 As Threading.Thread


    '以下爲WindowsFrom處理定義
    Public ProductionSite As String '生成地點
    Dim ComboxSelectItem(3) As Integer
    Public StopGetData, MultiPCB, WarmingUp As Boolean
#End Region
    ''' <summary>
    ''' 延時函數
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Declare Function timeGetTime Lib "winmm.dll" () As Integer
    Public Function Sleep1(ByVal T As Integer)
        Dim Savetime As Integer
        Savetime = timeGetTime
        While timeGetTime < Savetime + T
            Application.DoEvents()
        End While
        Return 1
    End Function
    ''' <summary>
    ''' Modbus 通信協議CRC校驗返回十進制
    ''' </summary>
    ''' <param name="Checkdata"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function CRC_CLC16(ByVal Checkdata() As Byte) As Int32
        Dim CRC_L, CRC_H As Byte        '初值0XFFFF
        Dim CRCL, CRCH As Byte          '多項式碼0XA001
        Dim CRC_TEMP_L, CRC_TEMP_H As Byte
        Dim num As Integer = 0
        CRC_L = &HFF
        CRC_H = &HFF
        CRCL = &H1
        CRCH = &HA0
        For num = 0 To UBound(Checkdata)
            CRC_L = (CRC_L Xor Checkdata(num))
            For n = 0 To 7
                CRC_TEMP_H = CRC_H
                CRC_TEMP_L = CRC_L
                CRC_H = CRC_H >> 1
                CRC_L = CRC_L >> 1
                If ((CRC_TEMP_H And &H1) = &H1) Then
                    CRC_L = (CRC_L Or &H80)
                End If
                If ((CRC_TEMP_L And &H1) = &H1) Then
                    CRC_H = (CRC_H Xor CRCH)
                    CRC_L = (CRC_L Xor CRCL)
                End If
            Next
        Next
        Return CRC_L * 256 + CRC_H
    End Function
    ''' <summary>
    '''========================发送十六进制数=========================================='
    ''' </summary>
    ''' <param name="Serial"></param>
    ''' <param name="sbuf"></param>
    ''' <remarks></remarks>
    Sub Send_H(ByVal Serial As System.IO.Ports.SerialPort, ByVal sbuf As String)

        Dim i, j, k As Integer
        sbuf = sbuf.Replace(" ", "").Trim
        j = Len(sbuf)
        Dim buf(j / 2) As Byte
        k = 0
        For i = 0 To j / 2 - 1
            buf(i) = "&H" & sbuf.Substring(k, 2)
            k += 2
        Next
        Serial.Write(buf, 0, j / 2)
    End Sub
    ''' <summary>
    ''' ========================接收十六进制数=========================================='
    ''' </summary>
    ''' <param name="Serial"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Read_H(ByVal Serial As System.IO.Ports.SerialPort) As Byte()
        Dim i, num As Integer
        Dim errreceive(2) As Byte
        errreceive = {0, 0}
        num = 0
        num = Serial.BytesToRead
        If num > 0 Then
            Dim rebuf(num - 1), crc(num - 3) As Byte
            Serial.Read(rebuf, 0, num)
            For i = 0 To num - 3
                crc(i) = rebuf(i)
            Next
            Dim s As Integer = CRC_CLC16(crc)
            If (rebuf(num - 2) * 256 + rebuf(num - 1)) = s Then
                Return rebuf
            Else
                Return errreceive
            End If
        End If
        Return errreceive
    End Function
    ''' <summary>
    ''' 讀取指定傳感器溫度
    ''' </summary>
    ''' <param name="Serial"></param>
    ''' <param name="Data_Changle"></param>
    ''' <param name="Data_Num"></param>
    ''' <remarks></remarks>
    Private Function Get_Data(ByVal Serial As System.IO.Ports.SerialPort, ByVal Data_Changle As Integer, ByVal Data_Num As Integer) As Single()
        Dim Data_Value(15) As Single
        If Data_Num = 1 Then
            Select Case Data_Changle
                Case 1 : Send_H(Serial, "01 03 00 64 00 01 C5 D5")
                Case 2 : Send_H(Serial, "01 03 00 65 00 01 94 15")
                Case 3 : Send_H(Serial, "01 03 00 66 00 01 64 15")
                Case 4 : Send_H(Serial, "01 03 00 67 00 01 35 D5")
                Case 5 : Send_H(Serial, "01 03 00 68 00 01 05 D6")
                Case 6 : Send_H(Serial, "01 03 00 69 00 01 54 16")
                Case 7 : Send_H(Serial, "01 03 00 6A 00 01 A4 16")
                Case 8 : Send_H(Serial, "01 03 00 6B 00 01 F5 D6")
                Case 9 : Send_H(Serial, "01 03 00 6C 00 01 44 17")
                Case 10 : Send_H(Serial, "01 03 00 6D 00 01 15 D7")
                Case 11 : Send_H(Serial, "01 03 00 6E 00 01 E5 D7")
                Case 12 : Send_H(Serial, "01 03 00 6F 00 01 B4 17")
                Case 13 : Send_H(Serial, "01 03 00 70 00 01 85 D1")
                Case 14 : Send_H(Serial, "01 03 00 71 00 01 D4 11")
                Case 15 : Send_H(Serial, "01 03 00 72 00 01 24 11")
                Case 16 : Send_H(Serial, "01 03 00 73 00 01 75 D1")
            End Select
        ElseIf Data_Num = 2 Then
            Select Case Data_Changle
                Case 1 : Send_H(Serial, "01 03 00 64 00 02 85 D4")
                Case 2 : Send_H(Serial, "01 03 00 65 00 02 D4 14")
                Case 3 : Send_H(Serial, "01 03 00 66 00 02 24 14")
                Case 4 : Send_H(Serial, "01 03 00 67 00 02 75 D4")
                Case 5 : Send_H(Serial, "01 03 00 68 00 02 45 D7")
                Case 6 : Send_H(Serial, "01 03 00 69 00 02 14 17")
                Case 7 : Send_H(Serial, "01 03 00 6A 00 02 E4 17")
                Case 8 : Send_H(Serial, "01 03 00 6B 00 02 B5 D7")
                Case 9 : Send_H(Serial, "01 03 00 6C 00 02 04 16")
                Case 10 : Send_H(Serial, "01 03 00 6D 00 02 55 D6")
                Case 11 : Send_H(Serial, "01 03 00 6E 00 02 A5 D6")
                Case 12 : Send_H(Serial, "01 03 00 6F 00 02 F4 16")
                Case 13 : Send_H(Serial, "01 03 00 70 00 02 C5 D0")
                Case 14 : Send_H(Serial, "01 03 00 71 00 02 94 10")
                Case 15 : Send_H(Serial, "01 03 00 72 00 02 64 10")
            End Select
        ElseIf Data_Num = 4 Then
            Select Case Data_Changle
                Case 1 : Send_H(Serial, "01 03 00 64 00 04 05 D6")
                Case 5 : Send_H(Serial, "01 03 00 68 00 04 C5 D5")
                Case 9 : Send_H(Serial, "01 03 00 6C 00 04 84 14")
                Case 13 : Send_H(Serial, "01 03 00 70 00 04 45 D2")
            End Select
        ElseIf Data_Num = 8 Then
            Select Case Data_Changle
                Case 1 : Send_H(Serial, "01 03 00 64 00 08 05 D3")
                Case 9 : Send_H(Serial, "01 03 00 6C 00 08 84 11")
            End Select
        Else
            Send_H(Serial, "01 03 00 64 00 10 05 D9")
        End If
        Dim rebuf() As Byte = Read_H(Serial)
        If rebuf.Length > 2 Then
            Dim i, j As Integer
            j = 0
            For i = Data_Changle - 1 To Data_Changle - 1 + Data_Num - 1
                Data_Value(i) = (rebuf(3 + j) * 256 + rebuf(4 + j)) / 10
                j += 2
            Next
        End If
        Return Data_Value
    End Function
    ''' <summary>
    ''' label的委託任務
    ''' </summary>
    ''' <param name="i"></param>
    ''' <param name="obj"></param>
    ''' <remarks></remarks>
    Delegate Sub bl(ByRef i As String, ByVal obj As Label)
    Sub tx(ByRef i As String, ByVal obj As Label)
        Dim st As Label
        st = obj
        st.Text = i.ToString
    End Sub
    ''' <summary>
    ''' TextBox的委託任務
    ''' </summary>
    ''' <param name="i"></param>
    ''' <param name="obj"></param>
    ''' <remarks></remarks>
    Delegate Sub TextboxInvoke(ByRef i As String, ByVal obj As TextBox)
    Sub TextboxIn(ByRef i As String, ByVal obj As TextBox)
        Dim st As TextBox
        st = obj
        st.Text = i.ToString
    End Sub
    ''' <summary>
    ''' Listbox委託任務
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <remarks></remarks>
    Delegate Sub ListboxInvoke(ByVal obj As ListBox)
    Sub ListboxIn(ByVal obj As ListBox)
        Dim list As ListBox
        Dim i As Integer
        list = obj
        For i = 1 To 20
            If TestSN(i - 1) <> "" Then
                list.Items.Item(i) = i.ToString + ":" + TestSN(i - 1)
            Else
                list.Items.Item(i) = i.ToString + ":N"
            End If
        Next
    End Sub
    Delegate Sub ComboxInvoke(ByVal obj As ComboBox, ByVal PathString As String)
    Sub ComboxIn(ByVal obj As ComboBox, ByVal PathString As String)
        Dim Box As ComboBox = obj
        '获取所有Baseprofie文件
        Box.Items.Clear()
        Do While PathString <> ""
            If PathString <> "." And PathString <> ".." And (PathString.ToUpper.LastIndexOf(".XL") > 0 Or PathString.ToUpper.LastIndexOf(".CSV") > 0) Then
                Box.Items.Add(PathString)
            End If
            PathString = Dir()
        Loop
    End Sub
    Delegate Sub DataViewInvoke(ByVal obj1 As DataGridView, ByVal obj2 As DataVisualization.Charting.Chart, ByVal TtNow As Single, ByVal TestNumOfNow As Integer, ByVal SelectChannel As Integer)
    Sub DataView(ByVal obj1 As DataGridView, ByVal obj2 As DataVisualization.Charting.Chart, ByVal TtNow As Single, ByVal TestNumOfNow As Integer, ByVal SelectChannel As Integer)
        Dim DataView As DataGridView = obj1
        Dim ChartView As DataVisualization.Charting.Chart = obj2
        ChartView.Series(0).Points.AddXY(TtNow / 2, TestData(TestNumOfNow, 0, TtNow))
        ChartView.Series(1).Points.AddXY(TtNow / 2, TestData(TestNumOfNow, 2, TtNow))
        ChartView.Series(2).Points.AddXY(TtNow / 2, TestData(TestNumOfNow, 3, TtNow))
        For i = 0 To 4
            DataView.Rows(i + 1).Cells(SelectChannel + 2).Value = TestData(TestNumOfNow, i, TtNow).ToString("0.0")
        Next
    End Sub
    ''' <summary>
    ''' 顯示當前時間
    ''' </summary>
    ''' <remarks></remarks>
    Sub test1()
        Dim lb As New bl(AddressOf tx)
        Dim st As String
        While 1
            st = Now.ToString("yyyy-MM-dd HH:mm:ss")
            Me.Invoke(lb, st, Label4)
            Threading.Thread.Sleep(500)
        End While
    End Sub
    ''' <summary>
    ''' 進板出板
    ''' </summary>
    ''' <remarks></remarks>
    Sub InOut()
        Dim lb As New TextboxInvoke(AddressOf TextboxIn)
        Dim lb2 As New ListboxInvoke(AddressOf ListboxIn)
        Dim lb3 As New Ex.WriteTestLogFileInvoke(AddressOf Ex.WriteTestLogFile)
        Dim lb4 As New Ex.CopyTestFileInvoke(AddressOf Ex.CopyTestFile)
        While 1
            If InSenser And TestSN(InNum) <> "" Then
                InSenser = False
                If InNum >= 20 Then
                    InNum = 0
                End If
                Me.Invoke(lb, "", TextBox1)
                InNum += 1
            End If
            If OutSenser Then
                OutSenser = False
                '多連板處理方法爲1片當多片數據
                If TestSN(OutNum).IndexOf("|") > 0 Then
                    Dim SN() As String = Split(TestSN(OutNum), "|")
                    Me.Invoke(lb3, Application.StartupPath + "\TestLog\" + ModleName + "_" + SN(0) + ".xlsx", TestData, Offset, TestLength(OutNum), OutNum, SN(0), ProgressBar1)
                    InserToSQL(SN(0), "\TestLog\" + ModleName + "_" + SN(0) + ".xlsx")
                    For i = 1 To SN.Length - 2
                        Me.Invoke(lb4, Application.StartupPath + "\TestLog\" + ModleName + "_" + SN(0) + ".xlsx", Application.StartupPath + "\TestLog\" + ModleName + "_" + SN(i) + ".xlsx", SN(i))
                        InserToSQL(SN(i), "\TestLog\" + ModleName + "_" + SN(i) + ".xlsx")
                    Next
                Else
                    Me.Invoke(lb3, Application.StartupPath + "\TestLog\" + ModleName + "_" + TestSN(OutNum) + ".xlsx", TestData, Offset, TestLength(OutNum), OutNum, TestSN(OutNum), ProgressBar1)
                    InserToSQL(TestSN(OutNum), "\TestLog\" + ModleName + "_" + TestSN(OutNum) + ".xlsx")
                End If
                TestSN(OutNum) = ""
                If OutNum >= 20 Then
                    OutNum = 0
                End If
                OutNum += 1
            End If
            Me.Invoke(lb2, ListBox2)
            Sleep(100)
        End While
    End Sub
    ''' <summary>
    ''' 數據處理
    ''' </summary>
    ''' <remarks></remarks>
    Sub GetData()
        Dim lb As New DataViewInvoke(AddressOf DataView)
        Dim lb2 As New Ex.CopyTestFileInvoke(AddressOf Ex.CopyTestFile)
        Dim lb3 As New Ex.WriteBaseProFileInvoke(AddressOf Ex.WriteBaseProFile)
        Dim lb4 As New Ex.WriteTestLogFileInvoke(AddressOf Ex.WriteTestLogFile)
        Dim lb5 As New bl(AddressOf tx)
        Dim PathString As String = Dir(Application.StartupPath & "\Base_Log\", vbDirectory)
        Dim OldData(16) As Single '上一次采集的数据
        Dim k(20) As Single  '记录采集次数
        Dim TestNumOfNow As Integer '记录当前测试序号
        Dim SelectChangle(20) As Integer '记录当前到达的时刻
        WarmingUp = True
        MaxUpValue = {0, 0, 0, 0, 0} '最大上升斜率
        MaxDownValue = {0, 0, 0, 0, 0} '最小下降斜率
        MaxTestValue = {0, 0, 0, 0, 0} '最大測試溫度
        SoakTime = {0, 0, 0, 0, 0} '恆溫時間
        RefluxingTime = {0, 0, 0, 0, 0}  '迴流時間
        While 1
            '若有數據爲0則取上一次的數據
            If DataA_Value(0) <> 0 Then
                OldData = DataA_Value
            End If
            DataA_Value = Get_Data(TemperatureA, 1, 16)
            'DataB_Value = Get_Data(TemperatureB, 1, 16)
            If DataA_Value(0) = 0 Then
                DataA_Value = OldData
            End If
            For TestNumOfNow = 0 To 20
                If TestSN(TestNumOfNow) <> "" Then
                    Sleep(1)
                    If k(TestNumOfNow) / 2 <= tt(0, SelectChangle(TestNumOfNow)) Then
                        TestData(TestNumOfNow, 0, k(TestNumOfNow)) = DataA_Value(SelectChangle(TestNumOfNow)) + Offset(k(TestNumOfNow), 0)
                        TestData(TestNumOfNow, 1, k(TestNumOfNow)) = DataB_Value(SelectChangle(TestNumOfNow)) + Offset(k(TestNumOfNow), 1)
                        TestData(TestNumOfNow, 2, k(TestNumOfNow)) = DataA_Value(SelectChangle(TestNumOfNow)) + Offset(k(TestNumOfNow), 2)
                        TestData(TestNumOfNow, 3, k(TestNumOfNow)) = DataA_Value(SelectChangle(TestNumOfNow)) + Offset(k(TestNumOfNow), 3)
                        TestData(TestNumOfNow, 4, k(TestNumOfNow)) = DataB_Value(SelectChangle(TestNumOfNow)) + Offset(k(TestNumOfNow), 4)
                    ElseIf k(TestNumOfNow) / 2 > tt(0, SelectChangle(TestNumOfNow)) And k(TestNumOfNow) / 2 <= tt(2, SelectChangle(TestNumOfNow)) And SelectChangle(TestNumOfNow) < 15 Then
                        TestData(TestNumOfNow, 0, k(TestNumOfNow)) = DataA_Value(SelectChangle(TestNumOfNow) + 1) + Offset(k(TestNumOfNow), 0)
                        TestData(TestNumOfNow, 1, k(TestNumOfNow)) = DataB_Value(SelectChangle(TestNumOfNow) + 1) + Offset(k(TestNumOfNow), 1)
                        TestData(TestNumOfNow, 2, k(TestNumOfNow)) = DataA_Value(SelectChangle(TestNumOfNow)) + Offset(k(TestNumOfNow), 2)
                        TestData(TestNumOfNow, 3, k(TestNumOfNow)) = DataA_Value(SelectChangle(TestNumOfNow)) + Offset(k(TestNumOfNow), 3)
                        TestData(TestNumOfNow, 4, k(TestNumOfNow)) = DataB_Value(SelectChangle(TestNumOfNow)) + Offset(k(TestNumOfNow), 4)
                    ElseIf k(TestNumOfNow) / 2 > tt(2, SelectChangle(TestNumOfNow)) And k(TestNumOfNow) / 2 <= tt(4, SelectChangle(TestNumOfNow)) And SelectChangle(TestNumOfNow) < 15 Then
                        TestData(TestNumOfNow, 0, k(TestNumOfNow)) = DataA_Value(SelectChangle(TestNumOfNow) + 1) + Offset(k(TestNumOfNow), 0)
                        TestData(TestNumOfNow, 1, k(TestNumOfNow)) = DataB_Value(SelectChangle(TestNumOfNow) + 1) + Offset(k(TestNumOfNow), 1)
                        TestData(TestNumOfNow, 2, k(TestNumOfNow)) = DataA_Value(SelectChangle(TestNumOfNow) + 1) + Offset(k(TestNumOfNow), 2)
                        TestData(TestNumOfNow, 3, k(TestNumOfNow)) = DataA_Value(SelectChangle(TestNumOfNow)) + Offset(k(TestNumOfNow), 3)
                        TestData(TestNumOfNow, 4, k(TestNumOfNow)) = DataB_Value(SelectChangle(TestNumOfNow)) + Offset(k(TestNumOfNow), 4)
                    ElseIf k(TestNumOfNow) / 2 > tt(4, SelectChangle(TestNumOfNow)) And k(TestNumOfNow) / 2 <= T(SelectChangle(TestNumOfNow) + 1) Then
                        SelectChangle(TestNumOfNow) += 1
                        If SelectChangle(TestNumOfNow) >= 15 Then
                            TestLength(TestNumOfNow) = k(TestNumOfNow)
                            If TestNumOfNow = 20 Then
                                Me.Invoke(lb3, Application.StartupPath + "\Base_Log\" + EnterString + "V" + Velocity.ToString + ".xlsx", Offset, TestLength(20), ProgressBar1)
                            Else
                                '多連板處理方法爲1片當多片數據
                                If TestSN(OutNum).IndexOf("|") > 0 Then
                                    Dim SN() As String = Split(TestSN(OutNum), "|")
                                    Me.Invoke(lb3, Application.StartupPath + "\TestLog\" + ModleName + "_" + SN(0) + ".xlsx", TestData, Offset, TestLength(OutNum), OutNum, SN(0), ProgressBar1)
                                    InserToSQL(SN(0), "\TestLog\" + ModleName + "_" + SN(0) + ".xlsx")
                                    For i = 1 To SN.Length - 2
                                        Me.Invoke(lb2, Application.StartupPath + "\TestLog\" + ModleName + "_" + SN(0) + ".xlsx", Application.StartupPath + "\TestLog\" + ModleName + "_" + SN(i) + ".xlsx", SN(i))
                                        InserToSQL(SN(i), "\TestLog\" + ModleName + "_" + SN(i) + ".xlsx")
                                    Next
                                Else
                                    Me.Invoke(lb4, Application.StartupPath + "\TestLog\" + ModleName + "_" + TestSN(OutNum) + ".xlsx", TestData, Offset, TestLength(OutNum), OutNum, TestSN(OutNum), ProgressBar1)
                                    InserToSQL(TestSN(OutNum), "\TestLog\" + ModleName + "_" + TestSN(OutNum) + ".xlsx")
                                End If
                                OutNum += 1
                            End If
                            k(TestNumOfNow) = 0
                            TestSN(TestNumOfNow) = ""
                            SelectChangle(TestNumOfNow) = 0
                            Continue For
                        End If
                        TestData(TestNumOfNow, 0, k(TestNumOfNow)) = DataA_Value(SelectChangle(TestNumOfNow)) + Offset(k(TestNumOfNow), 0)
                        TestData(TestNumOfNow, 1, k(TestNumOfNow)) = DataB_Value(SelectChangle(TestNumOfNow)) + Offset(k(TestNumOfNow), 1)
                        TestData(TestNumOfNow, 2, k(TestNumOfNow)) = DataA_Value(SelectChangle(TestNumOfNow)) + Offset(k(TestNumOfNow), 2)
                        TestData(TestNumOfNow, 3, k(TestNumOfNow)) = DataA_Value(SelectChangle(TestNumOfNow)) + Offset(k(TestNumOfNow), 3)
                        TestData(TestNumOfNow, 4, k(TestNumOfNow)) = DataB_Value(SelectChangle(TestNumOfNow)) + Offset(k(TestNumOfNow), 4)
                    End If
                    k(TestNumOfNow) += 1
                End If
                If SelectChangle(TestNumOfNow) <= 15 Then
                    TestLength(TestNumOfNow) = k(TestNumOfNow)
                End If
            Next
            If InNum >= 1 Then
                If TestSN(InNum - 1) <> "" Then
                    Me.Invoke(lb, DataGridView3, Chart2, k(InNum - 1) - 1, InNum - 1, SelectChangle(InNum - 1))
                End If
                For i = 0 To 4
                    If k(InNum - 1) > 3 Then
                        If k(InNum - 1) > 3 Then
                            '計算最大測試溫度，計算所有的
                            If MaxTestValue(i) < TestData(InNum - 1, i, k(InNum - 1) - 1) Then
                                MaxTestValue(i) = TestData(InNum - 1, i, k(InNum - 1) - 1)
                            End If
                        End If
                        '計算最大上升斜率，只計算開始的
                        If MaxUpValue(i) < TestData(InNum - 1, i, k(InNum - 1) - 1) - TestData(InNum - 1, i, k(InNum - 1) - 2) Then
                            MaxUpValue(i) = TestData(InNum - 1, i, k(InNum - 1) - 1) - TestData(InNum - 1, i, k(InNum - 1) - 2)
                        End If
                        '計算最大下降斜率，只計算開始的
                        If MaxDownValue(i) > TestData(InNum - 1, i, k(InNum - 1) - 1) - TestData(InNum - 1, i, k(InNum - 1) - 2) Then
                            MaxDownValue(i) = TestData(InNum - 1, i, k(InNum - 1) - 1) - TestData(InNum - 1, i, k(InNum - 1) - 2)
                        End If
                        '計算恆溫時間150~190,測試數據減去上一次數據±1℃,並且超過190後則不在計算，只計算開始的
                        If MaxValueTime = False And TestData(InNum - 1, i, k(InNum - 1) - 1) >= 150 And TestData(InNum - 1, i, k(InNum - 1) - 1) <= 190 And TestData(InNum - 1, i, k(InNum - 1) - 1) - TestData(InNum - 1, i, k(InNum - 1) - 2) > -1 And TestData(InNum - 1, i, k(InNum - 1) - 1) - TestData(InNum - 1, i, k(InNum - 1) - 2) < 1 Then
                            SoakTime(i) += 0.5
                        End If
                        '計算迴流時間>=220，只計算開始的
                        If TestData(InNum - 1, i, k(InNum - 1) - 1) >= 220 Then
                            RefluxingTime(i) += 0.5
                            MaxValueTime = True
                        End If
                    End If
                Next
                '2.分析CPK
                If k(InNum - 1) > 10 Then
                    DataGridView4.Rows(0).Cells(0).Value = "CPK=" + Cpk(MaxUpValue, 3, 1).ToString  '最大上升斜率1-3
                    DataGridView4.Rows(0).Cells(1).Value = "CPK=" + Cpk(MaxDownValue, -1, -3).ToString  '最大下降斜率-3--1
                    DataGridView4.Rows(0).Cells(2).Value = "CPK=" + Cpk(SoakTime, 120, 80).ToString  '恆溫時間80-120
                    DataGridView4.Rows(0).Cells(3).Value = "CPK=" + Cpk(RefluxingTime, 40, 75).ToString  '迴流時間40-75
                    DataGridView4.Rows(0).Cells(4).Value = "CPK=" + Cpk(MaxTestValue, 245, 235).ToString '峯值溫度235-245

                End If
            End If
            If TestSN(20) <> "" Then
                '生成基準數據
                For i As Integer = 0 To 4
                    Offset(k(20) - 1, i) = BaseData(i, k(20) - 1) - TestData(20, i, k(20) - 1)
                    TestData(20, i, k(20) - 1) += Offset(k(20) - 1, i)
                Next
                Me.Invoke(lb, DataGridView3, Chart2, k(20) - 1, 20, SelectChangle(20))
            End If
            If StopGetData = True Then
                TestSN(20) = ""
                StopGetData = False
                Me.Invoke(lb3, Application.StartupPath + "\Base_Log\" + EnterString + "V" + Velocity.ToString + ".xlsx", Offset, TestLength(20), ProgressBar1)
                Sleep(50)
                k(20) = 0
                SelectChangle(20) = 0
            End If
            '以下爲統計信息
            Me.Invoke(lb5, "當前統計數量：" + TestallNum.ToString, Label19)
            '1.分析製程
            DataGridView1.Rows(1).Cells(0).Value = MaxUpValue.Max.ToString '最大上升斜率
            DataGridView1.Rows(1).Cells(1).Value = MaxDownValue.Max.ToString '最大下降斜率
            DataGridView1.Rows(1).Cells(2).Value = SoakTime.Max.ToString  '恆溫時間
            DataGridView1.Rows(1).Cells(3).Value = RefluxingTime.Max.ToString '迴流時間
            DataGridView1.Rows(1).Cells(4).Value = MaxTestValue.Max.ToString '峯值溫度
            ''2.分析CPK
            'If k(InNum - 1) > 10 Then
            '    DataGridView4.Rows(0).Cells(0).Value = "CPK=" + Cpk(MaxUpValue, 3, 1).ToString  '最大上升斜率1-3
            '    DataGridView4.Rows(0).Cells(1).Value = "CPK=" + Cpk(MaxDownValue, -1, -3).ToString  '最大下降斜率-3--1
            '    DataGridView4.Rows(0).Cells(2).Value = "CPK=" + Cpk(SoakTime, 120, 80).ToString  '恆溫時間80-120
            '    DataGridView4.Rows(0).Cells(3).Value = "CPK=" + Cpk(RefluxingTime, 40, 75).ToString  '迴流時間40-75
            '    DataGridView4.Rows(0).Cells(4).Value = "CPK="  '峯值溫度235-245

            'End If
            '3.分析爐區溫度
            For i = 1 To 16
                DataGridView5.Rows(1).Cells(i).Value = DataA_Value(i - 1) - DataGridView6.Rows(0).Cells(i).Value
                DataGridView5.Rows(2).Cells(i).Value = DataB_Value(i - 1) - DataGridView6.Rows(2).Cells(i).Value
            Next
            Sleep(500)
        End While
    End Sub
    ''' <summary>
    ''' 关闭进程
    ''' </summary>
    ''' <param name="ProcessString"></param>
    ''' <remarks></remarks>
    ''' 
    Private Sub killexcel(ByVal ProcessString As String)
        Try
            For Each proc As Process In Process.GetProcessesByName(ProcessString)
                If proc IsNot Nothing AndAlso proc.MainWindowTitle = "" Then
                    proc.Kill()
                End If
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' 關閉程序
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Test_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        task1.Abort()
        task2.Abort()
        If WarmingUp = True Then
            task3.Abort()
        End If
        killexcel("EXCEL")
    End Sub
    ''' <summary>
    ''' 啓動多線程任務，顯示當前時間並獲取當前時間刻度
    ''' </summary>
    ''' <remarks></remarks>
    Sub main()
        task1 = New Threading.Thread(AddressOf test1)
        task1.Start()
        task2 = New Threading.Thread(AddressOf InOut)
        task2.Start()
        task3 = New Threading.Thread(AddressOf GetData)
    End Sub
    ''' <summary>
    ''' 數據庫更新
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub LoadStart()
        Try
            DataGridView3.Rows.Add(6)
            ListBox1.Items.Clear()
            OledbConnection = cn.New_Datacon()
            With OledbCommand
                .Connection = OledbConnection
                .CommandType = CommandType.Text
                .CommandText = "SELECT   [Value] FROM ([Set])"
            End With
            Dim OledbReader As OleDbDataReader
            OledbReader = OledbCommand.ExecuteReader
            '獲取機器長度
            OledbReader.Read()
            Length_Machine = Val(OledbReader.GetString(0))
            ListBox1.Items.Add("迴流爐長度： " + Length_Machine.ToString + "cm")
            '獲取傳感器距離
            OledbReader.Read()
            Senser_Distant = Val(OledbReader.GetString(0))
            ListBox1.Items.Add("傳感器距離： " + Senser_Distant.ToString + "cm")
            '獲取傳感器距離
            OledbReader.Read()
            Senser_Num = Val(OledbReader.GetString(0))
            ListBox1.Items.Add("傳感器數量： " + Senser_Num.ToString)
            '获取串口並打開
            OledbReader.Read()
            If TemperatureA.IsOpen = True Then
                TemperatureA.Close()
            End If
            TemperatureA.PortName = OledbReader.GetString(0)
            If TemperatureA.IsOpen = False Then
                TemperatureA.Open()
            End If
            ListBox1.Items.Add("連接串口： " + TemperatureA.PortName)
            '獲取默認鍊速
            OledbReader.Read()
            Velocity = Val(OledbReader.GetString(0))
            ListBox1.Items.Add("當前链速： " + Velocity.ToString + "cm/s")
            '獲取默認板邊
            OledbReader.Read()
            PCB_Edges = Val(OledbReader.GetString(0)) / 10
            ListBox1.Items.Add("板邊： " + (PCB_Edges * 10).ToString + "mm")
            ListBox1.Items.Add("")
            OledbCommand.Dispose()
            OledbReader.Close()
            '設置爐溫參數
            DataGridView6.Rows.Add(4)
            DataGridView6.Rows(0).Cells(0).Value = "溫區A(℃)"
            DataGridView6.Rows(1).Cells(0).Value = "偏差(℃)"
            DataGridView6.Rows(2).Cells(0).Value = "溫區B(℃)"
            DataGridView6.Rows(3).Cells(0).Value = "偏差(℃)"
            For i = 1 To 10
                DataGridView6.Rows(0).Cells(i).Value = 160 + i * 10
                DataGridView6.Rows(2).Cells(i).Value = 160 + i * 10
            Next
            For i = 11 To 16
                DataGridView6.Rows(0).Cells(i).Value = 260 - (i - 10) * 10
                DataGridView6.Rows(2).Cells(i).Value = 260 - (i - 10) * 10
            Next
            '獲取所有已有機種
            ComboBox1.Items.Clear()
            ComboBox1.ResetText()
            With OledbCommand
                .Connection = OledbConnection
                .CommandType = CommandType.Text
                .CommandText = "SELECT   [Model] FROM ([Modle])"
            End With
            OledbReader = OledbCommand.ExecuteReader
            While OledbReader.Read
                ComboBox1.Items.Add(OledbReader.GetString(0))
            End While
            '將所有不相同機種添加到combox中
            OledbReader.Close()
            ComboBox4.Items.Clear()
            ComboBox4.Items.Add("ALL")
            OledbCommand.CommandText = "SELECT   distinct [Model] FROM ([TestData])"
            OledbReader = OledbCommand.ExecuteReader
            While OledbReader.Read
                ComboBox4.Items.Add(OledbReader.GetString(0))
            End While
            '將TestData數據放入歷史數據中
            Dim dt As DataSet = cn.Getds(OledbConnection, "Select * from [TestData]")
            DataGridView2.DataSource = dt.Tables(0)
            '關閉數據庫
            OledbCommand.Dispose()
            OledbReader.Close()
            OledbConnection.Close()
            'WindowsFrom界面處理
            ComboxSelectItem(0) = -1
            ComboxSelectItem(1) = -1
            ComboxSelectItem(2) = -1
            Button4.Enabled = False
            RadioButton1.Select()
            WarmingUp = False
            InNum = 0
            OutNum = 0
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            End
        End Try
    End Sub
    Sub DescriptionOfDatagridview()
        '1.分析製程
        DataGridView1.Rows.Clear()
        DataGridView1.Rows.Add(2)
        DataGridView1.Rows(0).Cells(0).Value = "1~3℃/s"
        DataGridView1.Rows(0).Cells(1).Value = "-3~-1℃/s"
        DataGridView1.Rows(0).Cells(2).Value = "80~120s"
        DataGridView1.Rows(0).Cells(3).Value = "40~75s"
        DataGridView1.Rows(0).Cells(4).Value = "235~245℃"
        '2.分析CPK
        DataGridView4.Rows.Clear()
        DataGridView4.Rows.Add()
        DataGridView4.Rows(0).Cells(0).Value = "CPK="
        DataGridView4.Rows(0).Cells(1).Value = "CPK="
        DataGridView4.Rows(0).Cells(2).Value = "CPK="
        DataGridView4.Rows(0).Cells(3).Value = "CPK="
        DataGridView4.Rows(0).Cells(4).Value = "CPK="
        '3.分析迴流爐爐區溫度
        DataGridView5.Rows.Clear()
        DataGridView5.Rows.Add(3)
        DataGridView5.Rows(0).Cells(0).Value = "標準(℃)"
        For i = 1 To 16
            DataGridView5.Rows(0).Cells(i).Value = DataGridView6.Rows(0).Cells(i).Value.ToString + "±5"
        Next
        DataGridView5.Rows(1).Cells(0).Value = "溫區A偏差(℃)"
        DataGridView5.Rows(2).Cells(0).Value = "溫區B偏差(℃)"
    End Sub
    ''' <summary>
    ''' 數據初始化
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub DataInitialization()
        Dim i As Integer
        OledbConnection = cn.New_Datacon()
        With OledbCommand
            .Connection = OledbConnection
            .CommandType = CommandType.Text
            .CommandText = "SELECT   [PCBLength] FROM ([Modle]) where Model='" + ComboBox1.SelectedItem + "'"
        End With
        PCB_Length = OledbCommand.ExecuteScalar / 10
        ListBox1.Items(6) = "PCB板長度： " + (PCB_Length * 10).ToString + "mm"
        OledbCommand.CommandText = "SELECT   [PCB-Edges] FROM ([Modle]) where Model='" + ComboBox1.SelectedItem + "'"
        PCB_Edges = OledbCommand.ExecuteScalar / 10
        ListBox1.Items(5) = "板邊： " + (PCB_Edges * 10).ToString + "mm"
        OledbCommand.CommandText = "SELECT   [Base] FROM ([Modle]) where Model='" + ModleName + "'"
        BaseModle = OledbCommand.ExecuteScalar
        '是否爲多連板
        OledbCommand.CommandText = "SELECT   [MultiplePCB] FROM ([Modle]) where Model='" + ModleName + "'"
        MultiPCB = OledbCommand.ExecuteScalar
        '如果是則打開多連板掃描窗口，否則關閉多連板掃描窗口
        If MultiPCB Then
            MultiplePCB.Show()
        Else
            If MultiplePCB.IsHandleCreated Then
                MultiplePCB.Close()
            End If
        End If
        OledbCommand.Dispose()
        OledbConnection.Close()
        '計算各個區域的時間Tn
        T(0) = Senser_Distant / Velocity
        For i = 1 To 16
            T(i) = T(0) * (i + 1)
        Next
        tt(0, 0) = ((Senser_Distant + PCB_Edges) / Velocity).ToString("0") 't1t2在#1區域時間
        tt(1, 0) = ((Senser_Distant + PCB_Edges) / Velocity).ToString("0") 't1t2在#1區域時間
        tt(2, 0) = ((2 * Senser_Distant + PCB_Length) / (2 * Velocity)).ToString("0") 't3在#1區域時間
        tt(3, 0) = ((Senser_Distant - PCB_Edges + PCB_Length) / Velocity).ToString("0") 't4t5在#1區域時間
        tt(4, 0) = ((Senser_Distant - PCB_Edges + PCB_Length) / Velocity).ToString("0") 't4t5在#1區域時間
        DifferenceValue(0) = tt(2, 0) - tt(1, 0)
        DifferenceValue(1) = tt(4, 0) - tt(1, 0)
        '計算各區域的時間ttn
        For i = 1 To 15
            tt(0, i) = (Senser_Distant * (i + 1) + PCB_Edges) / Velocity
            tt(1, i) = (Senser_Distant * (i + 1) + PCB_Edges) / Velocity 't1t2在#(i+1)區域時間
            tt(2, i) = tt(1, i) + DifferenceValue(0) 't3在#(i+1)區域時間
            tt(3, i) = tt(1, i) + DifferenceValue(1)
            tt(4, i) = tt(1, i) + DifferenceValue(1) 't4t5在#(i+1)區域時間
        Next
        ChartInitialization()
    End Sub
    ''' <summary>
    ''' 表格曲線圖像初始化
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ChartInitialization()
        Chart2.Series.Clear()
        Chart2.Series.Add("#1")
        'Chart2.Series.Add("#2")
        Chart2.Series.Add("#3")
        Chart2.Series.Add("#4")
        'Chart2.Series.Add("#5")
        Chart2.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline
        Chart2.Series(1).ChartType = DataVisualization.Charting.SeriesChartType.Spline
        Chart2.Series(2).ChartType = DataVisualization.Charting.SeriesChartType.Spline
        '表格處理
        DataGridView3.Rows.Clear()
        DataGridView3.Rows.Add(6)
        DataGridView3.Rows(0).Cells(0).Value = Velocity.ToString + "(cm/s)"
        DataGridView3.Rows(0).Cells(1).Value = "间距(秒)"
        For i = 0 To 15
            DataGridView3.Rows(0).Cells(i + 2).Value = T(i).ToString("0.0")
        Next
        DataGridView3.Rows(1).Cells(1).Value = "#1(℃)"
        DataGridView3.Rows(2).Cells(1).Value = "#2(℃)"
        DataGridView3.Rows(3).Cells(1).Value = "#3(℃)"
        DataGridView3.Rows(4).Cells(1).Value = "#4(℃)"
        DataGridView3.Rows(5).Cells(1).Value = "#5(℃)"
        DescriptionOfDatagridview()
    End Sub
    ''' <summary>
    ''' 将測試数据存放到数据库
    ''' </summary>
    ''' <param name="SavePath"></param>
    ''' <remarks></remarks>
    Public Sub InserToSQL(ByVal Sn As String, ByVal SavePath As String)
        OledbConnection = cn.New_Datacon()
        With OledbCommand
            .Connection = OledbConnection
            .CommandType = CommandType.Text
            .CommandText = "INSERT INTO TestData  (Model, StartTime, EndTime, SN, Velocity, BoardLength, PlateEdges, BaseModel, OffSetModel, SavePath) VALUES(" +
           "'" + ModleName + "'," +
           "'" + t0(OutNum) + "'," +
           "'" + Now.ToString("yyyy/MM/dd HH:mm:ss") + "'," +
           "'" + Sn + "'," +
            Velocity.ToString + "," +
            (PCB_Length * 10).ToString + "," +
            (PCB_Edges * 10).ToString + "," +
            "'" + BaseModle + "'," +
            "'" + OffSetModle + "'," +
           "'" + SavePath + "')"
            .ExecuteNonQuery()
        End With
        OledbCommand.Dispose()
        OledbConnection.Close()
    End Sub
    ''' <summary>
    ''' 自动调整窗口控件
    ''' </summary>
    ''' <param name="inObj"></param>
    ''' <remarks></remarks>
    Public Sub AutoCtlSize(ByVal inObj As Control)     '自动调整控件大小 
        If inObj Is Nothing Then
            Exit Sub
        End If
        '显示分辨率与窗体工作区的大小的关系：分辨率width*height--工作区（没有工具栏）width*(height-46) 
        '即分辨率为*600时,子窗体只能为*554 
        '上述情况还要windows状态栏自动隐藏，如果不隐藏，则height还要减少，结果为：*524 
        '检测桌面显示分辨率(Visual Basic)请参见 
        '此示例以像素为单位确定桌面的宽度和高度。 
        Dim DeskTopSize As Size = System.Windows.Forms.SystemInformation.PrimaryMonitorSize
        Dim FontSize As Single
        DeskTopSize.Width -= DeskTopSize.Width * 0.25
        DeskTopSize.Height -= DeskTopSize.Height * 0.25

        '控件本身**** 
        '控件大小 
        inObj.Size = New Size(Int(inObj.Size.Width * DeskTopSize.Width / 800), Int(inObj.Size.Height * DeskTopSize.Height / 600))
        '控件位置 
        inObj.Location = New Point(Int(inObj.Location.X * DeskTopSize.Width / 800), Int(inObj.Location.Y * DeskTopSize.Height / 600))
        '如果控件为Form,则设置固定边框 
        'Dim mType As Type
        'Dim mProperty As System.Reflection.PropertyInfo
        'mType = inObj.GetType

        'mProperty = mType.GetProperty("FormBorderStyle")
        'If Not mProperty Is Nothing Then
        '    MessageBox.Show(mType.ToString)
        '    mProperty.SetValue(inObj, FormBorderStyle.FixedSingle, Nothing)
        'End If
        '子控件===== 
        Dim n As Integer
        For n = 0 To inObj.Controls.Count - 1
            If inObj.Controls.Item(n).Name = "Label3" Then
                FontSize = 36 * DeskTopSize.Height / 600
            Else
                FontSize = 9 * DeskTopSize.Height / 600
            End If
            '只调整子控件的字体。（如果调整窗体的字体，再调用窗体的show方法时，会引发resize从而导致控件的大小和布局改变） 
            inObj.Controls.Item(n).Font = New Font(inObj.Controls.Item(n).Font.FontFamily, FontSize)
            '递归调用（穷举所有子控件） 
            AutoCtlSize(inObj.Controls.Item(n))
        Next
    End Sub
    ''' <summary>
    '''    啓動
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Test_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        AutoCtlSize(Me)
        FormBorderStyle = Windows.Forms.FormBorderStyle.FixedSingle
        LoadStart()
        DescriptionOfDatagridview()
        '初始化listbox2當前迴流爐鍋爐序號：
        ListBox2.Items.Add("當前迴流爐鍋爐序號：")
        For i = 0 To 20
            ListBox2.Items.Add("")
        Next
        For i = 0 To 15
            Chart1.Series(0).Points.InsertXY(i, "#" + (i + 1).ToString, 150)
        Next
        For i = 0 To 15
            Chart3.Series(0).Points.InsertXY(i, "#" + (i + 1).ToString, 150)
        Next
        '获取所有Baseprofie文件
        Dim PathString As String = Dir(Application.StartupPath & "\Base\", vbDirectory)
        ComboBox2.Items.Clear()
        Do While PathString <> ""
            If PathString <> "." And PathString <> ".." And (PathString.ToUpper.LastIndexOf(".XL") > 0 Or PathString.ToUpper.LastIndexOf(".CSV") > 0) Then
                ComboBox2.Items.Add(PathString)
            End If
            PathString = Dir()
        Loop
        '获取所有Baseprofie文件
        ComboBox5.Items.Clear()
        PathString = Dir(Application.StartupPath & "\Base_Log\", vbDirectory)
        Do While PathString <> ""
            If PathString <> "." And PathString <> ".." And (PathString.ToUpper.LastIndexOf(".XL") > 0 Or PathString.ToUpper.LastIndexOf(".CSV") > 0) Then
                ComboBox5.Items.Add(PathString)
            End If
            PathString = Dir()
        Loop
        '执行主程序
        main()
    End Sub
    ''' <summary>
    ''' 開始採集數據
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub TextBox1_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        '按下回車&條碼長度大於等於3
        If e.KeyValue = 13 And TextBox1.TextLength >= 3 Then
            If ComboBox1.SelectedItem <> "" And ComboBox5.SelectedItem <> "" Then
                InSenser = True
                '記錄當前條碼
                t0(InNum) = Now.ToString("yyyy/MM/dd HH:mm:ss")
                TestSN(InNum) = TextBox1.Text
                TestallNum += 1
                MaxUpValue(InNum) = 0 '最大上升斜率
                MaxDownValue(InNum) = 0 '最小下降斜率
                MaxTestValue(InNum) = 25 '最大測試溫度
                SoakTime(InNum) = 0  '恆溫時間
                RefluxingTime(InNum) = 0  '迴流時間
                MaxValueTime = False '是否達到峯值時間
                Sleep1(300)
                ChartInitialization()
            Else
                MessageBox.Show("請選擇機種與基準文件")
            End If
        End If
    End Sub
    ''' <summary>
    ''' 選擇新機種
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ComboBox1_SelectionChangeCommitted(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted
        Dim enter As String
        enter = MsgBox("确认选择该机种？", MsgBoxStyle.OkCancel)
        If enter = vbOK Then
            ModleName = ComboBox1.SelectedItem
            ComboxSelectItem(0) = ComboBox1.SelectedIndex
            DataInitialization() '數據初始化
            If ComboBox1.SelectedItem <> "" Then
                If WarmingUp = False Then
                    task3.Start()
                End If
                '選擇機種後復位初始化
                For i = 0 To 20
                    TestSN(i) = ""
                Next
                InNum = 0
                OutNum = 0
                TestallNum = 0
                Button4.Enabled = True
            End If
        Else
            ComboBox1.SelectedIndex = ComboxSelectItem(0)
        End If
    End Sub
    ''' <summary>
    ''' 調入參考文件獲取BaseData
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ComboBox2_SelectionChangeCommitted(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectionChangeCommitted
        Dim enter As String
        enter = MsgBox("确认选择该參考文件？", MsgBoxStyle.OkCancel)
        If enter = vbOK Then
            ComboxSelectItem(1) = ComboBox2.SelectedIndex
            BaseData = Ex.GetBaseProFile(Application.StartupPath + "\Base\" + ComboBox2.SelectedItem, ProgressBar1)  '調入參考文件獲取BaseData
            BaseModle = ComboBox2.SelectedItem
        Else
            ComboBox2.SelectedIndex = ComboxSelectItem(1)
        End If
        Button1.Focus()
    End Sub
    ''' <summary>
    ''' 調入補償文件獲取OffSet
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ComboBox5_SelectionChangeCommitted(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox5.SelectionChangeCommitted
        If ComboBox1.SelectedItem <> "" Then
            Dim enter As String
            enter = MsgBox("确认选择该補償文件？", MsgBoxStyle.OkCancel)
            If enter = vbOK Then
                ComboxSelectItem(2) = ComboBox5.SelectedIndex
                Offset = Ex.GetOffSetFile(Application.StartupPath + "\Base_Log\" + ComboBox5.SelectedItem, ProgressBar1) '調入參考文件獲取BaseData
                OffSetModle = ComboBox5.SelectedItem
                '將基準文件寫入Modle機種數據庫,以便下次打開機種時同時打開基準文件
                OledbConnection = cn.New_Datacon()
                With OledbCommand
                    .Connection = OledbConnection
                    .CommandType = CommandType.Text
                    .CommandText = "UPDATE Modle SET [BaseLog] ='" + OffSetModle + "' where Model='" + ModleName + "'"
                    .ExecuteNonQuery()
                End With
                OledbCommand.Dispose()
                OledbConnection.Close()
            Else
                ComboBox5.SelectedIndex = ComboxSelectItem(2)
            End If
            If ComboBox5.SelectedItem <> "" Then
                TextBox1.Enabled = True
            End If
            TextBox1.Focus()
        Else
            ComboBox5.SelectedIndex = ComboxSelectItem(2)
            MessageBox.Show("請先選擇測試機種")
        End If
    End Sub
    ''' <summary>
    ''' 啓動單組採集
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Button1.Text = "開始採集" Then
            If ComboBox1.SelectedItem <> "" Then
                If ComboBox2.SelectedItem <> "" Then
                    Try
                        EnterString = InputBox("請輸入需要生成的基準文件名：")
                        If EnterString <> "" Then
                            '復位所有參數
                            ReDim Offset(1500, 4)
                            ReDim TestData(20, 4, 1500)
                            Button1.Text = "停止採集"
                            ChartInitialization() '曲線圖像初始化
                            TestSN(20) = EnterString
                        End If
                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try
                Else
                    MessageBox.Show("請選擇參考文件！")
                End If
            Else
                MessageBox.Show("請選擇需要添加TEST文件的機種！")
            End If
        Else
            If TestSN(20) <> "" Then
                StopGetData = True
                While StopGetData
                    Sleep1(10)
                End While
            End If
            Button1.Text = "開始採集"
        End If
    End Sub
    ''' <summary>
    ''' 獲取當前爐溫值
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim i As Integer
        Dim DataAString(15), DataBString(15) As Integer '將數據轉換爲字符串
        Try
            '若已進行採集則直接提取數據，否則進行直接採集
            If task3.IsAlive Then
                For i = 0 To 15
                    DataAString(i) = DataA_Value(i).ToString("0.0")
                    DataBString(i) = DataB_Value(i).ToString("0.0")
                Next
            Else
                DataA_Value = Get_Data(TemperatureA, 1, 16)
                'DataB_Value = Get_Data(TemperatureB, 1, 16)
                For i = 0 To 15
                    DataAString(i) = DataA_Value(i).ToString("0.0")
                    DataBString(i) = DataB_Value(i).ToString("0.0")
                Next
            End If
            Chart1.Series(0).Points.DataBindY(DataAString)
            Chart3.Series(0).Points.DataBindY(DataBString)
            For i = 1 To 16
                DataGridView6.Rows(1).Cells(i).Value = DataA_Value(i - 1) - DataGridView6.Rows(0).Cells(i).Value
                DataGridView6.Rows(3).Cells(i).Value = DataA_Value(i - 1) - DataGridView6.Rows(2).Cells(i).Value
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message + vbCrLf + "通信異常！")
        End Try
    End Sub
    ''' <summary>
    ''' 添加新機種
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        NewModle.ShowDialog()
    End Sub
    ''' <summary>
    ''' 修改參數
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Edit.ShowDialog()
    End Sub
    ''' <summary>
    ''' 雙擊打開數據PDF
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub DataGridView2_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView2.MouseDoubleClick
        Dim Topath, Model, SavePath, SN As String
        Dim RowsIndex As Integer
        Topath = Application.StartupPath + "\PDFLog\"
        RowsIndex = DataGridView2.CurrentRow.Index
        Model = DataGridView2.Rows(RowsIndex).Cells("Model").Value.ToString
        SN = DataGridView2.Rows(RowsIndex).Cells("SN").Value.ToString
        SavePath = Application.StartupPath + DataGridView2.Rows(RowsIndex).Cells("SavePath").Value.ToString
        If My.Computer.FileSystem.FileExists(SavePath) Then
            If My.Computer.FileSystem.FileExists(Topath + SN + ".pdf") = False Then
                Ex.SavePDF(SavePath, Topath + SN + ".pdf")
            End If
            System.Diagnostics.Process.Start(Topath + SN + ".pdf")
        Else
            MessageBox.Show("档案丢失")
        End If
    End Sub
    ''' <summary>
    ''' 刷新数据库
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '獲取所有已有機種
        ComboBox4.Items.Clear()
        ComboBox4.ResetText()
        ComboBox4.Items.Add("ALL")
        OledbConnection = cn.New_Datacon()
        With OledbCommand
            .Connection = OledbConnection
            .CommandType = CommandType.Text
            .CommandText = "SELECT   distinct [Model] FROM ([TestData])"
        End With
        Dim OledbReader As OleDbDataReader
        OledbReader = OledbCommand.ExecuteReader
        While OledbReader.Read
            ComboBox4.Items.Add(OledbReader.GetString(0))
        End While
        Dim dt As DataSet = cn.Getds(OledbConnection, "Select * from [TestData]")
        DataGridView2.DataSource = dt.Tables(0)
        OledbReader.Close()
        OledbConnection.Close()
    End Sub
    ''' <summary>
    ''' 多連板測試
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Label15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label15.Click
        If MultiplePCB.IsHandleCreated Then
            MultiplePCB.Close()
        Else
            MultiplePCB.Show()
        End If
    End Sub
    ''' <summary>
    ''' 刪除數據庫資料
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If DataGridView2.Rows.Count > 0 Then
            Dim Model, StartTime, SN, SavePath, Topath, ComText As String
            Dim RowsIndex As Integer
            Try
                Topath = Application.StartupPath + "\PDFLog\"
                RowsIndex = DataGridView2.CurrentRow.Index
                Model = DataGridView2.Rows(RowsIndex).Cells("Model").Value.ToString
                SN = DataGridView2.Rows(RowsIndex).Cells("SN").Value.ToString
                StartTime = DataGridView2.Rows(RowsIndex).Cells("StartTime").Value.ToString
                SavePath = DataGridView2.Rows(RowsIndex).Cells("SavePath").Value.ToString
                '以下爲刪除數據庫資料
                ComText = "DELETE FROM TestData where model='" + Model + "' and SN='" + SN + "' and SavePath='" + SavePath + "'"
                OledbConnection = cn.New_Datacon()
                With OledbCommand
                    .Connection = OledbConnection
                    .CommandType = CommandType.Text
                    .CommandText = ComText
                    .ExecuteNonQuery()
                End With
                '刷新數據庫
                Dim dt As DataSet = cn.Getds(OledbConnection, "Select * from [TestData]")
                DataGridView2.DataSource = dt.Tables(0)
                OledbConnection.Close()
                SavePath = Application.StartupPath + SavePath
                '以下爲刪除原始數據TestLog&PdfLog
                'TestLOG
                If My.Computer.FileSystem.FileExists(SavePath) Then
                    My.Computer.FileSystem.DeleteFile(SavePath)
                End If
                'PDFLOG
                If My.Computer.FileSystem.FileExists(Topath + SN + ".pdf") Then
                    My.Computer.FileSystem.DeleteFile(Topath + SN + ".pdf")
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        Else
            MessageBox.Show("請選擇需要刪除的記錄")
        End If
    End Sub
    ''' <summary>
    ''' 執行查詢
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim dt As DataSet
        OledbConnection = cn.New_Datacon()
        If RadioButton1.Checked And ComboBox4.SelectedIndex <= 0 Then
            dt = cn.Getds(OledbConnection, "Select * from [TestData]")
        ElseIf RadioButton1.Checked And ComboBox4.SelectedIndex > 0 Then
            dt = cn.Getds(OledbConnection, "Select * from [TestData] Where Model='" + ComboBox4.SelectedItem + "'")
        ElseIf RadioButton2.Checked And ComboBox4.SelectedIndex <= 0 Then
            dt = cn.Getds(OledbConnection, "Select * from [TestData] Where StartTime>'" + DateTimePicker2.Text + "' and EndTime <='" + DateTimePicker1.Text + "'")
        ElseIf RadioButton2.Checked And ComboBox4.SelectedIndex > 0 Then
            dt = cn.Getds(OledbConnection, "Select * from [TestData] Where Model='" + ComboBox4.SelectedItem + "' and StartTime>'" + DateTimePicker2.Text + "' and EndTime <='" + DateTimePicker1.Text + "'")
        Else
            dt = cn.Getds(OledbConnection, "Select * from [TestData]")
        End If
        DataGridView2.DataSource = dt.Tables(0)
        OledbConnection.Close()
    End Sub
    ''' <summary>
    ''' 模擬出板傳感器
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        OutSenser = True
    End Sub
    'Private Sub DataGridView1_CellPainting(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs) Handles DataGridView1.CellPainting
    '    '纵向合并 
    '    For Each fieldHeaderText In colsHeaderText_H
    '        If (e.ColumnIndex >= 0 And DataGridView1.Columns(e.ColumnIndex).HeaderText = fieldHeaderText And e.RowIndex >= 0) Then
    '            Using gridBrush As Brush = New SolidBrush(DataGridView1.GridColor),
    '                backColorBrush As Brush = New SolidBrush(e.CellStyle.BackColor)
    '            Using gridLinepen As Pen = New Pen(gridBrush)
    '                ' 擦除原单元格背景 
    '                e.Graphics.FillRectangle(backColorBrush, e.CellBounds)
    '                '不是最后一行且单元格的值不为null 
    '                    If (e.RowIndex < DataGridView1.RowCount - 1 And DataGridView1.Rows(e.RowIndex + 1).Cells(e.ColumnIndex).Value <> Nothing) Then
    '                        '若与下一单元格值不同 
    '                        If (e.Value.ToString() <> DataGridView1.Rows(e.RowIndex + 1).Cells(e.ColumnIndex).Value.ToString()) Then
    '                            '下边缘的线 
    '                            e.Graphics.DrawLine(gridLinepen, e.CellBounds.Left, e.CellBounds.Bottom - 1,
    '                            e.CellBounds.Right - 1, e.CellBounds.Bottom - 1)
    '                            '绘制值 
    '                            If (e.Value <> Nothing) Then
    '                                e.Graphics.DrawString(e.Value.ToString(), e.CellStyle.Font,
    '                                    Brushes.Crimson, e.CellBounds.X + 2,
    '                                    e.CellBounds.Y + 2, StringFormat.GenericDefault)
    '                            End If
    '                            '若与下一单元格值相同   
    '                        Else
    '                            '背景颜色 
    '                            'e.CellStyle.BackColor = Color.LightPink;   //仅在CellFormatting方法中可用 
    '                            DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.BackColor = Color.LightBlue
    '                            DataGridView1.Rows(e.RowIndex + 1).Cells(e.ColumnIndex).Style.BackColor = Color.LightBlue
    '                            '只读（以免双击单元格时显示值） 
    '                            DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).ReadOnly = True
    '                            DataGridView1.Rows(e.RowIndex + 1).Cells(e.ColumnIndex).ReadOnly = True
    '                        End If
    '                        '最后一行或单元格的值为null 
    '                    Else
    '                        '下边缘的线 
    '                        e.Graphics.DrawLine(gridLinepen, e.CellBounds.Left, e.CellBounds.Bottom - 1,
    '                        e.CellBounds.Right - 1, e.CellBounds.Bottom - 1)

    '                        '绘制值 
    '                        If (e.Value <> Nothing) Then
    '                            e.Graphics.DrawString(e.Value.ToString(), e.CellStyle.Font,
    '                                        Brushes.Crimson, e.CellBounds.X + 2,
    '                                        e.CellBounds.Y + 2, StringFormat.GenericDefault)
    '                        End If
    '                    End If
    '                    '右侧的线 
    '                    e.Graphics.DrawLine(gridLinepen, e.CellBounds.Right - 1,
    '                        e.CellBounds.Top, e.CellBounds.Right - 1,
    '                        e.CellBounds.Bottom - 1)
    '                    '设置处理事件完成（关键点），只有设置为ture,才能显示出想要的结果。 
    '                    e.Handled = True
    '                End Using
    '            End Using
    '        End If
    '    Next
    '    For Each fieldHeaderText In colsHeaderText_V
    '        '橫向合并
    '        If (e.ColumnIndex >= 0 And DataGridView1.Columns(e.ColumnIndex).HeaderText = fieldHeaderText And e.RowIndex >= 0) Then
    '            Using gridBrush As Brush = New SolidBrush(DataGridView1.GridColor),
    '                 backColorBrush As Brush = New SolidBrush(e.CellStyle.BackColor)
    '                Using gridLinepen As Pen = New Pen(gridBrush)
    '                    ' 擦除原单元格背景 
    '                    e.Graphics.FillRectangle(backColorBrush, e.CellBounds)
    '                    '不是最后一行且单元格的值不为null 
    '                    If (e.ColumnIndex < DataGridView1.ColumnCount - 1 And DataGridView1.Rows(e.RowIndex + 1).Cells(e.ColumnIndex + 1).Value <> Nothing) Then
    '                        '若与下一单元格值不同 
    '                        If (e.Value.ToString() <> DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex + 1).Value.ToString()) Then
    '                            '下边缘的线 
    '                            e.Graphics.DrawLine(gridLinepen, e.CellBounds.Right - 1, e.CellBounds.Top,
    '                            e.CellBounds.Right - 1, e.CellBounds.Bottom - 1)
    '                            '绘制值 
    '                            If (e.Value <> Nothing) Then
    '                                e.Graphics.DrawString(e.Value.ToString(), e.CellStyle.Font,
    '                                    Brushes.Crimson, e.CellBounds.X + 2,
    '                                    e.CellBounds.Y + 2, StringFormat.GenericDefault)
    '                            End If
    '                            '若与下一单元格值相同   
    '                        Else
    '                            '背景颜色 
    '                            'e.CellStyle.BackColor = Color.LightPink;   //仅在CellFormatting方法中可用 
    '                            DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.BackColor = Color.LightPink
    '                            DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex + 1).Style.BackColor = Color.LightPink
    '                            '只读（以免双击单元格时显示值） 
    '                            DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).ReadOnly = True
    '                            DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex + 1).ReadOnly = True
    '                        End If
    '                        '最后一行或单元格的值为null 
    '                    Else

    '                        '右侧的线 
    '                        e.Graphics.DrawLine(gridLinepen, e.CellBounds.Right - 1,
    '                            e.CellBounds.Top, e.CellBounds.Right - 1,
    '                            e.CellBounds.Bottom - 1)
    '                        '绘制值 
    '                        If (e.Value <> Nothing) Then
    '                            e.Graphics.DrawString(e.Value.ToString(), e.CellStyle.Font,
    '                                        Brushes.Crimson, e.CellBounds.X + 2,
    '                                        e.CellBounds.Y + 2, StringFormat.GenericDefault)
    '                        End If
    '                    End If
    '                    '下边缘的线 
    '                    e.Graphics.DrawLine(gridLinepen, e.CellBounds.Left, e.CellBounds.Bottom - 1,
    '                        e.CellBounds.Right - 1, e.CellBounds.Bottom - 1)
    '                    '设置处理事件完成（关键点），只有设置为ture,才能显示出想要的结果。 
    '                    e.Handled = True
    '                End Using
    '            End Using
    '        End If
    '    Next
    'End Sub
    ''' <summary>
    ''' 輸入生產地點
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub TextBox2_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyValue = 13 And TextBox2.Text.Length > 0 Then
            TextBox2.Enabled = False
            ProductionSite = TextBox2.Text
        End If
    End Sub
    ''' <summary>
    ''' 修改生成地點
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        TextBox2.Enabled = True
        TextBox2.Focus()
    End Sub
    ''' <summary>
    ''' 點擊取消
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub TextBox2_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TextBox2.MouseClick
        TextBox2.Text = ""
    End Sub
    Public Function Cpk(ByVal DataCpk() As Single, ByVal Updata As Single, ByVal DownData As Single) As Single
        Dim CpkValue As Single = 0
        Dim Avg As Single
        Dim BiaoZhuncha As Single
        '計算標準差
        If DataCpk.Count > 0 Then
            Avg = DataCpk.Average
            Dim datad(DataCpk.Count - 1) As Single
            For i = 0 To DataCpk.Count - 1
                datad(i) = (DataCpk(i) - Avg) * (DataCpk(i) - Avg)
            Next
            Dim Sum1 As Single = datad.Sum()
            BiaoZhuncha = Math.Sqrt(Sum1 / (DataCpk.Count - 1))
            '計算Cpu
            Dim Cpu = (Updata - Avg) / (3 * BiaoZhuncha)
            '計算Cp1
            Dim Cp1 = (Avg - DownData) / (3 * BiaoZhuncha)
            CpkValue = Min(Cpu, Cp1)
            Return CpkValue
        End If
        Return 0
    End Function
End Class
