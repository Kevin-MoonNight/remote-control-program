Imports System.Net
Imports System.Net.Sockets
Imports System.Text
Imports System.Threading
Imports System.IO
Public Class Form1
    '主機端
    Dim picture_transmission_listen As TcpListener '畫面傳輸socket
    Dim control_command_client As New TcpClient '鍵鼠控制socket
    Dim client_socket As Socket
    Delegate Sub button_text_change(ByVal str As String) '委派 更改按鈕文字
    Delegate Sub label_text_change(ByVal str As String) '委派 更改標籤文字
    Delegate Sub timer_start(ByVal t As Boolean) '委派 開關計時器
    Private Sub change_label(ByVal str As String)
        Label1.Text = str
    End Sub '更改 標籤文字
    Private Sub change_button(ByVal str As String)
        Button1.Text = str
    End Sub '更改 按鈕文字
    Private Sub picture_start(ByVal t As Boolean)
        If t = True Then
            picture.Start()
        Else
            picture.Stop()
        End If
    End Sub '開關 計時器傳輸畫面
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If sender.text = "開始" Then
            '連接 socket
            '背景執行
            Dim connect As New Thread(AddressOf connect_socket) '監聽
            connect.IsBackground = True : connect.Start() '開始執行
        ElseIf sender.text = "斷開連接" Then
            disconnect()
        End If
    End Sub
    Sub connect_socket()
        '監聽 畫面傳輸socket
        listen_picture_transmission()
        '連接 鍵鼠控制socket
        connect_control_command()
        '更新畫面
        change_controls_text("連線成功!", "斷開連接")
        '開始傳送圖片
        Me.Invoke(New timer_start(AddressOf picture_start), New Object() {True}) '開啟 計時器傳送畫面
        '接收 控制命令
        '背景執行
        Dim receive_control As New Thread(AddressOf receive_control_command)
        receive_control.IsBackground = True : receive_control.Start()
    End Sub '連接 socket

    Sub setip()
        Dim ip As String = Dns.GetHostEntry(Dns.GetHostName).AddressList(1).ToString '找出自己ip
        picture_transmission_listen = New TcpListener(IPAddress.Parse(ip), 4444) '設定ip、port
        '更新畫面
        change_controls_text("本機IP:" & ip & vbNewLine & "等待連線...", "等待連線")
    End Sub '設定ip
    Sub listen_picture_transmission()
        '設定ip 
        setip()
        picture_transmission_listen.Start()  '開始監聽
        client_socket = picture_transmission_listen.AcceptSocket '回傳Socket與新進連接的用戶端來通訊
    End Sub '監聽 畫面傳輸socket
    Sub connect_control_command()
        control_command_client = New TcpClient
        control_command_client.Connect(Split(client_socket.RemoteEndPoint.ToString, ":")(0), 4444)
    End Sub '連接 鍵鼠控制socket

    Private Sub picture_tick(sender As Object, e As EventArgs) Handles picture.Tick '每16毫秒傳送一張畫面(60幀)
        Try '螢幕截圖轉成Byte寫入資料流
            client_socket.Send(bitmap_to_byte(screenshot))
        Catch ex As Exception
            disconnect()
        End Try
    End Sub '傳輸 螢幕畫面
    Function screenshot() As Image
        '創造一個跟螢幕解析度一樣的畫布
        Dim primaryscreen As New Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height)
        Dim g As Graphics = Graphics.FromImage(primaryscreen)
        '將螢幕畫面截圖放到畫布中
        g.CopyFromScreen(0, 0, 0, 0, New Drawing.Size(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height))
        '再創造一個輸出用的畫布
        Dim output_picture As New Bitmap(1280, 720)
        g = Graphics.FromImage(output_picture)
        '將截圖的畫面已指定縮放到指定大小
        g.DrawImage(primaryscreen, 0, 0, 1280, 720)
        Return output_picture
    End Function '螢幕截圖
    Function bitmap_to_byte(ByVal primaryscreen As Bitmap) As Byte()
        '將圖片放到資料流
        Dim picture_stream As New MemoryStream
        primaryscreen.Save(picture_stream, Imaging.ImageFormat.Bmp)
        '將資料流的資料轉成Byte
        Dim picture_byte(picture_stream.Length) As Byte
        picture_byte = picture_stream.ToArray
        Return picture_byte
    End Function '圖片轉成Byte


    Sub receive_control_command()
        Dim networkstream As NetworkStream = control_command_client.GetStream
        Do
            Threading.Thread.Sleep(0)
            If networkstream.CanRead = True Then '有資料可以讀取
                Try
                    '判斷控制命令文字 執行控制動作
                    judgment_control(Strings.Replace(byte_to_string(read_control_data(networkstream)), vbNullChar, ""))
                Catch ex As Exception
                    If client_socket.Connected = False Then
                        Exit Do
                    End If
                End Try
            End If
        Loop
    End Sub '接收 控制命令
    Function read_control_data(ByVal networkstream As NetworkStream) As Byte()
        Dim control_byte(50) As Byte
        networkstream.Read(control_byte, 0, 50)
        Return control_byte
    End Function '讀取 控制命令
    Function byte_to_string(ByVal control_byte() As Byte) As String
        Return Encoding.GetEncoding(950).GetString(control_byte)
    End Function 'byte轉回string格式

    Sub judgment_control(ByVal str As String)
        Dim control_command() As String = Split(str, " ")
        Select Case control_command(0)
            Case "move" '如果是滑鼠移動
                move_mouse((Val(control_command(1)) * Screen.PrimaryScreen.Bounds.Width).ToString("0"), (Val(control_command(2)) * Screen.PrimaryScreen.Bounds.Height).ToString("0"))
            Case "mouse" '如果是滑鼠類的控制
                mouse(Val(control_command(1)), control_command(2))
            Case "wheel"
                wheel(Val(control_command(1)))
            Case "keyboard" '如果是鍵盤類的控制
                keyboard(Val(control_command(1)), control_command(2))
        End Select
    End Sub '判斷 控制命令
    Sub mouse(ByVal control_command1 As Integer, ByVal control_command2 As String)
        Select Case control_command1 '判斷是按下哪個鍵
            Case MouseButtons.Left '如果是左鍵
                If control_command2 = "down" Then
                    left_down()
                ElseIf control_command2 = "up" Then
                    left_up()
                End If
            Case MouseButtons.Right '如果是右鍵
                If control_command2 = "down" Then
                    right_down()
                ElseIf control_command2 = "up" Then
                    right_up()
                End If
            Case MouseButtons.Middle '如果是中鍵 
                If control_command2 = "down" Then
                    middle_down()
                ElseIf control_command2 = "up" Then
                    middle_up()
                End If
        End Select
    End Sub '執行滑鼠動作
    Sub keyboard(ByVal control_command1 As Integer, ByVal control_command2 As String)
        If control_command2 = "down" Then '如果是按下
            keyboard_down(control_command1)
        ElseIf control_command2 = "up" Then '如果是放開
            keyboard_up(control_command1)
        End If
    End Sub '執行鍵盤動作

    Sub change_controls_text(ByVal labelstr As String, ByVal buttonstr As String)
        Me.Invoke(New label_text_change(AddressOf change_label), New Object() {labelstr}) '委派更改標籤文字 
        Me.Invoke(New button_text_change(AddressOf change_button), New Object() {buttonstr}) '委派更改按鈕文字
    End Sub '更新畫面
    Sub disconnect()
        '初始化
        client_socket.Close()
        picture_transmission_listen.Stop()
        control_command_client.Close() : control_command_client = New TcpClient
        '更新畫面
        change_controls_text("", "開始")
        '開關計時器傳送畫面
        Me.Invoke(New timer_start(AddressOf picture_start), New Object() {False})
    End Sub '中斷連線
End Class