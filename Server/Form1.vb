Imports System.Net
Imports System.Net.Sockets
Imports System.Threading
Imports System.IO
Public Class Form1
    '主機端
    Dim picture_transmission_listen As TcpListener '畫面傳輸socket
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
        '更新畫面
        change_controls_text("連線成功!", "斷開連接")
        '開始傳送圖片
        Me.Invoke(New timer_start(AddressOf picture_start), New Object() {True}) '開啟 計時器傳送畫面
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


    Sub change_controls_text(ByVal labelstr As String, ByVal buttonstr As String)
        Me.Invoke(New label_text_change(AddressOf change_label), New Object() {labelstr}) '委派更改標籤文字 
        Me.Invoke(New button_text_change(AddressOf change_button), New Object() {buttonstr}) '委派更改按鈕文字
    End Sub '更新畫面
    Sub disconnect()
        '初始化
        client_socket.Close()
        picture_transmission_listen.Stop()
        '更新畫面
        change_controls_text("", "開始")
        '開關計時器傳送畫面
        Me.Invoke(New timer_start(AddressOf picture_start), New Object() {False})
    End Sub '中斷連線
End Class