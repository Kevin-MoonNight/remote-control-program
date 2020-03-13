Imports System.Net
Imports System.Net.Sockets
Imports System.Text
Imports System.Threading
Imports System.IO
Public Class Form1
    '控制端
    Dim picture_transmission_client As New TcpClient '畫面傳輸socket
    Dim control_command_listener As TcpListener '鍵鼠控制socket
    Dim client_socket As Socket
    Delegate Sub picturebox_image_change(ByVal pic As Bitmap) '委派更改圖片
    Delegate Sub button_text_change(ByVal Str As String) '委派更改按鈕文字
    Delegate Sub textbox_text_change(ByVal T As Boolean) '委派更改可否輸入文字
    Sub change_button(ByVal Str As String)
        Button1.Text = Str
    End Sub '委派更改按鈕文字
    Sub change_textbox(ByVal T As Boolean)
        TextBox1.Enabled = True
    End Sub '委派更改可否輸入文字
    Sub change_image(ByVal pic As Bitmap)
        PictureBox1.Image = pic
    End Sub '委派更改圖片

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Button1.Text = "連接" Then
            connect_socket(TextBox1.Text, 4444)
        Else
            disconnect()
        End If
    End Sub '連接Socket
    Sub connect_socket(ByVal ip As String, ByVal port As Integer)
        Try '嘗試連接
            '連接 畫面傳輸Socket
            connect_picture_transmission(ip, port)
            '監聽 鍵鼠控制Socket
            listen_control_command(port)
            '接收傳送過來的畫面
            receive_picture_transmission()
            '禁止更改ip
            TextBox1.Enabled = False : Button1.Text = "斷開連接"
        Catch ex As Exception '如果連接失敗
            MsgBox("無法連線到此IP位置")
        End Try
    End Sub '連接

    Sub connect_picture_transmission(ByVal ip As String, ByVal port As Integer)
        picture_transmission_client.Connect(ip, port) '連接
    End Sub '連接 畫面傳輸Socket
    Sub listen_control_command(ByVal port As Integer)
        '用於傳輸滑鼠按鍵資料
        Dim ip As String = Dns.GetHostEntry(Dns.GetHostName).AddressList(1).ToString '查看自己的ip
        control_command_listener = New TcpListener(IPAddress.Parse(ip), port) '設定IP、Port
        control_command_listener.Start() '開始接聽
        client_socket = control_command_listener.AcceptSocket() '回傳Socket與新進連接的用戶端來通訊
    End Sub '監聽 鍵鼠控制Socket

    Sub receive_picture_transmission()
        '將接收圖片設為背景執行
        Dim receive_picture As New Thread(AddressOf read_picture_data)
        receive_picture.IsBackground = True : receive_picture.Start()
    End Sub '接收傳送過來的畫面
    Sub read_picture_data()
        Dim networkstream As NetworkStream = picture_transmission_client.GetStream
        Dim pic As Bitmap
        Do
            Threading.Thread.Sleep(0)
            '如果資料流可以讀取
            If NetWorkStream.CanRead = True Then
                Try
                    '將資料轉成Image
                    pic = byte_to_image(NetWorkStream, pic)
                    '輸出畫面 
                    Me.Invoke(New picturebox_image_change(AddressOf change_image), New Object() {pic}) '委派更改 圖片
                Catch ex As Exception
                End Try
            Else
                PictureBox1.Image = Nothing : Exit Do
            End If
        Loop
    End Sub '讀取資料流裡的畫面資料
    Function byte_to_image(ByVal networkstream As NetworkStream, ByVal pic As Bitmap) As Bitmap
        Dim newpic As Bitmap
        Dim picbyte(3686454) As Byte
        '從資料流上面讀取圖片的Byte 
        networkstream.Read(picbyte, 0, 3686454)
        '將Byte放到資料流裡面
        Dim picstream As New MemoryStream(picbyte)
        Try
            newpic = New Bitmap(Image.FromStream(picstream)) '把資料流的資料轉成圖片
        Catch ex As Exception
            Return pic
        End Try
        Return newpic
    End Function '將圖片從Byte轉成Image

    Sub sendkey(ByVal Str As String)
        If picture_transmission_client.Connected = True Then '如果有連線
            '將控制動作變成文字在編碼
            Dim control_byte() As Byte = Encoding.GetEncoding(950).GetBytes(Str)
            '將按鍵編碼寫入資料流
            Try
                Dim output_byte(50) As Byte
                For i = 0 To control_byte.Length - 1
                    output_byte(i) = control_byte(i)
                Next
                client_socket.Send(output_byte)
            Catch ex As Exception '如果連線中斷
                disconnect()
            End Try
        End If
    End Sub '傳送按鍵滑鼠控制

    Sub disconnect()
        '初始化
        picture_transmission_client.Close() : picture_transmission_client = New TcpClient
        control_command_listener.Stop()
        client_socket.Close()
        AcceptButton = Button1
        Me.Invoke(New textbox_text_change(AddressOf change_textbox), New Object() {True}) '委派設定 ip欄可否更改
        Me.Invoke(New button_text_change(AddressOf change_button), New Object() {"連接"}) '委派更改 按鈕文字
        Me.Invoke(New picturebox_image_change(AddressOf change_image), New Object() {Nothing}) '委派更改 圖片
    End Sub '中斷連線

    Private Sub PictureBox1_MouseDown(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseDown
        sendkey("mouse " & e.Button & " down") '進入事件
    End Sub  '如果滑鼠按下
    Private Sub PictureBox1_MouseUp(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseUp
        sendkey("mouse " & e.Button & " up") '進入事件
    End Sub '如果滑鼠放開
    Private Sub PictureBox1_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseMove
        sendkey("move " & (e.X / 1280) & " " & (e.Y / 720))
    End Sub '如果滑鼠移動
    Private Sub PictureBox1_MouseWheel(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseWheel
        sendkey("wheel " & e.Delta) '進入事件
    End Sub  '如果滑鼠滾輪移動
    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        sendkey("keyboard " & e.KeyCode & " down") '進入事件
    End Sub  '如果鍵盤按下
    Private Sub Form1_KeyUp(sender As Object, e As KeyEventArgs) Handles MyBase.KeyUp
        sendkey("keyboard " & e.KeyCode & " up") '進入事件
    End Sub '如果鍵盤放開 
    Private Sub Button1_Focus() Handles Button1.GotFocus
        Label1.Focus()
    End Sub '控制焦點
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.PictureBox1.Select()
    End Sub '讓PictureBox可以觸發按鍵事件
End Class







