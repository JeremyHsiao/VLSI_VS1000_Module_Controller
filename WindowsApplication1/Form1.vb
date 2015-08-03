Imports System
Imports System.IO.Ports
Imports System.Threading

Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        InitRS232()
        ListBox1.Text = "AUDIO001"
        ListBox2.Text = "001"
    End Sub

    Private Sub Form1_Closing(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.FormClosing
        StopRS232()
    End Sub

    Private Sub Delay_loop(ByVal millisec As UInt32)
        While (millisec > 0)
            Thread.Sleep(1)
            Application.DoEvents()
            millisec -= 1
        End While
    End Sub

    Private Sub TextWindowWriteLine(ByVal show_text As String)
        Application.DoEvents()
        RichTextBox1.AppendText(show_text + Chr(13) + Chr(10))
        RichTextBox1.ScrollToCaret()
        RichTextBox2.AppendText(show_text + Chr(13) + Chr(10))
        RichTextBox2.ScrollToCaret()
        Application.DoEvents()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If (comport.IsOpen = False) Then
            TextWindowWriteLine("Connecting...Please Wait.")
            StartRS232()
            GroupBox1.Enabled = True
            GroupBox2.Enabled = True
            TextWindowWriteLine("RS232 Connected.")
        Else
            TextWindowWriteLine("Disconnecting...Please Wait.")
            GroupBox1.Enabled = False
            GroupBox2.Enabled = False
            StopRS232()
            TextWindowWriteLine("RS232 Disconnected.")
        End If

    End Sub

    Private comport As SerialPort

    Private Sub InitRS232()
        comport = New SerialPort()
        comport.PortName = "COM14"
        comport.BaudRate = 115200
        comport.DataBits = 8
        comport.Parity = Parity.None
        comport.StopBits = StopBits.One
        comport.Handshake = Handshake.None
    End Sub

    Private t As Thread

    Private Sub StartRS232()
        If (comport.IsOpen = False) Then
            comport.Open()
            receiving = True
            t = New Thread(AddressOf DoReceive)
            t.IsBackground = True
            t.Start()
        Else
            comport.DiscardInBuffer()
            comport.DiscardOutBuffer()
        End If
    End Sub

    Private Sub StopRS232()
        If (comport.IsOpen = True) Then
            receiving = False
            comport.DiscardInBuffer()
            comport.DiscardOutBuffer()
            comport.Close()
            Delay_loop(500)
        End If
    End Sub

    Private Function WriteRS232Byte(ByVal tx_byte As Byte) As Char
        Dim temp_str As String = Chr(tx_byte)
        Dim temp_hex
        RichTextBox1.AppendText(temp_str)
        RichTextBox1.ScrollToCaret()

        ' Show leading "0" for value <0x10 in HEX WINDOW
        If (tx_byte < 16) Then    ' Add leading '0' for one-digit hex value
            temp_hex = "0" + Hex(tx_byte)
        Else
            temp_hex = Hex(tx_byte)
        End If
        RichTextBox2.AppendText(temp_hex + " ")
        RichTextBox2.ScrollToCaret()
        comport.Write(temp_str)
        Application.DoEvents()
    End Function

    Private Delegate Sub Display(ByVal buffer As Byte())
    Private Sub DisplayBufferHex(ByVal buffer As Byte())
        For index As Int32 = 0 To buffer.Length - 1
            Dim temp_str
            Dim temp_hex
            Dim temp_byte

            temp_byte = buffer(index)
            read_data.Enqueue(temp_byte)

            ' Show either char or "." for non-char in TEXT WINDOW
            If ((temp_byte >= 32) And (temp_byte < 128)) Then
                temp_str = Chr(temp_byte)
            Else
                temp_str = "."
            End If
            RichTextBox1.AppendText(temp_str)

            ' Show " " for 2nd hex value and later
            If (index > 0) Then     ' Add space to separate hex
                temp_hex = " "
            Else
                temp_hex = ""
            End If

            ' Show leading "0" for value <0x10 in HEX WINDOW
            If (temp_byte < 16) Then    ' Add leading '0' for one-digit hex value
                temp_hex += "0"
            End If
            temp_hex = temp_hex + Hex(temp_byte)
            RichTextBox2.AppendText(temp_hex)
        Next index

        RichTextBox1.AppendText(Chr(13) + Chr(10))
        RichTextBox1.ScrollToCaret()
        RichTextBox2.AppendText(Chr(13) + Chr(10))
        RichTextBox2.ScrollToCaret()
    End Sub

    Private read_data As Queue(Of Byte) = New Queue(Of Byte)
    Private receiving As Boolean
    Private Sub DoReceive()
        Dim buffer(1023) As Byte
        While receiving = True
            If comport.BytesToRead > 0 Then
                Dim length As Int32 = comport.Read(buffer, 0, buffer.Length)
                Array.Resize(buffer, length)

                'For temp_index As Integer = 0 To length - 1
                '    read_data.Enqueue(buffer(temp_index))
                'Next temp_index

                Dim d As New Display(AddressOf DisplayBufferHex)
                Me.Invoke(d, New Object() {buffer})
                Array.Resize(buffer, 1024)
            End If
            Delay_loop(16)
        End While
    End Sub

    Private Dont_care_length As Int16 = 20000
    Private Default_wait_time As Integer = 200
    ' Cmd as char version
    Private Sub SendWaitAck(ByVal cmd_str As Byte, ByRef ret_str As String, _
                            ByVal len As Int16, ByVal wait_time As Integer)
        WriteRS232Byte(cmd_str)

        Dim elpased_time As Integer = 0

        If (len = Dont_care_length) Then            ' Dont care, dont wait
            While ((elpased_time < wait_time))
                elpased_time += 1
            End While
        Else
            While ((read_data.Count < len) AndAlso (elpased_time < wait_time))
                elpased_time += 1
                Delay_loop(1)
            End While
        End If

        'ret_str = read_data.ToString()

    End Sub


    Private Sub SendWaitAck(ByVal cmd_str As Char, ByRef ret_str As String, _
                            ByVal len As Int16, ByVal wait_time As Integer)
        WriteRS232Byte(CByte(Asc(cmd_str)))

        Dim elpased_time As Integer = 0

        If (len = Dont_care_length) Then            ' Dont care, dont wait
            While ((elpased_time < wait_time))
                elpased_time += 1
            End While
        Else
            While ((read_data.Count < len) AndAlso (elpased_time < wait_time))
                elpased_time += 1
                Delay_loop(1)
            End While
        End If

        'ret_str = read_data.ToString()

    End Sub

    ' Cmd as string version
    Private Sub SendWaitAck(ByVal cmd_str As String, ByRef ret_str As String, _
                            ByVal len As Int16, ByVal wait_time As Integer)
        For index As Int32 = 0 To cmd_str.Length - 1
            WriteRS232Byte(CByte(Asc(cmd_str(index))))
        Next

        Dim elpased_time As Integer = 0

        If (len = Dont_care_length) Then
            While ((elpased_time < wait_time))
                elpased_time += 1
            End While
        Else
            While ((elpased_time < wait_time) AndAlso (read_data.Count < len))
                elpased_time += 1
                Delay_loop(1)
            End While
        End If

        'ret_str = read_data.ToString()

    End Sub

    '
    ' Command in Continuous mode
    '

    Private Function File_Play_Mode(Optional ByVal with_0x0a As Boolean = False) As String
        ' f - switch to file play mode
        Dim cmd_str As String = "f"
        Dim ret_str As String = ""
        Dim len As Int16 = Dont_care_length
        Dim wait_time As Integer = Default_wait_time

        If (with_0x0a = True) Then
            cmd_str = cmd_str + Chr(10)
        End If

        SendWaitAck(cmd_str, ret_str, len, wait_time)
        Return ret_str
    End Function

    Private Function Cont_Play_Mode(Optional ByVal with_0x0a As Boolean = False) As String
        ' c - switch to continuous play mode
        Dim cmd_str As String = "c"
        Dim ret_str As String = ""
        Dim len As Int16 = Dont_care_length
        Dim wait_time As Integer = Default_wait_time

        If (with_0x0a = True) Then
            cmd_str = cmd_str + Chr(10)
        End If

        SendWaitAck(cmd_str, ret_str, len, wait_time)
        Return ret_str

    End Function

    Private Function Cancel_Play(Optional ByVal with_0x0a As Boolean = False) As String
        ' C - cancel play, return to play loop, responds with c
        Dim cmd_str As String = "C"
        Dim ret_str As String = ""
        Dim len As Int16 = Dont_care_length
        Dim wait_time As Integer = Default_wait_time

        If (with_0x0a = True) Then
            cmd_str = cmd_str + Chr(10)
        End If

        SendWaitAck(cmd_str, ret_str, len, wait_time)
        Return ret_str

    End Function

    Private Function Vol_Up(Optional ByVal with_0x0a As Boolean = False) As String
        ' + - volume up, responds with two-byte current volume level
        Dim cmd_str As String = "+"
        Dim ret_str As String = ""
        Dim len As Int16 = 2
        Dim wait_time As Integer = Default_wait_time

        If (with_0x0a = True) Then
            cmd_str = cmd_str + Chr(10)
        End If

        SendWaitAck(cmd_str, ret_str, len, wait_time)
        Return ret_str

    End Function

    Private Function Vol_Down(Optional ByVal with_0x0a As Boolean = False) As String
        ' - - volume down, responds with two-byte current volume level
        Dim cmd_str As String = "-"
        Dim ret_str As String = ""
        Dim len As Int16 = 2
        Dim wait_time As Integer = Default_wait_time

        If (with_0x0a = True) Then
            cmd_str = cmd_str + Chr(10)
        End If

        SendWaitAck(cmd_str, ret_str, len, wait_time)
        Return ret_str

    End Function

    Private Function Pause_On(Optional ByVal with_0x0a As Boolean = False) As String
        ' = - pause on, responds with =
        Dim cmd_str As String = "="
        Dim ret_str As String = ""
        Dim len As Int16 = 1
        Dim wait_time As Integer = Default_wait_time

        If (with_0x0a = True) Then
            cmd_str = cmd_str + Chr(10)
        End If

        SendWaitAck(cmd_str, ret_str, len, wait_time)
        Return ret_str

    End Function

    Private Function Normal_Play(Optional ByVal with_0x0a As Boolean = False) As String
        ' > - play (normal speed), responds with >
        Dim cmd_str As String = ">"
        Dim ret_str As String = ""
        Dim len As Int16 = 1
        Dim wait_time As Integer = Default_wait_time

        If (with_0x0a = True) Then
            cmd_str = cmd_str + Chr(10)
        End If

        SendWaitAck(cmd_str, ret_str, len, wait_time)
        Return ret_str

    End Function

    Private Function Faster_Play(Optional ByVal with_0x0a As Boolean = False) As String
        ' (0xbb) - faster play, responds with the new play speed
        Dim cmd_str As Byte = 187
        Dim ret_str As String = ""
        Dim len As Int16 = 1
        Dim wait_time As Integer = Default_wait_time

        SendWaitAck(cmd_str, ret_str, len, wait_time)
 
        Return ret_str

    End Function

    Private Function Next_song(Optional ByVal with_0x0a As Boolean = False) As String
        ' n - next song, responds with n
        Dim cmd_str As String = "n"
        Dim ret_str As String = ""
        Dim len As Int16 = 1
        Dim wait_time As Integer = Default_wait_time

        If (with_0x0a = True) Then
            cmd_str = cmd_str + Chr(10)
        End If

        SendWaitAck(cmd_str, ret_str, len, wait_time)
        Return ret_str

    End Function

    Private Function Previous_song(Optional ByVal with_0x0a As Boolean = False) As String
        ' p - previous song, responds with p
        Dim cmd_str As String = "p"
        Dim ret_str As String = ""
        Dim len As Int16 = 1
        Dim wait_time As Integer = Default_wait_time

        If (with_0x0a = True) Then
            cmd_str = cmd_str + Chr(10)
        End If

        SendWaitAck(cmd_str, ret_str, len, wait_time)
        Return ret_str

    End Function

    Private Function Get_Info(Optional ByVal with_0x0a As Boolean = False) As String
        ' p - previous song, responds with p
        Dim cmd_str As String = "?"
        Dim ret_str As String = ""
        Dim len As Int16 = 4 + 1
        Dim wait_time As Integer = Default_wait_time

        If (with_0x0a = True) Then
            cmd_str = cmd_str + Chr(10)
        End If

        SendWaitAck(cmd_str, ret_str, len, wait_time)
        Return ret_str

    End Function

    '
    ' File Mode Command
    '
    '    OFFnn - powers down
    ' cnn - switch to continuous play mode
    ' 
    ' PFILENAMEOGGnn - play by name (capital P), a 8.3-character uppercase name without the ”.” .
    ' 

    Private Function Powers_Down(Optional ByVal with_0x0a As Boolean = True) As String
        ' OFF\n - powers down
        Dim cmd_str As String = "OFF"
        Dim ret_str As String = ""
        Dim len As Int16 = Dont_care_length
        Dim wait_time As Integer = Default_wait_time

        If (with_0x0a = True) Then
            cmd_str = cmd_str + Chr(10)
        End If

        SendWaitAck(cmd_str, ret_str, len, wait_time)
        Return ret_str

    End Function

    Private Function List_Files(Optional ByVal with_0x0a As Boolean = True) As String
        ' L\n - list files
        Dim cmd_str As String = "L"
        Dim ret_str As String = ""
        Dim len As Int16 = Dont_care_length
        Dim wait_time As Integer = 1000

        If (with_0x0a = True) Then
            cmd_str = cmd_str + Chr(10)
        End If

        SendWaitAck(cmd_str, ret_str, len, wait_time)
        Return ret_str

    End Function

    Private Function Play_Filename(ByVal file_name As String, _
                                   Optional ByVal with_0x0a As Boolean = True) As String
        ' PFILENAMEOGGnn - play by name (capital P), a 8.3-character uppercase name without the ”.” .
        Dim cmd_str As String = "P" + file_name
        Dim ret_str As String = ""
        Dim len As Int16 = Dont_care_length
        Dim wait_time As Integer = Default_wait_time

        If (with_0x0a = True) Then
            cmd_str = cmd_str + Chr(10)
        End If

        SendWaitAck(cmd_str, ret_str, len, wait_time)

        ' Force HEX window to enter new line for better visualization
        If (with_0x0a = True) Then
            RichTextBox2.AppendText(Chr(13) + Chr(10))
        End If

        Return ret_str

    End Function

    Private Function Play_FileNumber(ByVal file_number As Integer, _
                                   Optional ByVal with_0x0a As Boolean = True) As String
        ' pnumbernn - play file by number (small-case p)
        Dim cmd_str As String = "p" + CStr(file_number)
        Dim ret_str As String = ""
        Dim len As Int16 = Dont_care_length
        Dim wait_time As Integer = Default_wait_time

        If (with_0x0a = True) Then
            cmd_str = cmd_str + Chr(10)
        End If

        SendWaitAck(cmd_str, ret_str, len, wait_time)
        Return ret_str

    End Function

    Private Function Empty_Cmd() As String
        ' 
        Dim cmd_str As String = Chr(10)
        Dim ret_str As String = ""
        Dim len As Int16 = Dont_care_length
        Dim wait_time As Integer = Default_wait_time

        SendWaitAck(cmd_str, ret_str, len, wait_time)
        Return ret_str

    End Function


    '
    '
    '
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        File_Play_Mode()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Cont_Play_Mode()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Cancel_Play()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Vol_Up()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Vol_Down()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Pause_On()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Normal_Play()
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Faster_Play()
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Previous_song()
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Next_song()
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Get_Info()
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        Powers_Down()
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        List_Files()
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        Play_Filename(ListBox1.Text + "OGG", True)
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        Play_FileNumber(CInt(ListBox2.Text))
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        Cancel_Play(True)
    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        Cont_Play_Mode(True)
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        Empty_Cmd()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        StartPlaying()
    End Sub


    Enum MyEnumKeyword
        DONE = 0
        PLAY
        FILES
        NoFAT
        FAT
        NotSD
        SD
        MAX_KEYWORD
    End Enum

    Private KeyWordList() As String = {"done", "play", "files", "nofat", "fat", "!SD", "SD"}
    Private KeyWordOptionLen() As Integer = {0, 0, 4, 0, 2, 11, 0}
    Private KeyWordCouter() As Integer = {0, 0, 0, 0, 0, 0, 0}
    Private ParseWordQueue As Queue(Of String) = New Queue(Of String)

    Private Sub ParseEvent()

    End Sub

    Private Sub WaitString(ByVal wait_str As String)
        Dim dummy_buffer(1) As Byte
        WaitStringAndOption(wait_str, 0, dummy_buffer)
    End Sub

    Private Sub WaitStringAndOption(ByVal wait_str As String, _
                                    ByVal option_len As Integer, _
                                    ByRef option_buf() As Byte)
        Dim my_str(1000) As Byte
        Dim wait_len = wait_str.Length()
        Dim my_str_index As Integer = 0
        Dim my_match_index As Integer = 0
        Dim my_match_position As Integer = 0
        Dim found_flag As Boolean = False

        ' Loop until wait string found
        Do
            If (read_data.Count > 0) Then
                Dim temp As Byte = read_data.Dequeue()

                ' Check if this one is matching next waiting char
                If (temp = Asc(wait_str(my_match_index))) Then
                    my_match_position = my_str_index
                    my_match_index += 1
                    ' Match all string
                    If (my_match_index = wait_len) Then
                        found_flag = True
                    End If
                Else
                    my_match_index = 0
                End If
                ' Store
                my_str(my_str_index) = temp
                my_str_index += 1
            Else
                Delay_loop(2)
            End If
        Loop While (found_flag = False)

        Dim my_option_index As Integer = 0
        ' Loop until option received
        While (my_option_index < option_len)
            If (read_data.Count > 0) Then
                Dim temp As Byte = read_data.Dequeue()
                option_buf(my_option_index) = temp
                my_option_index += 1
            Else
                Delay_loop(2)
            End If
        End While

    End Sub

    Dim fat_return(5 - 1) As Byte
    Dim files_return(3 - 1) As Byte

    Private Sub StartPlaying()
        ' Clear to 0
        For temp_index = 0 To (3 - 1)
            files_return(temp_index) = 0
        Next
        WaitString(KeyWordList(MyEnumKeyword.SD) + Chr(10))
        WaitStringAndOption(KeyWordList(MyEnumKeyword.FAT), 5, fat_return)
        WaitStringAndOption(KeyWordList(MyEnumKeyword.FILES), 3, files_return)
        File_Play_Mode()
        Cancel_Play(True)
        WaitString(KeyWordList(MyEnumKeyword.DONE) + Chr(10))
        Play_Filename("AUDIO002OGG")
    End Sub

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        Dim previous_song_no = ListBox1.Items.Count()
        Dim new_song_no = files_return(0) * 256 + files_return(1)

        If (new_song_no > previous_song_no) Then
            For temp_index As Integer = (previous_song_no + 1) To new_song_no
                Dim temp_no_str As String
                If (temp_index < 10) Then
                    temp_no_str = "00" + Chr(Asc("0") + temp_index)
                ElseIf (temp_index < 100) Then
                    Dim temp_digit0 = temp_index Mod 10
                    Dim temp_digit1 = (temp_index - temp_digit0) / 10
                    temp_no_str = "0" + Chr(Asc("0") + temp_digit1) + Chr(Asc("0") + temp_digit0)
                Else
                    Dim temp_digit0, temp_digit1, temp_digit2
                    temp_digit1 = temp_index Mod 100
                    temp_digit2 = (temp_index - temp_digit1) / 100
                    temp_digit0 = temp_index Mod 10
                    temp_digit1 = (temp_digit1 - temp_digit0) / 10
                    temp_no_str = Chr(Asc("0") + temp_digit2) + Chr(Asc("0") + temp_digit1) + Chr(Asc("0") + temp_digit0)
                End If

                ListBox1.Items.Add("AUDIO" + temp_no_str + "OGG")
            Next
        Else

        End If

    End Sub

    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        Dim cmd_str As String = "F" + Chr(10)
        Dim ret_str As String = ""
        Dim len As Int16 = 2
        Dim wait_time As Integer = Default_wait_time
        SendWaitAck(cmd_str, ret_str, len, wait_time)
    End Sub

 End Class
