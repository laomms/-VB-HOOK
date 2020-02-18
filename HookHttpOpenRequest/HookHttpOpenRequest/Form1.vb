Imports System.IO
Imports System.Runtime.InteropServices
Imports HookApi

'函数原型

'HttpOpenRequest　　HINTERNET HttpOpenRequest （HINTERNET hInternet ，LPCTSTR lpszUrl ，LPCTSTR lpszHeaders ，DWORD dwHeadersLength ，DWORD dwFlags ，DWORD_PTR dwContext）

Public Class Form1

    Friend Shared CheatFileHandle As IntPtr = IntPtr.Zero   '要替换的文件的句柄，来源于HttpOpenRequest的返回值。
    Friend Shared CheatFile() As Byte = File.ReadAllBytes(My.Application.Info.DirectoryPath & "\abc.jpg")    '用于替换的文件
    Private Shared curcnt As Integer = 0

    <DllImport("wininet.dll")>
    Public Shared Function HttpOpenRequestW(hConnect As IntPtr, szVerb As IntPtr, szURI As IntPtr, szHttpVersion As IntPtr, szReferer As IntPtr, accetpType As IntPtr, dwflags As Integer, dwcontext As IntPtr) As IntPtr
    End Function
    <DllImport("wininet.dll", SetLastError:=True)>
    Public Shared Function InternetReadFile(ByVal hFile As IntPtr, ByVal lpBuffer As IntPtr, ByVal dwNumberOfBytesToRead As Integer, ByRef lpdwNumberOfBytesRead As Integer) As Boolean
    End Function

    Private HttpOpenRequestW_Hook As New APIHOOK()
    Private InternetReadFile_Hook As New APIHOOK()
    '定义一个引用变量以防止垃圾回收机制回收回调

    Private HttpOpenRequestWDelegate As HttpOpenRequestWCallback
    Private InternetReadFileDelegate As InternetReadFileCallback

    Private Delegate Function HttpOpenRequestWCallback(hConnect As IntPtr, szVerb As IntPtr, szURI As IntPtr, szHttpVersion As IntPtr, szReferer As IntPtr, accetpType As IntPtr, dwflags As Integer, dwcontext As IntPtr) As IntPtr
    Private Delegate Function InternetReadFileCallback(ByVal hFile As IntPtr, ByVal lpBuffer As IntPtr, ByVal dwNumberOfBytesToRead As Integer, ByRef lpdwNumberOfBytesRead As Integer) As Boolean



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click



        'If HttpOpenRequestW_Hook.Installed Then HttpOpenRequestW_Hook.Uninstall()

        HttpOpenRequestWDelegate = New HttpOpenRequestWCallback(AddressOf HttpOpenRequestWProc)
        InternetReadFileDelegate = New InternetReadFileCallback(AddressOf InternetReadFileProc)

        HttpOpenRequestW_Hook.Install("wininet.dll", "HttpOpenRequestW", Marshal.GetFunctionPointerForDelegate(HttpOpenRequestWDelegate))
        InternetReadFile_Hook.Install("wininet.dll", "InternetReadFile", Marshal.GetFunctionPointerForDelegate(InternetReadFileDelegate))

        HttpOpenRequestW_Hook.Hook()



    End Sub



    Private Function HttpOpenRequestWProc(hConnect As IntPtr, szVerb As IntPtr, szURI As IntPtr, szHttpVersion As IntPtr, szReferer As IntPtr, accetpType As IntPtr, dwflags As Integer, dwcontext As IntPtr) As IntPtr

        '注意：在钩CreateFile等函数时可能需要修改调试选项以便可以从非托管进入托管，并且不应直接使用debug.print等函数进行显示



        '卸载钩子以便调用原函数

        HttpOpenRequestW_Hook.UnHook()

        '调用原函数
        Dim uri As String = Marshal.PtrToStringUni(szURI)
        Dim ret As Integer = HttpOpenRequestW(hConnect, szVerb, szURI, szHttpVersion, szReferer, accetpType, dwflags, dwcontext)
        If uri.Contains("/8MXBCxc") Then '根据名称区分要替换的图片. 
            InternetReadFile_Hook.Hook()
            CheatFileHandle = ret
        End If
        '加载钩子以便继续获取数据

        'HttpOpenRequest_Hook.Hook()

        Return ret

    End Function

    Private Function InternetReadFileProc(ByVal hFile As IntPtr, ByVal lpBuffer As IntPtr, ByVal dwNumberOfBytesToRead As Integer, ByRef lpdwNumberOfBytesRead As Integer) As Boolean

        If hFile = CheatFileHandle Then
            If curcnt = CheatFile.Length Then
                CheatFileHandle = IntPtr.Zero
                curcnt = 0
                lpdwNumberOfBytesRead = 0
            Else
                If curcnt + dwNumberOfBytesToRead <= CheatFile.Length Then
                    lpdwNumberOfBytesRead = dwNumberOfBytesToRead
                    Marshal.Copy(CheatFile, curcnt, lpBuffer, lpdwNumberOfBytesRead)
                    curcnt += dwNumberOfBytesToRead
                Else
                    lpdwNumberOfBytesRead = CheatFile.Length - curcnt
                    Marshal.Copy(CheatFile, curcnt, lpBuffer, lpdwNumberOfBytesRead)
                    curcnt = CheatFile.Length
                End If
            End If
            Return True
        Else
            Return InternetReadFile(hFile, lpBuffer, dwNumberOfBytesToRead, lpdwNumberOfBytesRead)
        End If
    End Function


    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        HttpOpenRequestW_Hook.Uninstall()
        InternetReadFile_Hook.Uninstall()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        WebBrowser1.Navigate("https://ibb.co/8MXBCxc")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        HttpOpenRequestW_Hook.Uninstall()
        InternetReadFile_Hook.Uninstall()
    End Sub
End Class