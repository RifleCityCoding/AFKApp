Imports System.Drawing
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Diagnostics

Public Class Form1
    Inherits Form

    Private Const MOD_CTRL As Integer = &H2
    Private Const VK_P As Integer = &H50
    Private Const WH_MOUSE_LL As Integer = 14
    Private Const WM_MOUSEMOVE As Integer = &H200
    Private Const WM_HOTKEY As Integer = &H312
    Private Const HOTKEY_ID As Integer = 1 ' ID for the hotkey registration

    Private Shared hookID As IntPtr = IntPtr.Zero
    Private Shared hookProc As LowLevelMouseProc = New LowLevelMouseProc(AddressOf HookCallback)
    Private Shared enableMouseHook As Boolean = True ' Control flag for the mouse hook
    Private Shared instance As Form1 ' Static reference to the form

    Public Sub New()
        InitializeComponent()
        StartPosition = FormStartPosition.CenterScreen
        FormBorderStyle = FormBorderStyle.None ' Remove any border
        WindowState = FormWindowState.Maximized ' Maximize window
        TopMost = True ' Stay on top
        ShowInTaskbar = False ' Optionally hide from taskbar
        instance = Me ' Set the static instance
        hookID = SetHook(hookProc)
        Me.BackgroundImage = Image.FromFile("C:/Users/JamesS/source/repos/AFKApp/Images/codebg.png")
        Me.BackgroundImageLayout = ImageLayout.Stretch
    End Sub

    Protected Overrides Sub OnLoad(e As EventArgs)
        MyBase.OnLoad(e)
        RegisterHotKey(Handle, HOTKEY_ID, MOD_CTRL, VK_P)
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        UnregisterHotKey(Handle, HOTKEY_ID)
        UnhookWindowsHookEx(hookID)
        MyBase.OnFormClosing(e)
    End Sub

    Private Sub ShowPopup()
        If Me.InvokeRequired Then
            Me.Invoke(New MethodInvoker(AddressOf ShowPopup))
            Return
        End If
        enableMouseHook = False ' Disable hook processing
        Dim result As DialogResult = MessageBox.Show("Are you the current user?", "Popup", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            ShowGifWindow()
        Else
            ShowPopup() ' Recursion in original should be reviewed
        End If
        enableMouseHook = True ' Re-enable hook processing
    End Sub

    Private Sub ShowGifWindow()
        Dim gifForm As New Form()
        gifForm.TopMost = True
        gifForm.StartPosition = FormStartPosition.CenterScreen
        gifForm.Size = New Size(2000, 2000)
        gifForm.FormBorderStyle = FormBorderStyle.FixedSingle

        Dim pictureBox As New PictureBox()
        pictureBox.Dock = DockStyle.Fill
        pictureBox.Image = Image.FromFile("C:/Users/JamesS/source/repos/AFKApp/Images/nedry.gif")
        pictureBox.SizeMode = PictureBoxSizeMode.StretchImage
        gifForm.Controls.Add(pictureBox)
        gifForm.ShowDialog()
    End Sub

    Private Delegate Function LowLevelMouseProc(nCode As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr

    Private Shared Function HookCallback(nCode As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
        If nCode >= 0 AndAlso wParam.ToInt32() = WM_MOUSEMOVE AndAlso enableMouseHook Then
            instance.Invoke(New MethodInvoker(AddressOf instance.ShowPopup))
        End If
        Return CallNextHookEx(hookID, nCode, wParam, lParam)
    End Function

    Protected Overrides Sub WndProc(ByRef m As Message)
        MyBase.WndProc(m)
        If m.Msg = WM_HOTKEY AndAlso m.WParam.ToInt32() = HOTKEY_ID Then
            Me.Close() ' Closes the form
        End If
    End Sub

    Private Shared Function SetHook(proc As LowLevelMouseProc) As IntPtr
        Using curProcess As Process = Process.GetCurrentProcess()
            Using curModule As ProcessModule = curProcess.MainModule
                Dim modHandle As IntPtr = GetModuleHandle(curModule.ModuleName)
                Return SetWindowsHookEx(WH_MOUSE_LL, proc, modHandle, 0)
            End Using
        End Using
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function SetWindowsHookEx(idHook As Integer, lpfn As LowLevelMouseProc, hMod As IntPtr, dwThreadId As UInteger) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function UnhookWindowsHookEx(hhk As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function CallNextHookEx(hhk As IntPtr, nCode As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function GetModuleHandle(lpModuleName As String) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Shared Function RegisterHotKey(hWnd As IntPtr, id As Integer, fsModifiers As Integer, vk As Integer) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function UnregisterHotKey(hWnd As IntPtr, id As Integer) As Boolean
    End Function
End Class
