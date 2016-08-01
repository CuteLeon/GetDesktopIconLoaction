Imports System.Runtime.InteropServices

Public Class Form1
#Region "声明区"
    Private Declare Function GetDesktopWindow Lib "user32" Alias "GetDesktopWindow" () As IntPtr
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Integer, ByVal hWnd2 As Integer, ByVal lpsz1 As String, ByVal lpsz2 As String) As IntPtr
    Private Declare Function VirtualAllocEx Lib "kernel32.dll" (ByVal hProcess As IntPtr, ByVal lpAddress As IntPtr, ByVal dwSize As Integer, ByVal flAllocationType As Integer, ByVal flProtect As Integer) As IntPtr
    Private Declare Function VirtualFreeEx Lib "kernel32.dll" (ByVal hProcess As IntPtr, ByRef lpAddress As IntPtr, ByRef dwSize As Integer, ByVal dwFreeType As Integer) As Integer
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As IntPtr) As Integer
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As Integer) As Integer
    Private Declare Function GetWindowThreadProcessId Lib "user32" Alias "GetWindowThreadProcessId" (ByVal hWnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
    Private Declare Function OpenProcess Lib "kernel32" Alias "OpenProcess" (ByVal dwDesiredAccess As Integer, ByVal bInheritHandle As Boolean, ByVal dwProcessId As Integer) As IntPtr
    Private Declare Function CloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As IntPtr) As Integer
    Private Declare Function ReadProcessMemory Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As IntPtr, ByRef lpBaseAddress As IntPtr, ByRef lpBuffer As Point, ByVal nSize As Integer, ByRef lpNumberOfBytesWritten As Integer) As Integer
    Private Declare Function ReadProcessMemory Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As IntPtr, ByRef lpBaseAddress As IntPtr, ByVal lpBuffer As IntPtr, ByVal nSize As Integer, ByRef lpNumberOfBytesWritten As Integer) As Integer
    Private Declare Function WriteProcessMemory Lib "kernel32" Alias "WriteProcessMemory" (ByVal hProcess As IntPtr, ByRef lpBaseAddress As IntPtr, ByRef lpBuffer As IntPtr, ByVal nSize As Integer, ByRef lpNumberOfBytesWritten As Integer) As Integer

    Public Const PROCESS_VM_READ = &H10
    Public Const PROCESS_VM_OPERATION = &H8
    Public Const PROCESS_VM_WRITE = &H20
    Public Const MEM_COMMIT = &H1000
    Public Const PAGE_READWRITE = &H4
    Public Const LVM_GETITEMPOSITION = &H1010
    Public Const LVM_SETITEMPOSITION = &H100F
    Public Const LVM_GETITEMCOUNT = &H1004
    Public Const LVM_GETITEMTEXT = &H1073
    Public Const MAX_PATH = 260

    Public Structure tagLVITEMW
        Public mask As UInteger
        Public iItem As Integer
        Public iSubItem As Integer
        Public state As UInteger
        Public stateMask As UInteger
        Public pszText As Int64
        Public cchTextMax As Integer
        Public iImage As Integer
        Public lParam As Integer
        Public iIndent As Integer
        Public iGroupId As Integer
        Public cColumns As UInteger
        Public puColumns As IntPtr
        Public piColFmt As IntPtr
        Public iGroup As Integer
    End Structure

    Public Structure DesktopIcon
        Public Text As String
        Public Position As Point
    End Structure

#End Region

#Region "窗体"

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GetIconInfo()
        End
    End Sub
#End Region

#Region "接口函数"
    Private Sub GetIconInfo()
        Dim Icons As ArrayList = New ArrayList
        Dim HandleSysListView32 As IntPtr = GetDesktopIconHandle()
        Dim dwProcessId As Integer
        GetWindowThreadProcessId(HandleSysListView32, dwProcessId)
        Dim hProcess As IntPtr = OpenProcess(PROCESS_VM_OPERATION + PROCESS_VM_READ + PROCESS_VM_WRITE, False, dwProcessId)

        Dim PV As IntPtr = VirtualAllocEx(hProcess, IntPtr.Zero, Marshal.SizeOf(GetType(Point)), MEM_COMMIT, PAGE_READWRITE)
        Dim pn As IntPtr = VirtualAllocEx(hProcess, IntPtr.Zero, MAX_PATH * 2, MEM_COMMIT, PAGE_READWRITE)
        Dim plvitem As IntPtr = VirtualAllocEx(hProcess, IntPtr.Zero, GetLvItemSize(), MEM_COMMIT, PAGE_READWRITE)
        Dim IconCount As Integer = ListView_GetItemCount(HandleSysListView32)

        For Index As Integer = 0 To IconCount - 1
            Dim TempIcon As DesktopIcon = New DesktopIcon
            ListView_GetItemText(HandleSysListView32, Index, hProcess, plvitem, pn, MAX_PATH * 2)
            ListView_GetItemPosition(HandleSysListView32, Index, PV)
            Dim lpNumberOfBytesRead As Integer
            ReadProcessMemory(hProcess, PV, TempIcon.Position, Marshal.SizeOf(GetType(Point)), lpNumberOfBytesRead)

            Dim pszName As IntPtr = Marshal.AllocCoTaskMem(MAX_PATH * 2)
            ReadProcessMemory(hProcess, pn, pszName, MAX_PATH * 2, lpNumberOfBytesRead)
            TempIcon.Text = Marshal.PtrToStringUni(pszName)
            Marshal.FreeCoTaskMem(pszName)
            Icons.Add(Icon)

            Debug.Print(Index & ".[" & TempIcon.Text & "] ： " & TempIcon.Position.X & " , " & TempIcon.Position.Y)
        Next

        VirtualFreeEx(hProcess, PV, Marshal.SizeOf(GetType(Point)), 0)
        VirtualFreeEx(hProcess, pn, MAX_PATH * 2, 0)
        VirtualFreeEx(hProcess, plvitem, GetLvItemSize(), 0)
        CloseHandle(hProcess)
    End Sub
#End Region

#Region "功能函数"

    Private Function GetDesktopIconHandle() As IntPtr
        Dim HandleDesktop As Integer = GetDesktopWindow
        Dim HandleTop As Integer = 0
        Dim LastHandleTop As Integer = 0
        Dim HandleSHELLDLL_DefView As Integer = 0
        Dim HandleSysListView32 As Integer = 0
        '在WorkerW里搜索
        Do Until HandleSysListView32 > 0
            HandleTop = FindWindowEx(HandleDesktop, LastHandleTop, "WorkerW", vbNullString)
            HandleSHELLDLL_DefView = FindWindowEx(HandleTop, 0, "SHELLDLL_DefView", vbNullString)
            If HandleSHELLDLL_DefView > 0 Then HandleSysListView32 = FindWindowEx(HandleSHELLDLL_DefView, 0, "SysListView32", "FolderView")
            LastHandleTop = HandleTop
            If LastHandleTop = 0 Then Exit Do
        Loop
        '如果找到了，立即返回
        If HandleSysListView32 > 0 Then Return HandleSysListView32
        '在Progman里搜索
        Do Until HandleSysListView32 > 0
            HandleTop = FindWindowEx(HandleDesktop, LastHandleTop, "Progman", "Program Manager")
            HandleSHELLDLL_DefView = FindWindowEx(HandleTop, 0, "SHELLDLL_DefView", vbNullString)
            If HandleSHELLDLL_DefView > 0 Then HandleSysListView32 = FindWindowEx(HandleSHELLDLL_DefView, 0, "SysListView32", "FolderView")
            LastHandleTop = HandleTop
            If LastHandleTop = 0 Then Exit Do : Return 0
        Loop
        Return HandleSysListView32
    End Function

    Public Function ListView_GetItemPosition(ByVal hwndLV As IntPtr, ByVal i As Integer, ByRef ppt As IntPtr) As Integer
        Return SendMessage(hwndLV, LVM_GETITEMPOSITION, i, ppt)
    End Function

    Public Function ListView_SetItemPosition(ByVal hwndLV As IntPtr, ByVal i As Integer, ByVal x As Integer, ByVal y As Integer) As Integer
        Return SendMessage(hwndLV, LVM_SETITEMPOSITION, i, (x + y * 65536))
    End Function

    Public Function ListView_GetItemCount(ByVal hwndLV As IntPtr) As Integer
        Return SendMessage(hwndLV, LVM_GETITEMCOUNT, 0, 0)
    End Function

    Public Function ListView_GetItemText(ByVal hwndLV As IntPtr, ByVal i As Integer, ByRef hProcess As IntPtr, ByRef plvitem As IntPtr, ByRef pszText_ As IntPtr, ByVal cchTextMax As Integer) As Integer
        Dim _macro_lvi As tagLVITEMW = New tagLVITEMW()
        _macro_lvi.iSubItem = 0
        _macro_lvi.cchTextMax = cchTextMax
        _macro_lvi.pszText = pszText_
        Dim lv As IntPtr = Marshal.AllocCoTaskMem(GetLvItemSize())
        Marshal.StructureToPtr(_macro_lvi, lv, False)
        Dim lpNumberOfBytesWritten As Integer
        Dim result As Integer = WriteProcessMemory(hProcess, plvitem, lv, GetLvItemSize(), lpNumberOfBytesWritten)
        Marshal.FreeCoTaskMem(lv)
        Return SendMessage(hwndLV, LVM_GETITEMTEXT, i, plvitem)
    End Function

    Public Function GetLvItemSize() As Integer
        Return Marshal.SizeOf(GetType(tagLVITEMW))
    End Function

#End Region

End Class
