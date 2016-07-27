'Developed by Kalvin Aung
'Company The World Management Pte Ltd
'Version 1.0 Revised
'Date : 2009 07 13
'Change Log : 1.0
'Add DialogTitle, DefaultExtension, Default Directory
'Use Get Process ID function for CheckInstance Module
'Combine Check Instance Module

'########################################################################################
'######                             Check  INSTANCE                                 #####
'########################################################################################
Imports System.Management 'Need to ADD REF System.Management
Imports System.Diagnostics
'########################################################################################
'######                             SHOW FILE DIALOG                                #####
'########################################################################################
Imports System.Threading
Imports System.Security.Permissions
Imports System.Windows.Forms
'########################################################################################
'######                        E N D  O F  I M P O R T                              #####
'########################################################################################

Module kOpenFileDialog

    '################################################################################################################
    '######                             SHOW FILE DIALOG BEGIN                                                  #####
    '################################################################################################################
    Public FileName As String
    Public cParentID As Integer
    Public cMachineName As String

    '####################################################################
    '#### This 4 Global Variables can be initialized during runtime  ####
    '####################################################################
    Public cDlgDefaultDir As String = ""
    Public cDlgDefaultExt As String = ""
    Public cDlgTitle As String = "Open File"
    Public cDlgFileFilter As String = "All files(*.)|*.*" 'Change File type filter
    '####################################################################

    Public Class WindowWrapper

        Implements System.Windows.Forms.IWin32Window
        Private _hwnd As IntPtr

        Public Sub New(ByVal handle As IntPtr)
            _hwnd = handle
        End Sub

        Public ReadOnly Property Handle() As System.IntPtr Implements System.Windows.Forms.IWin32Window.Handle
            Get
                Return _hwnd
            End Get
        End Property

    End Class

    Public Function FindFile(ByVal xApplication As SAPbouiCOM.Application) As String

        Try
            Dim cProcess As System.Diagnostics.Process = System.Diagnostics.Process.GetCurrentProcess()
            cParentID = getProcessParentID(cProcess.ProcessName, cProcess.Id)
            Dim cParent As System.Diagnostics.Process = System.Diagnostics.Process.GetProcessById(cParentID, cProcess.MachineName)
            cMachineName = cProcess.MachineName
            If Right(cProcess.ProcessName, 6) = "vshost" Then 'This part is for RUNTIME Mode
                xApplication.SetStatusBarMessage(cProcess.Id & " 2N " & cProcess.ProcessName & " << " & cParent.Id & " N " & cParent.ProcessName & " <<< This is DESIGN-TIME >>>", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                cParentID = 0
            Else
                'xApplication.SetStatusBarMessage(cProcess.Id & " 2N " & cProcess.ProcessName & " << " & cParent.Id & " N " & cParent.ProcessName, SAPbouiCOM.BoMessageTime.bmt_Short, False)
            End If
        Catch ex As Exception
            xApplication.SetStatusBarMessage("Process Reading Phail!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
        Dim ShowFolderBrowserThread As Threading.Thread
        Try
            ShowFolderBrowserThread = New Threading.Thread(AddressOf ShowFolderBrowser)
            If ShowFolderBrowserThread.ThreadState = System.Threading.ThreadState.Unstarted Then
                ShowFolderBrowserThread.SetApartmentState(System.Threading.ApartmentState.STA)
                ShowFolderBrowserThread.Start()
            ElseIf ShowFolderBrowserThread.ThreadState = System.Threading.ThreadState.Stopped Then
                ShowFolderBrowserThread.Start()
                ShowFolderBrowserThread.Join()
            End If
            While ShowFolderBrowserThread.ThreadState = Threading.ThreadState.Running
                Windows.Forms.Application.DoEvents()
            End While
            If FileName <> "" Then
                Return FileName
            End If
        Catch ex As Exception
            xApplication.MessageBox("FileFile" & ex.Message)
        End Try

        Return ""

    End Function

    Public Sub ShowFolderBrowser()
        FileName = ""
        Dim OpenFile As New OpenFileDialog
        Try
            OpenFile.Multiselect = False
            If cDlgDefaultDir <> "" Then OpenFile.InitialDirectory = cDlgDefaultDir
            If cDlgTitle <> "" Then OpenFile.Title = cDlgTitle
            If cDlgDefaultExt <> "" Then OpenFile.DefaultExt = cDlgDefaultExt
            OpenFile.Filter = cDlgFileFilter
            Dim filterindex As Integer = 0
            Try
                filterindex = 0
            Catch ex As Exception
            End Try

            OpenFile.FilterIndex = filterindex
            OpenFile.RestoreDirectory = True
            If cParentID = 0 Then
                Dim MyProcs() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcessesByName("SAP Business One")
                Dim MyWindow As New WindowWrapper(MyProcs(0).MainWindowHandle)
                Dim ret As DialogResult = OpenFile.ShowDialog(MyWindow)

                If ret = DialogResult.OK Then
                    FileName = OpenFile.FileName
                    OpenFile.Dispose()
                Else
                    System.Windows.Forms.Application.ExitThread()
                End If
            Else
                Dim kProcs As System.Diagnostics.Process = System.Diagnostics.Process.GetProcessById(cParentID, cMachineName)
                Dim xWindow As New WindowWrapper(kProcs.MainWindowHandle)
                Dim ret As DialogResult = OpenFile.ShowDialog(xWindow)

                If ret = DialogResult.OK Then
                    FileName = OpenFile.FileName
                    OpenFile.Dispose()
                Else
                    System.Windows.Forms.Application.ExitThread()
                End If
            End If
        Catch ex As Exception
            FileName = "Error " & ex.Message.ToString
        Finally
            OpenFile.Dispose()
        End Try

    End Sub
    '################################################################################################################
    '######                             SHOW FILE DIALOG END                                                    #####
    '################################################################################################################

    '################################################################################################################
    '# Usage                                                                                                        #
    '################################################################################################################
    ' Set File filter, default is all file
    ' cFileFilter = "Form File <*.xml> |*.xml"
    ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ' Call the function giving Application object as parameter
    ' strvariable = FindFile(SBO_Application)
    ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ' Return string is full path + file name of dialogresult OK
    ' Return Blank when dialogresult cancel
    '################################################################################################################

    Public Function checkInstance(Optional ByVal killFlag As Boolean = False) As Boolean ''Return Value : Returns FALSE if there's an running process
        Dim BufferFlag As Boolean = True
        Dim cProcess As Process = Process.GetCurrentProcess()
        Dim aProcesses() As Process = Process.GetProcessesByName(cProcess.ProcessName)

        Dim cParentID As Single = getProcessParentID(cProcess.ProcessName, cProcess.Id)
        Dim xParentID As Single = 0

        For Each xProcess As Process In aProcesses
            If xProcess.Id <> cProcess.Id Then 'ignore the current (self)
                If System.Reflection.Assembly.GetExecutingAssembly().Location = cProcess.MainModule.FileName Then                 'Check the running process with same EXE 
                    xParentID = getProcessParentID(xProcess.ProcessName, xProcess.Id)
                    If xParentID = cParentID Then

                        If killFlag = True Then
                            xProcess.Kill()
                            'MessageBox.Show("New / Parent = " & cProcess.Id & " / " & cParentID & " Old / Parent = " & xProcess.Id & " / " & xParentID, " Old Application was killed ")
                            MessageBox.Show("Running Addon for the same instance of SAP was terminated.", "old Process Killed", MessageBoxButtons.OK)
                            BufferFlag = True
                        Else
                            'If MessageBox.Show("New / Parent = " & cProcess.Id & " / " & cParentID & " Old / Parent = " & xProcess.Id & " / " & xParentID, " Wanna Kill Old Application ?", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                            If MessageBox.Show("Found Same Addon was running for the same instance of SAP, wan to terminate the old ?", "wanna kill old process ? ", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                                xProcess.Kill()
                                BufferFlag = True
                            Else
                                MessageBox.Show("Application is already running", "Program Terminated!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                                BufferFlag = False
                            End If
                        End If
                        Exit For
                    End If
                End If
            End If
        Next
        Return BufferFlag
    End Function
    'Need to Imports System.Management ( Add Reference )
    Public Function getProcessParentID(ByVal cName As String, ByVal cID As Integer) As Integer
        Dim query As SelectQuery = New SelectQuery("SELECT * FROM Win32_Process WHERE Name like '" & cName & ".exe' and ProcessId = " & cID)
        Dim mgmtSearcher As ManagementObjectSearcher = New ManagementObjectSearcher(query)
        Dim kRet As Integer = -1
        For Each p As ManagementObject In mgmtSearcher.Get()
            Dim s(1) As String
            p.InvokeMethod("GetOwner", DirectCast(s, Object()))
            ' Source Code link : http://www.vbdotnetforums.com/windows-services/4022-kill-specific-process.html
            ' More Object Reference at this link : http://msdn.microsoft.com/en-us/library/aa394372(VS.85).asp
            kRet = p("ParentProcessId")
        Next
        Return kRet
    End Function
    'USAGE of FileDialog

    'cDlgDefaultDir = "C:\temp"
    'cDlgDefaultExt = "*.xml"
    'cDlgTitle = "Open FORM xml file"
    'cDlgFileFilter = "Form File <*.xml> |*.xml"
    'SBO_Application.SetStatusBarMessage(FindFile(SBO_Application), SAPbouiCOM.BoMessageTime.bmt_Short, False)


    'USAGE of CheckInstance
    ' ###### Code Scrap in SubMain #####
    '    Try
    '        If checkInstance(True) = True Then
    '            Dim EventHandler As EventHandler
    '            EventHandler = New EventHandler
    '            System.Windows.Forms.Application.Run()
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.StackTrace & ":" & ex.Message)
    '    End Try
End Module
