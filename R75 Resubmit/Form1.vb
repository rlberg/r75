Option Explicit On

Imports System.Deployment.Application
Imports System.IO
Imports System.Diagnostics
Imports System.Net.Mail
Imports bgw = System.ComponentModel
Imports pcom = AutPSTypeLibrary

Public Class MainForm

    Dim OpenFileLocation As String
    Dim SavetoLocation As String

    Dim usrNm As String

    'RxClaim
    Dim objRx As pcom.AutPS
    Dim objWait As Object
    Dim objMgr, objMgr2 As Object
    Dim autECLConnList As Object
    Dim ObjSessionHandle As Integer

    'Excel Object Variables
    Dim objExcelFilePath As String
    Dim objExcel = CreateObject("Excel.Application")
    Dim objWorkbook1
    Dim objWorksheet1
    Dim objWorksheet1Count As Integer

    Dim bStopApp As Boolean = False

    '*****  This is for the columns in the spreadsheet that will track the results of each record  *****
    Dim col_FINAL_OUTCOME As Integer = 27
    Dim col_MESSAGE As Integer = 28
    Dim col_SUBMITDATE As Integer = 29
    Dim col_TAT As Integer = 30
    '***************************************************************************************************

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        Me.Cursor = Cursors.WaitCursor
        btnStart.Enabled = False

        'Username
        GetUsername()

        Do While bStopApp = False
            OpenSpreadsheetTemplate()

            If objWorksheet1Count > 1 Then
                Dim i As Integer

                OpenRxClaimSession()    'Open an RxClaim session/window

                Initialize_RxClaim_Screen()

                For i = 2 To objWorksheet1Count      'Start i on 2 because that is the 1st row we can start with (row1 is the header)

                    GoHome()        'will bring the RxClaim screen all the way back home

                    GetTo_JobScheduleList_Screen(i)

                    FindCAG(i)
                Next
            Else
                MsgBox("Your spreadsheet was empty.")

            End If

            bStopApp = True






            'Get to the right screen


            '1.  Type in RX6 [enter]        'RX6 or PPF
            '   RxClaims Library
            '2.  Type "3" [enter]           'Screen name should be:  CCT600 - RxCLAIM Plan Administrator Menu
            '   Manual Claim
            '3.  Type "2" [enter]           'Screen name should be:  CCT630S - RxCLAIM Manual Claim Menu
            '   D0 Manual Claim
            '4.  Type "2" [enter]           'Screen name should be:  CCT632 - RxCLAIM D0 Manual Claim Maintenance
            '   Member Reimbursement


        Loop

        btnStart.Enabled = True
        Me.Cursor = Cursors.Arrow

    End Sub

    Public Sub Initialize_RxClaim_Screen()
        'IF 19,2 for 11 = "Press Enter"  ...  This is usually the 1st screen that shows if you already have another session open
        If Trim(objRx.GetText(19, 2, 11)) = "Press Enter" Then
            objRx.SendKeys("[Enter]")
            waitOnMe(1000)
        End If

        'This will be the case if it is notifying you that you have x days until password expires
        If Trim(objRx.GetText(21, 2, 11)) = "Press Enter" Then
            objRx.SendKeys("[Enter]")
            waitOnMe(1000)
        End If

        waitForMe()

        'IsRightScreenName("RX6", 9, 45, 5000)      'this only works if every users has access to the exact same memu options
        IsRightScreenName("Prime", 1, 33, 5000)

        waitForMe()

        If LCase(cmbEnv.SelectedItem) = "prod03" Then
            objRx.SetText("PPF", 21, 7)
        Else
            objRx.SetText("RX6", 21, 7)
        End If

        waitForMe()
        MoveMe("enter", 1)
    End Sub

    Public Sub FindCAG(iRow As Integer)
        ''Find the right line of coverage (based on Fill Date (col 3))  ****************************************************************
        'Dim d As String = objWorksheet1.Cells(iRow, 3).Value
        'Dim sFillDate As String = d.Substring(0, 2) & "/" & d.Substring(2, 2) & "/" & d.Substring(4, 4)
        'Dim dFillDate As Date = CDate(sFillDate)

        Dim sCarrier, sAccount, sGroup As String

        sCarrier = Trim(objWorksheet1.Cells(iRow, 3).Value)
        sAccount = Trim(objWorksheet1.Cells(iRow, 4).Value)
        sGroup = Trim(objWorksheet1.Cells(iRow, 5).Value)


        Dim iRowCounter As Integer
        Dim IsActiveEligFound As Boolean = False

        'rows 9, 13, 17, 21

        For y As Integer = 1 To 4               'This will allow us to page down up to 4 times
            For z As Integer = 0 To 3           'This will allow us to look at up to 4 records per page

                'iRowCounter = (z + 9) + (z * 3)
                iRowCounter = (z + 8) + (z * 3)

                If Trim(objRx.GetText(iRowCounter, 35, 10)) = sCarrier And Trim(objRx.GetText(iRowCounter, 46, 16)) = sAccount And Trim(objRx.GetText(iRowCounter, 63, 15)) = sGroup Then
                    IsActiveEligFound = True

                    'Chose this one by entering a "1"
                    SettingText("1", iRowCounter, 2)   'Text, row, col

                    objRx.SendKeys("[Enter]")

                    waitForMe()
                    Exit For
                End If
            Next

            If IsActiveEligFound = True Then
                Exit For
            Else
                'pagedown
                MoveMe("roll up", 1)
                waitForMe()
            End If
        Next

        If IsActiveEligFound = False Then      'Member by ID screen
            objWorksheet1.Cells(iRow, col_FINAL_OUTCOME).Value = "Error - Member by Id screen"
            objWorksheet1.Cells(iRow, col_MESSAGE).Value = "Could not find Active Line of Coverage"
            Exit Sub
        End If
    End Sub

    Public Sub GetTo_JobScheduleList_Screen(iRow As Integer)
        ''IF 19,2 for 11 = "Press Enter"  ...  This is usually the 1st screen that shows if you already have another session open
        'If Trim(objRx.GetText(19, 2, 11)) = "Press Enter" Then
        '    objRx.SendKeys("[Enter]")
        '    waitOnMe(1000)
        'End If

        ''Now try connecting to that session ... we will wait 5 seconds
        ''This is a hard wait to ensure that the RxClaim session has started
        'IsRightScreenName("QPADEV", 1, 70, 5000)

        'waitForMe()

        'If LCase(cmbEnv.SelectedItem) = "prod03" Then
        '    objRx.SetText("PPF", 21, 7)
        'Else
        '    objRx.SetText("RX6", 21, 7)
        'End If

        'waitForMe()
        'MoveMe("enter", 1)
        'waitForMe()

        IsRightScreenName("CCT600", 1, 2, 5000)
        waitForMe()
        objRx.SetText("3", 21, 7)
        waitForMe()
        MoveMe("enter", 1)
        waitForMe()

        IsRightScreenName("CCT630S", 1, 2, 5000)
        waitForMe()
        objRx.SetText("2", 21, 7)
        waitForMe()
        MoveMe("enter", 1)
        waitForMe()

        IsRightScreenName("CCT632", 1, 2, 5000)
        waitForMe()
        objRx.SetText("2", 21, 7)
        waitForMe()
        MoveMe("enter", 1)
        waitForMe()

        IsRightScreenName("RCNCP050D", 1, 2, 5000)
        waitForMe()

        'press f6
        MoveMe("pf6", 1)
        waitForMe()

        IsRightScreenName("RCNCP056", 1, 2, 60000)

        'BIN
        MoveMe2("eraseeof", 4, 11)
        SettingText(Trim(objWorksheet1.Cells(iRow, 11).Value), 4, 11)   'Text, row, col


        'Proc Ctrl
        MoveMe2("eraseeof", 4, 38)
        SettingText(Trim(objWorksheet1.Cells(iRow, 12).Value), 4, 38)   'Text, row, col

        'Group
        MoveMe2("eraseeof", 4, 58)
        SettingText(Trim(objWorksheet1.Cells(iRow, 13).Value), 4, 58)   'Text, row, col

        '**************************************************************************************************************

        'RX #
        MoveMe2("eraseeof", 5, 38)
        SettingText(Trim(objWorksheet1.Cells(iRow, 15).Value), 5, 38)   'Text, row, col

        'Fill Date
        MoveMe2("eraseeof", 7, 11)
        SettingText(Trim(objWorksheet1.Cells(iRow, 17).Value), 7, 11)   'Text, row, col              '*Spreadsheet will be in this format "yyymmdd"...But needs to be in this format: "mmddyy"

        'MemberId
        MoveMe2("eraseeof", 7, 38)
        SettingText(Trim(objWorksheet1.Cells(iRow, 7).Value), 7, 38)   'Text, row, col

        objRx.SendKeys("[Enter]")

        waitForMe()


    End Sub

    Public Sub OpenSpreadsheetTemplate()
        Try
            objExcelFilePath = "C:\Users\rlberg\Desktop\R75 Resubmit Application Template.xlsx"

            'objExcelFilePath = "J:\shrproj\Benefit Operations\Paper Claims\Claims Processing\Eligibility Rejects Macro\Stage Two - Eligibility Rejects Output File.xlsx"

            If objExcelFilePath = Nothing Then
                MsgBox("Sorry we couldn't find that spreadsheet")
                Exit Sub
            End If

            'objExcel = CreateObject("Excel.Application")
            objWorkbook1 = objExcel.Workbooks.Open(objExcelFilePath)
            objWorksheet1 = objWorkbook1.Worksheets(1)

            objWorksheet1Count = objWorksheet1.Range("A1").CurrentRegion.Rows.Count()        'Claim Count

            objExcel.Visible = True     '--only do this if you want to see the progress

        Catch ex As Exception
            MsgBox("Experienced an exception on OpenSpreadsheetTemplate():  " & ex.ToString)
        End Try
    End Sub

    Public Sub OpenRxClaimSession()
        Try
            objRx = CreateObject("PCOMM.autECLPS")
            objWait = CreateObject("PCOMM.autECLOIA")
            objMgr = CreateObject("PCOMM.autECLConnMgr")
            autECLConnList = CreateObject("PCOMM.autECLConnList")

            OpenNewSession()

            waitOnMe(4000)

            objMgr2 = CreateObject("PCOMM.autECLConnMgr")

            waitOnMe(1000)

            Dim y As Integer = ManageSessions()

            ObjSessionHandle = objMgr2.autECLConnList(y).Handle             'Errors out here

            objRx.SetConnectionByHandle(ObjSessionHandle)
            objWait.SetConnectionByHandle(ObjSessionHandle)

            waitForMe()
        Catch ex As Exception
            bStopApp = True
            MsgBox("Experienced an exception on OpenRxClaimSession():  " & ex.ToString)
        End Try
    End Sub

    Public Sub OpenNewSession()
        Dim Envir As String

        Try
            'Now find the "File name" to open up based on their selection
            If LCase(cmbEnv.SelectedItem) = "dev01" Then
                Envir = "Dev01.AS4"
            ElseIf LCase(cmbEnv.SelectedItem) = "dev02" Then
                Envir = "Dev02.AS4"
            ElseIf LCase(cmbEnv.SelectedItem) = "prod03" Then
                Envir = "PROD03.AS4"
            ElseIf LCase(cmbEnv.SelectedItem) = "prod01" Then
                Envir = "PROD01.AS4"
            Else
                MsgBox("Environment was not found ... exiting.")
                Exit Sub
            End If


            '***********************************************

            Dim sDir As String = getMyDocs()

            If sDir.Length > 1 Then
                'Now we are trying to open up a session
                Try
                    Process.Start(sDir & "RxClaims Sessions\" & Envir)
                Catch
                    Try
                        Process.Start("C:\Users\Public\Desktop\RxClaims Sessions\" & Envir)
                    Catch ex As Exception
                        MsgBox("Please open up an RxClaim session and then press 'OK' to this message")
                    End Try
                End Try
            Else
                MsgBox("couldn't find Desktop")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("Experienced an exception on OpenNewSession():  " & ex.ToString)
        End Try
    End Sub



    Sub MoveMe(command, amount)
        'Do what the command says and do it as many times as the amount says
        'Most common commands will be "tab" and "pf12"

        Dim i As Integer

        Try
            For i = 1 To amount
                waitForMe()
                objRx.SendKeys("[" & command & "]")

                'MsgBox("Check here if we have a RED X")

                waitForMe()
            Next
        Catch ex As Exception
            MsgBox("Experienced an exception on MoveMe():  " & ex.ToString)
        End Try
    End Sub

    Sub MoveMe2(command, r, c)
        Try
            waitForMe()
            objRx.SendKeys("[" & command & "]", r, c)
            waitForMe()
        Catch ex As Exception
            MsgBox("Experienced an exception on MoveMe2():  " & ex.ToString)
        End Try
    End Sub

    Sub TypeMe(value)
        Try
            waitForMe()
            'Enter in the value provided
            objRx.SetText(value)
            waitForMe()
        Catch ex As Exception
            MsgBox("Experienced an exception on TypeMe():  " & ex.ToString)
        End Try
    End Sub

    Sub SettingText(text, row, col)
        Try
            waitForMe()
            objRx.SetText(text, row, col)
            waitForMe()
        Catch ex As Exception
            MsgBox("Experienced an exception on SettingText():  " & ex.ToString)
        End Try
    End Sub


    Sub GoHome()
        Try
            'This Subroutine will continue to check to see what screen we are on and
            'get us back to the "Home" screen (CCT600)

            Dim iCounter
            iCounter = 0

            waitForMe()

            'If Trim(objRx.GetText(19, 2, 11)) = "Press Enter" Then

            Do While Trim(objRx.GetText(1, 2, 6)) <> "CCT600"
                waitForMe()
                MoveMe("pf3", 1)
                waitForMe()

                iCounter = iCounter + 1

                If iCounter > 20 Then   'Just in-case it would get stuck in the loop...I wanted a semi-clean way to get out
                    MsgBox("we are exiting GoHome() Subroutine...We probably encountered an error.")
                    Exit Sub
                End If
            Loop
        Catch ex As Exception
            MsgBox("Error in:  GoHome()  ...  " & ex.ToString)
        End Try
    End Sub

    Sub waitOnMe(intHowLong)
        objRx.Wait(intHowLong)
    End Sub

    Sub waitForMe()
        objWait.WaitForAppAvailable()
        System.Threading.Thread.Sleep(10)
        objWait.WaitForInputReady()
    End Sub

    Sub GetUsername()
        Dim objNet      'This will get the username of the person logged into the PC running this Macro

        Try
            objNet = CreateObject("WScript.NetWork")
            usrNm = objNet.UserName
        Catch ex As Exception
            MsgBox("Experienced an exception on GetUsername():  " & ex.ToString)
        End Try
    End Sub

    Public Sub IsRightScreenName(scrName, row, col, mil)
        Try
            If (objRx.WaitForString(scrName, row, col, mil, True)) Then    'This will wait up to the Milliseconds provided
                'Do Nothing...because we are on the desired screen
            Else
                MsgBox("stop...we have detected that you are not on the expected screen.  Please look into.  scrName is:  " & scrName & " row is:  " & row & " col is:  " & col & " mil is:  " & mil)
            End If
        Catch ex As Exception
            MsgBox("Experienced an exception on IsRightScreenName():  " & ex.ToString)
        End Try
    End Sub

    Public Function ManageSessions()
        Dim intSessions, x, y As Integer

        Try
            intSessions = objMgr.autECLConnList.Count

            '** So...if we have session 1 (a), 2, (b) open and I close session 1 (a)...and now I want to open a new session...
            'the new session will be A...which is 1...that is the one I want to use.
            If intSessions > 0 Then

                For x = 1 To intSessions
                    y = 0

                    If x = 1 And LCase(CStr(objMgr.autECLConnList(x).Name)) <> "a" Then
                        y = 1
                    ElseIf x = 2 And LCase(CStr(objMgr.autECLConnList(x).Name)) <> "b" Then
                        y = 2
                    ElseIf x = 3 And LCase(CStr(objMgr.autECLConnList(x).Name)) <> "c" Then
                        y = 3
                    ElseIf x = 4 And LCase(CStr(objMgr.autECLConnList(x).Name)) <> "d" Then
                        y = 4
                    ElseIf x > 4 Then
                        'shouldn't have more than 5 sessions open...right?!
                        MsgBox("Sorry you have too many RxClaim sessions open." & Chr(13) & "Please close 1 or more and try again.")
                        ManageSessions = 0
                        Exit Function
                    End If

                    If y > 0 Then
                        Exit For
                    End If
                Next

                'If y = 0 Then y = intSessions + 1
                If y = 0 Then y = intSessions
                'If y = 0 Then y = 1                '2/5/16...not sure why this used to work (cuz now it does not work)

            ElseIf intSessions = 0 Then
                y = 1
            Else
                MsgBox("SOMETHING IS WRONG...")
                ManageSessions = 0
                Exit Function
            End If

            ManageSessions = y
        Catch ex As Exception
            ManageSessions = Nothing
            MsgBox("Experienced an exception on ManageSessions():  " & ex.ToString)
        End Try
    End Function

    Function getMyDocs() As String
        Dim WshShell As Object

        Try
            WshShell = CreateObject("WScript.Shell")
            getMyDocs = WshShell.SpecialFolders("Desktop") & "\"
        Catch ex As Exception
            getMyDocs = Nothing
            MsgBox("Experienced an exception on getMyDocs():  " & ex.ToString)
        End Try
    End Function

End Class
