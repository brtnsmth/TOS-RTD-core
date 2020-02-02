Option Explicit On

'  Code Block

' Excel interface
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

' StreanWriter interface
Imports System
Imports System.IO
Imports System.Text


Public Class Form1

#Region "Dims"
    ' Define Excel interface instance
    Private ReadOnly oXL As New Excel.Application
    Private oWB As Excel.Workbook
    Private oSheet As Excel.Worksheet
    Private oRng As Excel.Range


    ' Define Excel source cells
    Private XL_Cell_1, XL_Cell_2, XL_Cell_3, XL_Cell_4, XL_Cell_5 As Excel.Range
    Private XL_Cell_6, XL_Cell_7, XL_Cell_8, XL_Cell_9, XL_Cell_10 As Excel.Range
    Private XL_Cell_11, XL_Cell_12, XL_Cell_13, XL_Cell_14, XL_Cell_15 As Excel.Range
    Private XL_Cell_16, XL_Cell_17, XL_Cell_18, XL_Cell_19, XL_Cell_20 As Excel.Range


    ' Define DB output columns
    Private Data_Col_1, Data_Col_2, Data_Col_3, Data_Col_4, Data_Col_5 As String
    Private Data_Col_6, Data_Col_7, Data_Col_8, Data_Col_9, Data_Col_10 As String
    Private Data_Col_11, Data_Col_12, Data_Col_13, Data_Col_14, Data_Col_15 As String
    Private Data_Col_16, Data_Col_17, Data_Col_18, Data_Col_19, Data_Col_20 As String

    Private tsTimeStamp As DateTime                 ' timestamp for data field
    Private strTimeStamp As String                  ' time string used to create unique data file

    Private boolExcelLoaded As Boolean = False      ' make sure excel is running before asking for data

    Private strDayTime As String                    ' current time as hour minute to test against trigger below
    Private strDayTimeTrigger As String = "1759"    ' look for minute before 6 PM to start new data file

    Private boolNewFileCreated As Boolean = False   ' check to see if file exists for appending or if need to create


    '******************
    ' StreamWriter interface to CSV file
    Private swCSV As StreamWriter

    Private strCSV_FileNameBase As String = "C:\Temp\TOS Import Core\data\TOS data "      ' root of filename for streamwriter
    Private strCSV_FileName As String                                                     ' section of filename built JIT
    Private strCSV_FileNameExtension As String = ".csv"                                   ' defines file as csv format

    ' column headers written at file creation
    ' Private strCSV_Header As String = "Date" & ", " & "/ES" & ", " & "/NQ" & ", " & "/RTY" & ", " & "SPY" & ", " & "QQQ" & ", " & "IWM" & ", " & "AAPL" & ", " & "MSFT" & ", " & "NVDA" & ", " & "XLK" & ", " & "XLF" & ", " & "XLP" & ", " & "XLY" & ", " & "XTN" & ", " & "HYG" & ", " & "***" & ", " & "***" & ", " & "***" & ", " & "***"
    Private strCSV_Header As String = "Date" & ", " & "/ES" & ", " & "/NQ" & ", " & "/RTY" & ", " & "SPY" & ", " & "QQQ" & ", " & "IWM" & ", " & "AAPL" & ", " & "MSFT" & ", " & "NVDA" & ", " & "XLK" & ", " & "XLF" & ", " & "XLP" & ", " & "XLY" & ", " & "XTN" & ", " & "HYG" & ", " & "col 16" & ", " & "col 17" & ", " & "col 18" & ", " & "col 19" & ", " & "col 20"

    ' data string aligned with headers, built in save routine
    Private strCSV_Data As String


#End Region

#Region "Form Load / Close"


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ' ***************************************
        ' Initialize Excel workbook
        oXL.Visible = False
        oWB = oXL.Workbooks.Add
        oSheet = oWB.ActiveSheet

        ' run delay so Excel  can load 
        tmrExcelLoad.Interval = 10000
        tmrExcelLoad.Start()
        If tmrExcelLoad.Enabled = True Then Label1.Text = "Excel Load timer started"


        ' ***************************************
        ' Load variables and Excel formulas
        Call Define_Cell_Ranges()
        Call Define_Cell_Formulas()


        '****************************************
        ' start StreamWriter
        Call CreateNewCSVFile()


    End Sub


    Private Sub Define_Cell_Ranges()

        ' define cell ranges
        XL_Cell_1 = oSheet.Cells(1, 1)
        XL_Cell_2 = oSheet.Cells(2, 1)
        XL_Cell_3 = oSheet.Cells(3, 1)
        XL_Cell_4 = oSheet.Cells(4, 1)
        XL_Cell_5 = oSheet.Cells(5, 1)
        XL_Cell_6 = oSheet.Cells(6, 1)
        XL_Cell_7 = oSheet.Cells(7, 1)
        XL_Cell_8 = oSheet.Cells(8, 1)
        XL_Cell_9 = oSheet.Cells(9, 1)
        XL_Cell_10 = oSheet.Cells(10, 1)

        XL_Cell_11 = oSheet.Cells(11, 1)
        XL_Cell_12 = oSheet.Cells(12, 1)
        XL_Cell_13 = oSheet.Cells(13, 1)
        XL_Cell_14 = oSheet.Cells(14, 1)
        XL_Cell_15 = oSheet.Cells(15, 1)
        XL_Cell_16 = oSheet.Cells(16, 1)
        XL_Cell_17 = oSheet.Cells(17, 1)
        XL_Cell_18 = oSheet.Cells(18, 1)
        XL_Cell_19 = oSheet.Cells(19, 1)
        XL_Cell_20 = oSheet.Cells(20, 1)

    End Sub


    Private Sub Define_Cell_Formulas()

        ' define cell formulas
        XL_Cell_1.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""/ES:XCME"" )"
        XL_Cell_2.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""/NQ:XCME"" )"
        XL_Cell_3.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""/RTY:XCME"" )"
        XL_Cell_4.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""SPY"" )"
        XL_Cell_5.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""QQQ"" )"
        XL_Cell_6.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""IWM"" )"
        XL_Cell_7.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""AAPL"" )"
        XL_Cell_8.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""MSFT"" )"
        XL_Cell_9.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""NVDA"" )"
        XL_Cell_10.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""XLK"" )"

        XL_Cell_11.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""XLF"" )"
        XL_Cell_12.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""XLP"" )"
        XL_Cell_13.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""XLY"" )"
        XL_Cell_14.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""XTN"" )"
        XL_Cell_15.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""HYG"" )"
        'XL_Cell_16.Formula = "=RTD(""TOS.RTD"", , ""LAST"", "" X "" )"
        'XL_Cell_17.Formula = "=RTD(""TOS.RTD"", , ""LAST"", "" X "" )"
        'XL_Cell_18.Formula = "=RTD(""TOS.RTD"", , ""LAST"", "" X "" )"
        'XL_Cell_19.Formula = "=RTD(""TOS.RTD"", , ""LAST"", "" X "" )"
        'XL_Cell_20.Formula = "=RTD(""TOS.RTD"", , ""LAST"", "" X "" )"

    End Sub

    Private Sub CreateNewCSVFile()

        '****************************************
        ' build timestamp to embed in filename
        strTimeStamp = Format(Now(), "yyyy MM dd HH mm ss")

        Try
            ' build filename
            strCSV_FileName = strCSV_FileNameBase & strTimeStamp & strCSV_FileNameExtension
            ' Open NEW StreamWriter
            swCSV = My.Computer.FileSystem.OpenTextFileWriter(strCSV_FileName, True)
            ' write column headers to first record
            swCSV.WriteLine(strCSV_Header)
            swCSV.Close()
        Catch ex As IOException
            MsgBox(ex.ToString)
        End Try


        boolNewFileCreated = True

    End Sub


    Private Sub Form1_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing

        oRng = Nothing
        oSheet = Nothing
        oWB = Nothing
        oXL.Quit()

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL)

        ' TODO -find way to close Excel without external Save? msgbox
    End Sub

#End Region


#Region "Timer events"

    Private Sub tmrExcelLoad_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrExcelLoad.Tick

        '****************************************
        ' Excel load delay before connecting to TOS RTD is done 
        Label1.Text = "ExcelLoad timer done"
        tmrExcelLoad.Enabled = False

        '****************************************
        ' Test VB - Excel RTD connection
        Call Test_RTD_Connection()

        '****************************************
        ' VB - Excel RTD connection OK
        ' start main looping
        Timer1.Interval = 3000
        Timer1.Start()

    End Sub

    Private Sub Test_RTD_Connection()

        Try
            Label1.Text = "Testing RTD connection"
            Dim rg As Excel.Range = oSheet.Cells(1, 1)

            rg.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""/ES:XCME"" )"

            If CStr(rg.Value) = "" Then
                MessageBox.Show("No connection to XL TOS.RTD server." & vbCrLf & "Exit pgr, start TOS, restart pgr.")
            Else
                Label1.Text = "XL TOS.RTD connection tested OK"
            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message)

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL)

        End Try


    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Timer1.Enabled = False

        '***************************************
        ' Check for new futures trading day - starts at 6PM EST day before market day
        strDayTime = Format(Now(), "HHmm")

        ' If triggered, create file for new trading day's data
        If strDayTime = strDayTimeTrigger And boolNewFileCreated = False Then
            Call CreateNewCSVFile()
        End If

        ' once trigger cleared, clear boolNewFileCreated
        If strDayTime <> strDayTimeTrigger Then boolNewFileCreated = False


        '***************************************
        ' get and save new data
        Call Get_Values()
        Call CSV_Save()

        Timer1.Enabled = True

    End Sub



#End Region


#Region "get data"

    Private Sub Get_Values()

        Try

            ' Load values returned from RTDServer
            Data_Col_1 = CStr(XL_Cell_1.Value)
            Data_Col_2 = CStr(XL_Cell_2.Value)
            Data_Col_3 = CStr(XL_Cell_3.Value)
            Data_Col_4 = CStr(XL_Cell_4.Value)
            Data_Col_5 = CStr(XL_Cell_5.Value)
            Data_Col_6 = CStr(XL_Cell_6.Value)
            Data_Col_7 = CStr(XL_Cell_7.Value)
            Data_Col_8 = CStr(XL_Cell_8.Value)
            Data_Col_9 = CStr(XL_Cell_9.Value)
            Data_Col_10 = CStr(XL_Cell_10.Value)

            Data_Col_11 = CStr(XL_Cell_11.Value)
            Data_Col_12 = CStr(XL_Cell_12.Value)
            Data_Col_13 = CStr(XL_Cell_13.Value)
            Data_Col_14 = CStr(XL_Cell_14.Value)
            Data_Col_15 = CStr(XL_Cell_15.Value)
            Data_Col_16 = CStr(XL_Cell_16.Value)
            Data_Col_17 = CStr(XL_Cell_17.Value)
            Data_Col_18 = CStr(XL_Cell_18.Value)
            Data_Col_19 = CStr(XL_Cell_19.Value)
            Data_Col_20 = CStr(XL_Cell_20.Value)

        Catch ex As Exception
            Debug.Print(ex.Message)
            Debug.Print("exit exception at " & Now())
        End Try

    End Sub


#End Region


#Region "CSV saves"

    Private Sub CSV_Save()

        ' get current time
        tsTimeStamp = Now()

        '****************************************
        ' build data string
        strCSV_Data = CStr(tsTimeStamp) & ", " & Data_Col_1 & ", " & Data_Col_2 & ", " & Data_Col_3 & ", " & Data_Col_4 & ", " & Data_Col_5 & ", " & Data_Col_6 & ", " & Data_Col_7 & ", " & Data_Col_8 & ", " & Data_Col_9 & ", " & Data_Col_10 & ", " & Data_Col_11 & ", " & Data_Col_12 & ", " & Data_Col_13 & ", " & Data_Col_14 & ", " & Data_Col_15 & ", " & Data_Col_16 & ", " & Data_Col_17 & ", " & Data_Col_18 & ", " & Data_Col_19 & ", " & Data_Col_20

        ' write data 

        '****************************************
        ' Open NEW StreamWriter and write column headers to first line.

        Try
            swCSV = My.Computer.FileSystem.OpenTextFileWriter(strCSV_FileName, True)
            swCSV.WriteLine(strCSV_Data)
            swCSV.Close()

            ' Debug.Print("CSV data write success @ " & Now())
            Label1.Text = "CSV data write success @ " & Now()

        Catch ex As IOException
            MsgBox(ex.ToString)

            System.Windows.Forms.MessageBox.Show(ex.Message)
            Label1.Text = "CSV data write failed @ " & Now()
        End Try


    End Sub
#End Region




End Class