Imports oExcel = Microsoft.Office.Interop.Excel

Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "Please Select DataSheet .."
        OpenFileDialog1.InitialDirectory = "C:\DataTables"
        OpenFileDialog1.ShowDialog()

    End Sub

    Private Sub OpenFileDialog1_FileOk(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk


        TextBox1.Text = OpenFileDialog1.FileName.ToString()


    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        OpenFileDialog2.Title = "Please Select Datatable .."
        OpenFileDialog2.InitialDirectory = "C:\DataTables"
        OpenFileDialog2.ShowDialog()
    End Sub

    Private Sub OpenFileDialog2_FileOk(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog2.FileOk
        TextBox2.Text = OpenFileDialog2.FileName.ToString()

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Dim oExcel As New Excel.Application
        Dim oSource As Excel.Workbook
        Dim oDest As Excel.Workbook
        Dim counter, markedForUpdate As Integer
        Dim oS As Object
        Dim t, maximum, index, ds, row, i, WorksheetsCount, usedColumnCount, usedRowsCount, datasheet_counter As Integer
        Dim activesheet As Excel.Worksheet
        Dim cells As Excel.Range
        Dim Data As String
        Dim datasheet_items As New Dictionary(Of Integer, String)
        Dim dataTableValues As New Dictionary(Of Integer, String)
        Dim dataTableValuesTotals As Integer


        'DataTable
        oSource = oExcel.Workbooks.Open(TextBox2.Text)
        'DataSheet
        oDest = oExcel.Workbooks.Open(TextBox1.Text)

        ' RemoveAllUnwantedRowsFromPostBatchFile(TextBox2.Text)

        oS = oSource.Worksheets(1)
        oExcel.Visible = False

        dataTableValues = getDataFromDataTable(oS)
        dataTableValuesTotals = dataTableValues.Count
        ProgressBar1.Maximum = dataTableValuesTotals
        ProgressBar1.Step = 1
        counter = 1

        WorksheetsCount = oDest.Worksheets.count
        ds = 1
        row = 6 'web

        i = 1

        Dim arraytotal(WorksheetsCount - 1) As Integer

        For t = 0 To WorksheetsCount - 1
            arraytotal(t) = oDest.Worksheets(t + 1).UsedRange.Rows.Count()

        Next

        maximum = arraytotal(0)
        ' first value of the array
        Index = 0
        For cnt = 0 To arraytotal.Length - 1
            If arraytotal(cnt) > maximum Then
                maximum = arraytotal(cnt)
                Index = cnt
            End If
        Next

        Do
            activesheet = oDest.Worksheets(i)
            cells = activesheet.UsedRange
            usedColumnCount = activesheet.UsedRange.Columns.Count
            'usedRowsCount = activesheet.UsedRange.Rows.Count
            usedRowsCount = maximum
            For col = 1 To usedColumnCount
                Data = RTrim(UCase(cells(row, col).Value))

                datasheet_counter = datasheet_counter + 1

                If CStr(cells(row, col).value) <> "" Then
                    ds = ds + 1
                    ProgressBar1.PerformStep()
                    If dataTableValues.ContainsKey(ds) Then

                        If CStr(dataTableValues.Item(ds)) <> "" Then
                            ProgressBar1.PerformStep()
                            cells(row, col).Interior.ColorIndex = 4
                            cells(row, col).Value = dataTableValues.Item(ds)
                            markedForUpdate = markedForUpdate + 1
                        End If
                    End If
                End If

            Next

            If i >= WorksheetsCount Then 'go until last sheet for each row
                i = 1
            Else 'increament sheet number until it reaches the last sheet
                i = i + 1
            End If
            If i = 1 Then 'increament row number after the last sheet
                row = row + 1
            End If
        Loop Until row > usedRowsCount 'loop until all rows are executed


        Debug.Print(datasheet_items.Count)
        Debug.Print(ds - 1)

        Label5.Text = ds - 1
        Label6.Text = markedForUpdate

        oDest.Save()
        oDest.Close()
        oSource.Close()
        oExcel.Quit()

        MessageBox.Show("Update Complete")
        oExcel = Nothing
        oDest = Nothing
        oSource = Nothing

        ReleaseComObject(oDest)
        ReleaseComObject(oSource)
        ReleaseComObject(oExcel)
        ReleaseComObject(oExcel)
        ReleaseComObject(oS)
        ReleaseComObject(activesheet)
        ReleaseComObject(cells)

        GC.Collect()
        GC.WaitForPendingFinalizers()


    End Sub

    Friend Sub ReleaseComObject(ByVal o As Object)
        Try
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(o)
        Catch
        Finally
            If o IsNot Nothing Then o = Nothing
        End Try
    End Sub

    Public Function getDataFromDataTable(ByVal os As Worksheet) As Dictionary(Of Integer, String)

        Dim dataTableDict As New Dictionary(Of Integer, String)
        Dim cnt As Integer
        Dim rowTotal As Integer
        Dim mData As String

        cnt = 1

        rowTotal = os.UsedRange.Rows.Count

        Do Until cnt = rowTotal

            mData = CStr(os.Cells(cnt, 8).value)
            dataTableDict.Add(cnt, mData)
            cnt = cnt + 1

        Loop
        Label8.Text = rowTotal - 1
        Return dataTableDict
    End Function

    Public Sub RemoveAllUnwantedRowsFromPostBatchFile(ByVal DataTableName)
        Dim x As New Excel.Application
        Dim y As Excel.Workbook
        Dim z As Object
        Dim lastRow, row As Integer
        Dim RangeForDeletion As Range

        y = x.Workbooks.Open(DataTableName)
        z = y.Worksheets(1)

        LastRow = z.UsedRange.Rows.Count

        If lastRow = "65536" Then
            row = 1
            Do
                If z.Cells(row, 1).value = "" Then
                    Exit Do
                End If
                row = row + 1
            Loop


            RangeForDeletion = x.Range("A" & row & ":" & "A" & LastRow)
            RangeForDeletion.EntireRow.Delete()

            y.Save()
            y.Close()
            x.Quit()

            x = Nothing
            y = Nothing
            z = Nothing

            ReleaseComObject(x)
            ReleaseComObject(y)
            ReleaseComObject(z)
            GC.Collect()
            GC.WaitForPendingFinalizers()

        Else

            y.close()
            x.quit()
            x = Nothing
            y = Nothing
            z = Nothing

            ReleaseComObject(x)
            ReleaseComObject(y)
            ReleaseComObject(z)
            GC.Collect()
            GC.WaitForPendingFinalizers()

        End If





    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    

    End Sub
End Class
