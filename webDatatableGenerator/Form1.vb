Imports Excel = Microsoft.Office.Interop.Excel

Public Class form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim myexcel As Excel.Application
        Dim wkbk As Object



        myexcel = New Excel.Application
        wkbk = myexcel.Workbooks.Add()
        myexcel.DisplayAlerts = False
        myexcel.ActiveWorkbook.SaveAs("c:\datatables\datatable.xls", 56)
        myexcel.Quit()
        myexcel = Nothing
        wkbk = Nothing


        OpenFileDialog1.Title = "Please Select a Datasheet"
        OpenFileDialog1.InitialDirectory = "C:\datatables"
        OpenFileDialog1.ShowDialog()



    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub OpenFileDialog1_FileOk(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        Dim strm As System.IO.Stream

        strm = OpenFileDialog1.OpenFile()
        TextBox1.Text = OpenFileDialog1.FileName.ToString()

        If Not (strm Is Nothing) Then
            'insert code to read the file data
            strm.Close()

        End If


    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Dim oexcel As Excel.Application
        Dim oSource As Excel.Workbook
        Dim oDest As Excel.Workbook
        Dim oD As Object
        Dim WorksheetsCount As Integer
        Dim pos As Integer
        Dim row As Integer
        Dim i As Integer
        ' Dim activesheet As Excel.Worksheet
        Dim working_sheet As Excel.Worksheet
        Dim cells As Excel.Range
        Dim usedColumnCount As Integer
        Dim usedrowsCount As Integer
        Dim b, p, o, ot, a, Data As String
        Dim maximum As Integer
        Dim Index As Integer
        Dim cnt As Integer
        Dim bs As Integer = 0



        Dim objRange As Excel.Range
        Dim xlContinuous As Integer

        xlContinuous = 1

        oexcel = New Excel.Application
        oSource = oexcel.Workbooks.Open(TextBox1.Text)

        oDest = oexcel.Workbooks.Open("c:\datatables\datatable.xls")

        oexcel.Visible = True

        oD = oDest.Worksheets(1)




        WorksheetsCount = oSource.Worksheets.Count
        'this code adds an array so that we can find the maximum row of the X number of sheets
        Dim arraytotal(WorksheetsCount - 1) As Integer
        Dim arraycoltotal(WorksheetsCount - 1) As Integer
        Dim maximumCol, index2, cnt2, maximum2, celltotals As Integer


        For t = 0 To WorksheetsCount - 1
            arraytotal(t) = oSource.Worksheets(t + 1).UsedRange.Rows.Count()
            arraycoltotal(t) = oSource.Worksheets(t + 1).UsedRange.columns.Count()
        Next

        maximum = arraytotal(0)
        ' first value of the array
        Index2 = 0
        For cnt = 0 To arraytotal.Length - 1
            If arraytotal(cnt) > maximum Then
                maximum = arraytotal(cnt)
                Index = cnt
            End If
        Next

        maximumCol = arraycoltotal(0)
        ' first value of the array
        Index = 0
        For cnt2 = 0 To arraycoltotal.Length - 1
            If arraycoltotal(cnt2) > maximumCol Then
                maximum2 = arraycoltotal(cnt2)
                index2 = cnt2
            End If
        Next

        celltotals = maximum * maximum2

        '  MessageBox.Show(celltotals)

        oD.Range("A1:IV65536").numberformat = "@"

        pos = 1 ' this is the row datatable_row_positionition in the datatable
        oD.Cells(pos, 1).Value = "Browser"
        oD.Cells(pos, 2).Value = "Page"
        oD.Cells(pos, 3).Value = "Object"
        oD.Cells(pos, 4).Value = "ObjectType"
        oD.Cells(pos, 5).Value = "Action"
        oD.Cells(pos, 6).Value = "Data"
        oD.Cells(pos, 7).Value = "Step_Name"



        row = 6 'this is where the rows start in the datasheet
        i = 1 ' this is the worksheet index variable

        Do
            working_sheet = oSource.Worksheets(i)
            cells = working_sheet.UsedRange
            usedColumnCount = working_sheet.UsedRange.Columns.Count
            'usedrowsCount = working_sheet.UsedRange.Rows.Count
            usedrowsCount = maximum


            For col = 1 To usedColumnCount
            

                Data = working_sheet.Cells(row, col).value

                'cells(row, col).interior.ColorIndex = 17	
                If Data <> "" Then
                    bs = bs + 1


                    b = cells(1, col).Value
                    p = cells(2, col).Value
                    o = cells(3, col).Value
                    ot = cells(4, col).Value
                    a = cells(5, col).Value
                    'pos keeps track of  rowpos of datatable
                    pos = pos + 1
                    oD.Cells(pos, 1).Value = b
                    oD.Cells(pos, 2).Value = p
                    oD.Cells(pos, 3).Value = o
                    oD.Cells(pos, 4).Value = ot
                    oD.Cells(pos, 5).Value = UCase(a)

                    ' Debug.Print(a)

                    Select Case ot
                        Case "Button"
                            oD.Cells(pos, 6).Value = ""
                        Case "Link"
                            oD.Cells(pos, 6).Value = ""

                            'Case "Screen Text"
                            '  oD.Cells(pos, 6).Value = Data
                        Case Else
                            oD.Cells(pos, 6).Value = Data
                    End Select

                    If UCase(Trim(a)) = "CLOSEBROWSER" Then
                        oD.Cells(pos, 5).Value = "CLOSEBROWSER"
                        oD.Cells(pos, 6).Value = ""
                    End If


                    If UCase(Trim(a)) = "VERIFY" Then
                        oD.Cells(pos, 7).Value = a & " " & o
                    End If
                Else

                End If
            Next


            If i >= WorksheetsCount Then
                i = 1
            Else
                i = i + 1
            End If
            If i = 1 Then
                row = row + 1
            End If

        Loop Until row > usedrowsCount


        objRange = oD.UsedRange
        objRange.Font.Name = "System"
        objRange.Borders.LineStyle = xlContinuous
        objRange.Columns.AutoFit()

        oDest.Save()
        oDest.Close()
        oSource.Close()
        oexcel.Quit()
        oexcel = Nothing

        ReleaseComObject(oDest)
        ReleaseComObject(oSource)
        ReleaseComObject(oexcel)
        GC.Collect()
        GC.WaitForPendingFinalizers()

        MsgBox("Datatable.xls can be found in C:\datatables")
        Me.Close()


    End Sub
    Friend Sub ReleaseComObject(ByVal o As Object)
        Try
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(o)
        Catch
        Finally
            If o IsNot Nothing Then o = Nothing
        End Try
    End Sub


End Class
