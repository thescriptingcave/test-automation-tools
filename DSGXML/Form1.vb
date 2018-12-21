Imports System.Xml
Imports System.Xml.XPath
Imports Microsoft.Office.Interop.Excel
Imports System
Imports System.IO





Public Class Form1

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button1.Click

        Dim excelapp As New Microsoft.Office.Interop.Excel.Application

        Dim sheetorder As Integer
        Dim screen_names As String
        Dim xmlDoc As New XmlDocument
        Dim nsMgr As XmlNamespaceManager
        Dim selectedNodes As XmlNodeList
        Dim pagename As String = ""
        Dim browsername As String
        Dim childnodes As XmlNode
        Dim objecttype As XmlNode
        Dim objType As String
        Dim objTypeName As String
        Dim pageobjectName As String
        Dim retPage As String = "Wages"
        Dim cnt As Integer = 0
        Dim arrWebObjects() As String = {"Screen Text", "Images", "Text Input", "Radio Button", "Drop-Down", "Check-Box", "Link", "Button"}
        Const xlDescending = 1
        Const xlYes = 0
        Dim workingsheet As Object
        Dim wksheetRange As Object
        Dim TotalRowCount As Integer
        Dim objectTColRange As Object
        Dim CItemNumber As Integer
        Const xlPasteValues = -4163
        Const xlPasteSpecialOperationNone = -4142
        Dim formatrng As Object
        Dim Screentotals As Integer
        Dim excelsheetC As Integer
        Dim delta As Integer
        Dim irow, icol As Integer
        Const xlNone = -4142
        Const xlEdgeRight = 10
        Const xlEdgeBottom = 9
        Const xlEdgeTop = 8
        Const xlEdgeLeft = 7
        Const xlDiagonalDown = 5
        Const xlDiagonalUp = 6
        Const xlContinuous = 1
        Const xlInsideVertical = 11
        Const xlInsideHorizontal = 12
        Const xlThin = 2




        excelapp.Workbooks.Add()
        excelapp.Visible = False

        excelapp.Range("A1").Select()

        xmlDoc.Load(Form3.TextBox1.Text)
        ' xmlDoc.Load("c:\temp\sample.xml")

        nsMgr = New XmlNamespaceManager(xmlDoc.NameTable)
        nsMgr.AddNamespace("qtpRep", "http://www.mercury.com/qtp/ObjectRepository")

        selectedNodes = xmlDoc.SelectNodes("/qtpRep:ObjectRepository/qtpRep:Objects/qtpRep:Object[@Class=""Browser""]", nsMgr)

        Screentotals = ListBox2.Items.Count
        excelsheetC = excelapp.Worksheets.Count

        Form2.ProgressBar1.Maximum = Screentotals
        Form2.ProgressBar1.Step = 1
        Form2.Show()


        '    Debug.Print(Screentotals)

        If (Screentotals > excelsheetC) Then
            delta = (Screentotals - excelsheetC)
            If delta > 0 Then
                For i = 1 To delta
                    excelapp.Worksheets.Add()


                Next
            End If

        End If

        For sheetorder = 0 To ListBox2.Items.Count - 1
            screen_names = ListBox2.Items(sheetorder)
            excelapp.Worksheets(sheetorder + 1).activate()

            For Each selectedNode As XmlNode In selectedNodes

                browsername = selectedNode.Attributes(1).Value
                'Debug.Print(browsername)

                childnodes = selectedNode.ChildNodes.Item(4)

                For Each page As XmlNode In childnodes

                    pagename = page.Attributes(1).Value

                    If pagename = screen_names Then
                        pagename = page.Attributes(1).Value
                        pageobjectName = page.Attributes(0).Value

                        Debug.Print(pagename)
                        Debug.Print(pageobjectName)
                        Debug.Print("")

                        objecttype = page.ChildNodes.Item(4)
                        For Each obj As XmlNode In objecttype

                            objTypeName = obj.Attributes(1).Value
                            objType = obj.Attributes(0).Value

                            Select Case objType

                                Case "WebElement"
                                    objType = "Screen Text"

                                Case "WebEdit"
                                    objType = "Text Input"

                                Case "WebList"
                                    objType = "Drop-Down"

                                Case "WebRadioGroup"
                                    objType = "Radio Button"

                                Case "WebCheckBox"
                                    objType = "Check-Box"

                                Case "WebButton"
                                    objType = "Button"

                            End Select

                            '  Debug.Print(objType)

                            ' Debug.Print(browsername & " " & pagename & " " & objTypeName & "  " & objType)
                            '/Vertical order
                            excelapp.Range("A1").Offset(cnt, 0).Value = browsername
                            excelapp.Range("A1").Offset(cnt, 1).Value = pagename
                            excelapp.Range("A1").Offset(cnt, 2).Value = objTypeName
                            excelapp.Range("A1").Offset(cnt, 3).Value = objType

                            '/Horizontal order
                            'excelapp.Range("A1").Offset(0, cnt).Value = browsername
                            'excelapp.Range("A1").Offset(1, cnt).Value = pagename
                            'excelapp.Range("A1").Offset(2, cnt).Value = objTypeName
                            'excelapp.Range("A1").Offset(3, cnt).Value = objType

                            Select Case objType
                                Case "Screen Text"
                                    'excelapp.Range("A1").Offset(4, cnt).Value = "VERIFY"
                                    excelapp.Range("A1").Offset(cnt, 4).Value = "VERIFY"
                            End Select

                            cnt = cnt + 1

                        Next

                        Workingsheet = excelapp.Worksheets(sheetorder + 1)
                        wksheetRange = workingsheet.UsedRange
                        TotalRowCount = WorkingSheet.UsedRange.Rows.count
                        objectTColRange = WorkingSheet.Range("D1:" & "D" & TotalRowCount)
                        excelapp.Application.AddCustomList(arrWebObjects)
                        cItemNumber = excelapp.Application.GetCustomListNum(arrWebObjects)
                        WorkingSheet.sort.sortFields.Clear()
                        wksheetRange.Sort(objectTColRange, xlDescending, , , , , , xlYes, cItemNumber + 1, False)
                        wksheetRange.Columns.Autofit()
                        excelapp.Application.DeleteCustomList(CItemNumber)

                        wksheetRange.copy()
                        workingsheet.Range("G1").PasteSpecial(xlPasteValues, xlPasteSpecialOperationNone, , True)

                        workingsheet.columns("A:F").delete()
                        workingsheet.cells.Numberformat = "@"

                        formatrng = workingsheet.UsedRange
                        irow = formatrng.Rows.Count
                        icol = formatrng.Columns.count
                        If irow = 4 Then
                            workingsheet.Range("A5").value = "1"
                            formatrng = workingsheet.UsedRange
                            'formatRng = sheetobject.Range(formatRng.Cells(1, 1), formatRng.Cells(irow, icol))
                            formatrng.interior.colorIndex = 36
                            formatrng.Columns.Autofit()
                            With formatrng.Borders(xlEdgeLeft)
                                .LineStyle = xlContinuous
                                .ColorIndex = 0
                                .TintAndShade = 0
                                .Weight = xlThin
                            End With
                            With formatrng.Borders(xlEdgeTop)
                                .LineStyle = xlContinuous
                                .ColorIndex = 0
                                .TintAndShade = 0
                                .Weight = xlThin
                            End With
                            With formatrng.Borders(xlEdgeBottom)
                                .LineStyle = xlContinuous
                                .ColorIndex = 0
                                .TintAndShade = 0
                                .Weight = xlThin
                            End With
                            With formatrng.Borders(xlEdgeRight)
                                .LineStyle = xlContinuous
                                .ColorIndex = 0
                                .TintAndShade = 0
                                .Weight = xlThin
                            End With
                            With formatrng.Borders(xlInsideVertical)
                                .LineStyle = xlContinuous
                                .ColorIndex = 0
                                .TintAndShade = 0
                                .Weight = xlThin
                            End With
                            With formatrng.Borders(xlInsideHorizontal)
                                .LineStyle = xlContinuous
                                .ColorIndex = 0
                                .TintAndShade = 0
                                .Weight = xlThin
                            End With

                            workingsheet.Range("A5").value = ""
                            '  formatrng = Nothing
                            'workingsheet = Nothing
                        Else

                            formatrng.Columns.AutoFit()
                            formatrng.Borders(xlDiagonalDown).LineStyle = xlNone
                            formatrng.Borders(xlDiagonalUp).LineStyle = xlNone
                            With formatrng.Borders(xlEdgeLeft)
                                .LineStyle = xlContinuous
                                .ColorIndex = 0
                                .TintAndShade = 0
                                .Weight = xlThin
                            End With
                            With formatrng.Borders(xlEdgeTop)
                                .LineStyle = xlContinuous
                                .ColorIndex = 0
                                .TintAndShade = 0
                                .Weight = xlThin
                            End With
                            With formatrng.Borders(xlEdgeBottom)
                                .LineStyle = xlContinuous
                                .ColorIndex = 0
                                .TintAndShade = 0
                                .Weight = xlThin
                            End With
                            With formatrng.Borders(xlEdgeRight)
                                .LineStyle = xlContinuous
                                .ColorIndex = 0
                                .TintAndShade = 0
                                .Weight = xlThin
                            End With
                            With formatrng.Borders(xlInsideVertical)
                                .LineStyle = xlContinuous
                                .ColorIndex = 0
                                .TintAndShade = 0
                                .Weight = xlThin
                            End With
                            With formatrng.Borders(xlInsideHorizontal)
                                .LineStyle = xlContinuous
                                .ColorIndex = 0
                                .TintAndShade = 0
                                .Weight = xlThin
                            End With
                            formatrng.interior.colorIndex = 36
                      
                        End If


                        'formatrng.interior.colorIndex = 36
                        'formatrng.columns.autofit()
                        'formatrng.Columns.AutoFit()
                        'formatrng.Borders(xlDiagonalDown).LineStyle = xlNone
                        'formatrng.Borders(xlDiagonalUp).LineStyle = xlNone
                        'With formatrng.Borders(xlEdgeLeft)
                        '    .LineStyle = xlContinuous
                        '    .ColorIndex = 0
                        '    .TintAndShade = 0
                        '    .Weight = xlThin
                        'End With
                        'With formatrng.Borders(xlEdgeTop)
                        '    .LineStyle = xlContinuous
                        '    .ColorIndex = 0
                        '    .TintAndShade = 0
                        '    .Weight = xlThin
                        'End With
                        'With formatrng.Borders(xlEdgeBottom)
                        '    .LineStyle = xlContinuous
                        '    .ColorIndex = 0
                        '    .TintAndShade = 0
                        '    .Weight = xlThin
                        'End With
                        'With formatrng.Borders(xlEdgeRight)
                        '    .LineStyle = xlContinuous
                        '    .ColorIndex = 0
                        '    .TintAndShade = 0
                        '    .Weight = xlThin
                        'End With
                        'With formatrng.Borders(xlInsideVertical)
                        '    .LineStyle = xlContinuous
                        '    .ColorIndex = 0
                        '    .TintAndShade = 0
                        '    .Weight = xlThin
                        'End With
                        'With formatrng.Borders(xlInsideHorizontal)
                        '    .LineStyle = xlContinuous
                        '    .ColorIndex = 0
                        '    .TintAndShade = 0
                        '    .Weight = xlThin
                        'End With

                        'formatrng = Nothing

                        If Len(screen_names) >= 25 Then
                            screen_names = screen_names.Substring(0, 20)
                        Else
                            screen_names = screen_names

                        End If
                        workingsheet.name = screen_names & " " & sheetorder + 1

                        Debug.Print(workingsheet.name)

                        Form2.ProgressBar1.PerformStep()


                        cnt = 0
                    End If

                Next

            Next
        Next


        excelapp.Visible = True

        Me.Close()
        Form2.Close()
        Form3.Close()





    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

        Me.Left = (Screen.PrimaryScreen.WorkingArea.Width - Me.Width) / 2
        Me.Top = (Screen.PrimaryScreen.WorkingArea.Height - Me.Height) / 2

        Dim xmlDoc As New XmlDocument
        Dim nsMgr As XmlNamespaceManager
        Dim selectedNodes As XmlNodeList
        Dim pagename As String = ""
        Dim browsername As String
        Dim childnodes As XmlNode
        ' Dim combinedString As String

        xmlDoc.Load(Form3.TextBox1.Text)

        ' xmlDoc.Load("c:\temp\sample.xml")

        nsMgr = New XmlNamespaceManager(xmlDoc.NameTable)
        nsMgr.AddNamespace("qtpRep", "http://www.mercury.com/qtp/ObjectRepository")
        selectedNodes = xmlDoc.SelectNodes("/qtpRep:ObjectRepository/qtpRep:Objects/qtpRep:Object[@Class=""Browser""]", nsMgr)

        For Each selectedNode As XmlNode In selectedNodes

            browsername = selectedNode.Attributes(1).Value
            ' Debug.Print(browsername)
            childnodes = selectedNode.ChildNodes.Item(4)
            For Each page As XmlNode In childnodes
                pagename = page.Attributes(1).Value
                'pageobjectName = page.Attributes(0).Value
                'Debug.Print(pagename)
                ListBox1.Items.Add(pagename)


            Next
        Next
    End Sub

    Private Sub Button2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button2.Click

        ListBox2.Items.Add((ListBox1.Text))



    End Sub

    Private Sub Button3_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button3.Click
        ListBox2.Items.Remove(ListBox2.Text)
        ListBox2.AllowDrop = True
    End Sub
End Class
