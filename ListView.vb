Imports System.Drawing
Imports System.Drawing.Color
Public Class ListView

    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)

    Private strlvwName As String
    Private WithEvents lvwColumSorter As New ListViewColumnSorter()
    Private Sub ErrorMessage_ListViewColumnSorter(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String) Handles lvwColumSorter.ErrorMessage
        RaiseEvent ErrorMessage(errDesc, errNo, errTrace)
    End Sub
    Public Sub New()
    End Sub
    Public Sub StandardHeader(lvw As Windows.Forms.ListView)


        With lvw
            .OwnerDraw = True
            '    .DrawColumnHeader = New DrawListViewColumnHeaderEventHandler((sender, e) >= headerDraw(sender, e, System.Drawing.Color.Black, System.Drawing.Color.White)
            ')
            '    .DrawItem += New DrawListViewItemEventHandler(bodyDraw)
        End With




    End Sub
    Public Sub CopyCheckedRow(lvw_from As System.Windows.Forms.ListView, lvw_to As System.Windows.Forms.ListView)
        On Error GoTo Err

        Dim lvw_from_col As Integer = lvw_from.Columns.Count
        Dim lvw_from_row As Integer = lvw_from.Items.Count
        Dim lvw_to_col As Integer = lvw_to.Columns.Count
        Dim lvw_to_row As Integer = lvw_to.Items.Count
        Dim ID As String
        Dim SubItem As String
        Dim lViewItem As System.Windows.Forms.ListViewItem
        'check is both are the same cols 
        If lvw_from_col = lvw_to_col Then

            If lvw_from_row > 0 Then
                For i As Integer = 0 To lvw_from_row - 1
                    If lvw_from.Items(i).Checked = True Then
                        ID = lvw_from.Items(i).SubItems(0).Text
                        If CheckForValue(lvw_to, ID, 0) = False Then
                            lViewItem = New Windows.Forms.ListViewItem(ID)
                            For c As Integer = 1 To lvw_from_col - 1
                                SubItem = lvw_from.Items(i).SubItems(c).Text
                                lViewItem.SubItems.Add(SubItem)
                            Next
                            lvw_to.Items.Add(lViewItem)
                        End If
                    End If
                Next
            End If
        Else
            Dim strMessage As String = "The " + lvw_from.Name + " and " + lvw_to.Name + " listviews columns are the same " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
            RaiseEvent ErrorMessage("Columns counts are the same", 0, strMessage)
        End If



        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub

    Public Sub MoveSelectedRow(lvw_from As System.Windows.Forms.ListView, lvw_to As System.Windows.Forms.ListView)
        On Error GoTo Err

        Dim lvw_from_col As Integer = lvw_from.Columns.Count
        Dim lvw_from_row As Integer = lvw_from.Items.Count
        Dim lvw_to_col As Integer = lvw_to.Columns.Count
        Dim lvw_to_row As Integer = lvw_to.Items.Count
        Dim ID As String
        Dim SubItem As String
        Dim Row As Integer = 0
        Dim lvm As System.Windows.Forms.ListViewItem
        'check is both are the same cols 
        If lvw_from_col = lvw_to_col Then

            If lvw_from_row > 0 Then
                Do Until Row >= lvw_from_row
                    If lvw_from.Items(Row).Selected = True Then
                        ID = lvw_from.Items(Row).SubItems(0).Text
                        'Move the row from to
                        lvm = New Windows.Forms.ListViewItem(ID)
                        For c As Integer = 1 To lvw_from_col - 1
                            SubItem = lvw_from.Items(Row).SubItems(c).Text
                            lvm.SubItems.Add(SubItem)
                        Next
                        'Remove the selected row
                        lvw_to.Items.Add(lvm)
                        lvw_from.Items(Row).Remove()
                        lvw_from_row = lvw_from.Items.Count
                    End If
                    Row = Row + 1
                Loop
            End If



        Else
            Dim strMessage As String = "The " + lvw_from.Name + " and " + lvw_to.Name + " listviews columns are the same " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
            RaiseEvent ErrorMessage("Columns counts are the same", 0, strMessage)
        End If



        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
    '    Public Function GetValueFromColumn(lvw As Windows.Forms.ListView, selectIndex As Integer) As ArrayList

    '        Dim lngID As Integer
    '        Dim lngCol As Integer = lvw.Columns.Count
    '        Dim lngRow As Integer = lvw.Items.Count
    '        Dim strvalue As String
    '        Dim arr As New ArrayList


    '        On Error GoTo Err
    '        For c As Integer = 0 To lngCol - 1
    '            arr.Add(lvw.Items(selectIndex).SubItems(c).Text)
    '        Next

    '        Return arr

    '        Exit Function

    'Err:
    '        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
    '        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    '    End Function
    Public Function GetRowValues(lvw As Windows.Forms.ListView, selectIndex As Integer) As ArrayList

        Dim lngID As Integer
        Dim lngCol As Integer = lvw.Columns.Count
        Dim lngRow As Integer = lvw.Items.Count
        Dim strvalue As String
        Dim arr As New ArrayList


        On Error GoTo Err
        For c As Integer = 0 To lngCol - 1
            arr.Add(lvw.Items(selectIndex).SubItems(c).Text)
        Next

        Return arr

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Sub MoveAllRows(lvw_from As System.Windows.Forms.ListView, lvw_to As System.Windows.Forms.ListView)
        On Error GoTo Err

        Dim lvw_from_col As Integer = lvw_from.Columns.Count
        Dim lvw_from_row As Integer = lvw_from.Items.Count
        Dim lvw_to_col As Integer = lvw_to.Columns.Count
        Dim lvw_to_row As Integer = lvw_to.Items.Count
        Dim ID As String
        Dim SubItem As String

        Dim lvm As System.Windows.Forms.ListViewItem
        'check is both are the same cols 
        If lvw_from_col = lvw_to_col Then

            If lvw_from_row > 0 Then
                For Row As Integer = 0 To lvw_from_row - 1
                    ID = lvw_from.Items(Row).SubItems(0).Text
                    lvm = New Windows.Forms.ListViewItem(ID)
                    For c As Integer = 1 To lvw_from_col - 1
                        SubItem = lvw_from.Items(Row).SubItems(c).Text
                        lvm.SubItems.Add(SubItem)
                    Next
                    lvw_to.Items.Add(lvm)
                Next

                lvw_from.Items.Clear()
            End If
        Else
            Dim strMessage As String = "The " + lvw_from.Name + " and " + lvw_to.Name + " listviews columns are the same " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
            RaiseEvent ErrorMessage("Columns counts are the same", 0, strMessage)
        End If



        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
    Public Sub Load(ByRef lvw As System.Windows.Forms.ListView, ByRef ds As DataSet)

        On Error GoTo Err


        Dim lcols As Integer = lvw.Columns.Count
        Dim lrows As Integer = ds.Tables(0).Rows.Count
        Dim lViewItem As System.Windows.Forms.ListViewItem
        Dim r As Integer
        Dim colTag As String
        Dim dsField As String
        Dim dataType As String
        Dim CheckColBoolean As Boolean
        lvw.Items.Clear()
        If lrows > 0 Then

            For Each rows In ds.Tables(0).Rows
                lViewItem = New Windows.Forms.ListViewItem(ds.Tables(0).Rows(r).Item(0).ToString())
                For c As Integer = 1 To lcols - 1
                    dataType = rows.Table.Columns(c).DataType().ToString
                    ' colTag = lvw.Columns(c).Tag.ToString
                    If Not System.Convert.IsDBNull(ds.Tables(0).Rows(r).Item(c)) Then

                        Select Case dataType.ToString
                            Case "System.Bit"
                                If ds.Tables(0).Rows(r).Item(c) = 0 Then
                                    lViewItem.SubItems.Add("No")
                                Else
                                    lViewItem.SubItems.Add("Yes")
                                End If
                            Case "System.Boolean"
                                If ds.Tables(0).Rows(r).Item(c) = True Then
                                    lViewItem.SubItems.Add("Yes")
                                Else
                                    lViewItem.SubItems.Add("No")
                                End If
                            Case "System.Decimal"
                                lViewItem.SubItems.Add(CDbl(ds.Tables(0).Rows(r).Item(c)))
                            Case Else
                                lViewItem.SubItems.Add(ds.Tables(0).Rows(r).Item(c))
                        End Select

                    Else
                        lViewItem.SubItems.Add("")
                    End If
                Next
                lvw.Items.Add(lViewItem)
                r = r + 1
            Next
        End If

        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
    Public Sub LoadCheckBox(ByRef lvw As System.Windows.Forms.ListView, ByRef ds As DataSet)

        On Error GoTo Err


        Dim lcols As Integer = lvw.Columns.Count
        Dim lrows As Integer = ds.Tables(0).Rows.Count
        Dim lViewItem As System.Windows.Forms.ListViewItem
        Dim r As Integer
        Dim colTag As String
        Dim dsField As String
        Dim dataType As String
        Dim CheckColBoolean As Boolean
        lvw.Items.Clear()
        If lrows > 0 Then

            For Each rows In ds.Tables(0).Rows
                lViewItem = New Windows.Forms.ListViewItem(ds.Tables(0).Rows(r).Item(0).ToString())
                For c As Integer = 1 To lcols - 1
                    dataType = rows.Table.Columns(c).DataType().ToString
                    ' colTag = lvw.Columns(c).Tag.ToString
                    If Not System.Convert.IsDBNull(ds.Tables(0).Rows(r).Item(c)) Then

                        Select Case dataType.ToString
                            Case "System.Bit"
                                If ds.Tables(0).Rows(r).Item(c) = 0 Then
                                    lViewItem.SubItems.Add("No")
                                Else
                                    lViewItem.SubItems.Add("Yes")
                                End If
                            Case "System.Boolean"
                                If ds.Tables(0).Rows(r).Item(c) = True Then
                                    lViewItem.SubItems.Add("Yes")
                                Else
                                    lViewItem.SubItems.Add("No")
                                End If
                            Case "System.Decimal"
                                lViewItem.SubItems.Add(CDbl(ds.Tables(0).Rows(r).Item(c)))
                            Case Else
                                lViewItem.SubItems.Add(ds.Tables(0).Rows(r).Item(c))
                        End Select

                    Else
                        lViewItem.SubItems.Add("")
                    End If
                Next
                lvw.Items.Add(lViewItem)
                r = r + 1
            Next
        End If

        For rc As Integer = 0 To lvw.Items.Count - 1
            lvw.Items(rc).Checked = (CLng(lvw.Items(rc).SubItems(1).Text) = 1 Or lvw.Items(rc).SubItems(1).Text = "Yes")
        Next

        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
    Public Sub LoadCalendar(ByRef lvw As System.Windows.Forms.ListView, ByRef ds As DataSet)

        On Error GoTo Err


        Dim lcols As Integer = lvw.Columns.Count
        Dim lrows As Integer = ds.Tables(0).Rows.Count
        Dim lViewItem As System.Windows.Forms.ListViewItem
        Dim r As Integer
        Dim colTag As String
        Dim dsField As String
        Dim dataType As String
        Dim SameStartTime As String = String.Empty
        Const firstCol = 1

        lvw.Items.Clear()
        If lrows > 0 Then

            For Each rows In ds.Tables(0).Rows
                lViewItem = New Windows.Forms.ListViewItem(ds.Tables(0).Rows(r).Item(0).ToString())
                For c As Integer = 1 To lcols - 1
                    dataType = rows.Table.Columns(c).DataType().ToString
                    ' colTag = lvw.Columns(c).Tag.ToString
                    If Not System.Convert.IsDBNull(ds.Tables(0).Rows(r).Item(c)) Then

                        Select Case dataType.ToString
                            Case "System.Bit"
                                If ds.Tables(0).Rows(r).Item(c) = 0 Then
                                    lViewItem.SubItems.Add("No")
                                Else
                                    lViewItem.SubItems.Add("Yes")
                                End If
                            Case "System.Boolean"
                                If ds.Tables(0).Rows(r).Item(c) = True Then
                                    lViewItem.SubItems.Add("Yes")
                                Else
                                    lViewItem.SubItems.Add("No")
                                End If
                            Case "System.Decimal"
                                lViewItem.SubItems.Add(CDbl(ds.Tables(0).Rows(r).Item(c)))
                            Case Else
                                If firstCol = c Then
                                    If SameStartTime = ds.Tables(0).Rows(r).Item(c) Then
                                        lViewItem.SubItems.Add("")
                                    Else
                                        lViewItem.SubItems.Add(ds.Tables(0).Rows(r).Item(c))
                                        SameStartTime = ds.Tables(0).Rows(r).Item(c)
                                    End If
                                Else
                                    lViewItem.SubItems.Add(ds.Tables(0).Rows(r).Item(c))
                                End If

                        End Select

                    Else
                        lViewItem.SubItems.Add("")
                    End If

                Next
                r = r + 1
                lvw.Items.Add(lViewItem)
            Next


        End If

        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
    Public Function ID(lvw As System.Windows.Forms.ListView)
        On Error GoTo Err

        Dim _id As Integer
        If (lvw.SelectedItems.Count > 0) Then
            _id = lvw.SelectedItems(0).Text
        Else
            _id = 0
        End If
        Return _id
        Exit Function
Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function GetColumnTag(lvw As System.Windows.Forms.ListView, col As Integer) As String

        Dim tag As String = String.Empty

        Try
            If lvw.Columns(col).Tag = String.Empty Then
                tag = ""
            End If
        Catch ex As Exception
            If (lvw.Items.Count > 0) Then
                Select Case lvw.Columns(col).Tag.ToString.ToLower
                    Case "string", "int", "datetime", "decimal"
                        tag = lvw.Columns(col).Tag
                    Case Else
                        tag = ""
                End Select
            End If
        End Try

        GetColumnTag = tag.ToLower.ToString


    End Function
    Public Function getColValue(lvw As System.Windows.Forms.ListView, col As Integer) As String
        On Error GoTo Err

        Dim getValue As String

        getValue = lvw.Items(lvw.FocusedItem.Index).SubItems.Item(col).Text

        Return getValue

        Exit Function
Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

    Public Function SelectCheckBox(ByRef lvw As System.Windows.Forms.ListView, ByRef switch As Boolean)

        On Error GoTo Err


        Dim lrows As Integer = lvw.Items.Count
        Dim lViewItem As System.Windows.Forms.ListViewItem
        Dim r As Integer


        If lrows > 0 Then
            For i As Integer = 0 To lrows - 1
                lvw.Items(i).Checked = switch
            Next
        End If

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function

    Public Function TotalSummary(ByRef lvw As System.Windows.Forms.ListView, ByRef col As Integer) As Double

        On Error GoTo Err


        Dim lrows As Integer = lvw.Items.Count
        Dim lViewItem As System.Windows.Forms.ListViewItem
        Dim r As Integer
        Dim _total As Double = 0.0

        If lrows > 0 Then
            For i As Integer = 0 To lrows - 1
                If IsNumeric(lvw.Items(i).SubItems(col).Text) Then
                    _total = _total + CDbl(lvw.Items(i).SubItems(col).Text)
                End If
            Next
        End If

        Return _total

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function CheckForValue(lvw As System.Windows.Forms.ListView, ByRef value As String, col As Integer) As Boolean
        On Error GoTo Err


        Dim lrows As Integer = lvw.Items.Count
        Dim lViewItem As System.Windows.Forms.ListViewItem
        Dim r As Integer
        Dim returnValue As Boolean = False

        If lrows > 0 Then
            For i As Integer = 0 To lrows - 1
                If lvw.Items(i).SubItems(col).Text = value Then
                    returnValue = True
                    Exit For
                End If
            Next
        End If

        Return returnValue

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Sub CheckBoxRemove(lvw As System.Windows.Forms.ListView)
        On Error GoTo Err


        Dim lrows As Integer = lvw.Items.Count
        Dim lViewItem As System.Windows.Forms.ListViewItem
        Dim r As Integer
        Dim returnValue As Boolean = False

        If lrows > 0 Then
            For i As Integer = 0 To lrows - 1
                If i = lrows Then Exit For
                If lvw.Items(i).Checked = True Then
                    lvw.Items(i).Remove()
                    i = i - 1 : lrows = lrows - 1
                End If

            Next
        End If

        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
    Public Sub RowColour(lvw As System.Windows.Forms.ListView, Col As Integer, ColValue As String, colour As System.Drawing.Color)
        On Error GoTo Err


        Dim lrows As Integer

        If lvw.Items.Count > 0 Then
            lrows = lvw.Items.Count
            Dim lViewItem As System.Windows.Forms.ListViewItem
            Dim r As Integer
            Dim returnValue As Boolean = False

            If lrows > 0 Then
                For i As Integer = 0 To lrows - 1
                    If i = lrows Then Exit For
                    lvw.Items(i).BackColor = Drawing.Color.White
                    If UCase(lvw.Items(i).SubItems(Col).Text) = UCase(ColValue) Then
                        lvw.Items(i).BackColor = colour
                    End If
                Next
            End If
        End If
        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
    Public Function RowColourReturn(lvw As System.Windows.Forms.ListView, Col As Integer, ColValue As String, colour As System.Drawing.Color) As Integer
        On Error GoTo Err

        Dim SRow As Integer = 0
        Dim lrows As Integer = lvw.Items.Count
        Dim lViewItem As System.Windows.Forms.ListViewItem
        Dim r As Integer
        Dim returnValue As Boolean = False

        If lrows > 0 Then
            For i As Integer = 0 To lrows - 1
                If i = lrows Then Exit For
                If lvw.Items(i).SubItems(Col).Text = ColValue Then
                    lvw.Items(i).BackColor = colour
                    SRow = SRow + 1
                End If
            Next
        End If
        Return SRow
        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function CheckBoxSelected(lvw As System.Windows.Forms.ListView) As Boolean
        On Error GoTo Err


        Dim lrows As Integer = lvw.Items.Count
        Dim lViewItem As System.Windows.Forms.ListViewItem
        Dim r As Integer
        Dim returnValue As Boolean = False

        If lrows > 0 Then
            For i As Integer = 0 To lrows - 1
                If lvw.Items(i).Checked = True Then
                    returnValue = True
                    Exit For
                End If
            Next
        End If

        Return returnValue

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function CheckBoxSelectedAmount(lvw As System.Windows.Forms.ListView) As Integer
        On Error GoTo Err


        Dim lrows As Integer = lvw.Items.Count
        Dim lViewItem As System.Windows.Forms.ListViewItem
        Dim r As Integer
        Dim returnValue As Integer = 0

        If lrows > 0 Then
            For i As Integer = 0 To lrows - 1
                If lvw.Items(i).Checked = True Then
                    returnValue = returnValue + 1
                End If
            Next
        End If

        Return returnValue

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function sum(ByRef lvw As System.Windows.Forms.ListView, ByRef col As Integer, Checked As Boolean) As Double

        On Error GoTo Err


        Dim lrows As Integer = lvw.Items.Count
        Dim lViewItem As System.Windows.Forms.ListViewItem
        Dim r As Integer
        Dim _total As Double = 0.0
        If Checked Then
            If lrows > 0 Then
                For i As Integer = 0 To lrows - 1
                    If lvw.Items(i).Checked Then
                        _total = _total + CDbl(lvw.Items(i).SubItems(col).Text)
                    End If
                Next
            End If

        Else


            If lrows > 0 Then
                For i As Integer = 0 To lrows - 1
                    _total = _total + CDbl(lvw.Items(i).SubItems(col).Text)
                Next
            End If
        End If
        Return _total

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Sub Find(ByRef lvw As System.Windows.Forms.ListView, criteria As String)

        On Error GoTo Err


        Dim rows As Integer = lvw.Items.Count
        Dim cols As Integer = lvw.Columns.Count

        If rows > 0 Then
            For r As Integer = 0 To rows - 1
                For c As Integer = 0 To cols - 1
                    If criteria = lvw.Items(r).SubItems(c).Text.ToString() Then
                        lvw.Focus()
                        lvw.Items(r).Selected = True

                        Exit For
                    End If
                Next
            Next
        End If


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
    Public Sub Sort(ByRef lvw As System.Windows.Forms.ListView, e As System.Windows.Forms.ColumnClickEventArgs)


        On Error GoTo Err
        Dim dataType As String = GetColumnTag(lvw, e.Column)

        ' if the list view is the same list view as before, trigger the descend
        If strlvwName = lvw.Name Then
            If e.Column = lvwColumSorter.SortColumn Then
                If lvwColumSorter.Order = System.Windows.Forms.SortOrder.Ascending Then
                    lvwColumSorter.Order = System.Windows.Forms.SortOrder.Descending
                Else
                    lvwColumSorter.Order = System.Windows.Forms.SortOrder.Ascending
                End If
            Else
                lvwColumSorter = New ListViewColumnSorter(dataType, e.Column)
                lvw.ListViewItemSorter = lvwColumSorter
                lvwColumSorter.SortColumn = e.Column
                lvwColumSorter.Order = System.Windows.Forms.SortOrder.Ascending
            End If
        Else
            lvwColumSorter = New ListViewColumnSorter(dataType, e.Column)
            lvw.ListViewItemSorter = lvwColumSorter
            strlvwName = lvw.Name.ToString

            lvwColumSorter.SortColumn = e.Column
            lvwColumSorter.Order = System.Windows.Forms.SortOrder.Ascending
        End If
        lvw.Sort()


Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        If Err.Description = String.Empty And Err.Number = 0 Then
        Else
            RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
        End If

    End Sub

    Public Function GetCountValueInColumn(lvw As System.Windows.Forms.ListView, Col As Integer, Value As String) As String


        On Error GoTo Err


        Dim lrows As Integer
        Dim lViewItem As System.Windows.Forms.ListViewItem
        Dim ReturnCount As String = 0
        If lvw.Items.Count > 0 Then
            lrows = lvw.Items.Count


            If lrows > 0 Then
                For r As Integer = 0 To lrows - 1
                    If lvw.Items(r).SubItems.Item(Col).Text = Value Then
                        ReturnCount = ReturnCount + 1
                    End If
                Next
            End If
        End If

        Return ReturnCount

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

End Class
