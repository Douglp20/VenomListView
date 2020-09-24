Public Class ListViewColumnSorter
    Implements IComparer
    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)

    Private ColumnToSort As Integer 'Specify the column to sort
    Private OrderOfSort As System.Windows.Forms.SortOrder 'Specifies the order in which to sort (i.e. 'Ascending').
    Private dataType As String ' Specifies the data type in which to sort (i.e. 'Ascending').
    Public Property SortColumn() As Integer
        Get
            Return ColumnToSort
        End Get
        Set(value As Integer)
            ColumnToSort = value
        End Set
    End Property

    Public Property Order() As System.Windows.Forms.SortOrder
        Get
            Return OrderOfSort
        End Get
        Set(value As System.Windows.Forms.SortOrder)
            OrderOfSort = value
        End Set
    End Property

    Public Sub New()
        ColumnToSort = 0 'Initialize the column to '0'
        OrderOfSort = System.Windows.Forms.SortOrder.None 'Initialize the sort order to 'none'
    End Sub
    Public Sub New(dataType As String)
        ColumnToSort = 0 'Initialize the column to '0'
        OrderOfSort = System.Windows.Forms.SortOrder.None 'Initialize the sort order to 'none'
        Me.dataType = dataType 'Initalize the datatype
    End Sub
    Public Sub New(dataType As String, col As Integer)
        ColumnToSort = col 'Initialize the column to '0'
        OrderOfSort = System.Windows.Forms.SortOrder.None 'Initialize the sort order to 'none'
        Me.dataType = dataType 'Initalize the datatype
    End Sub
    Public Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare
        On Error GoTo Err
        Dim compareResult As Integer
        Dim lv1 As String = (CType(x, System.Windows.Forms.ListViewItem).SubItems(ColumnToSort).Text)
        Dim lv2 As String = (CType(y, System.Windows.Forms.ListViewItem).SubItems(ColumnToSort).Text)

        'compareResult = compare the two items by dataType
        Select Case Me.dataType.ToLower.ToString()
            Case "datetime"
                If lv1.Length = 0 Then lv1 = "#31/12/1899#"
                If lv2.Length = 0 Then lv2 = "#31/12/1899#"
                compareResult = CDate(lv1).CompareTo(CDate(lv2))
            Case "int"
                If lv1.Length = 0 Then lv1 = 0
                If lv2.Length = 0 Then lv2 = 0
                compareResult = CLng(CLng(lv1.ToString()).CompareTo(CLng(lv2)))
            Case Else
                compareResult = lv1.CompareTo(lv2)
        End Select

        'Calculate correct return value based on object comparison
        If OrderOfSort = System.Windows.Forms.SortOrder.Ascending Then
            Return compareResult 'Ascending sort is selected, return normal result of compare operation

        ElseIf OrderOfSort = System.Windows.Forms.SortOrder.Descending Then
            Return (-compareResult) 'Descending sort is selected, return negative result of compare operation
        Else
            Return 0
        End If


        Exit Function
Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
End Class

