Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop.Excel

Module Updated
    Private excel
    Private workbook

    Sub Main()

        Dim CustomerArrayObj = New List(Of Customer)
        Dim localArrayObj = OpenAndReadFile()
        Dim upperBound As Integer = localArrayObj.GetUpperBound(0)
        Dim lowerBound As Integer = localArrayObj.GetUpperBound(1)
        localArrayObj = ParseAndRemove(localArrayObj, lowerBound, upperBound)
        CustomerArrayObj = CreateAllCustomers(localArrayObj, CustomerArrayObj, lowerBound, upperBound)
        CustomerArrayObj = SortCustomersByValue(CustomerArrayObj)
        localArrayObj = ParseAndCheckMiddleName(localArrayObj, lowerBound, upperBound)
        PrintCustomerData(localArrayObj, lowerBound, upperBound)
        Console.WriteLine("Printing Class Data (Sorted Ascending By Account Amount): " + Environment.NewLine)
        CustomerArrayObj = SortCustomersByValue(CustomerArrayObj)
        PrintCustomerDataFromClass(CustomerArrayObj)

        Console.ReadLine()
    End Sub


    Public Function OpenAndReadFile() As Object(,)
        excel = New Application
        workbook = excel.workbooks.open("c:\users\chris\downloads\programmingtestreport")
        For count As Integer = 1 To workbook.sheets.count
            Dim sheet As Worksheet = workbook.sheets(count)
            Dim range As Range = sheet.UsedRange
            Dim customerData(,) As Object = range.Value(XlRangeValueDataType.xlRangeValueDefault)
            workbook.Close()
            Return customerData
        Next
        'Not sure how to return an object if the workbook doesn't exist here, throws a small warning, but no problems as the file exists'
    End Function

    Public Function ParseAndRemove(ByVal customerData As Object, lowerBound As Integer, upperBound As Integer) As Object(,)
        If customerData IsNot Nothing Then
            For row As Integer = 1 To upperBound
                For col As Integer = 1 To lowerBound
                    Dim s1 As String = customerData(row, col)
                    If (Not s1.Contains("PROCESS")) AndAlso (Not s1.Contains("REJECTS") AndAlso (Not s1.Contains("Insured")) AndAlso 'Remove all lines that have no customer information in them'
                        (Not s1.Contains("INTERNAL"))) Then
                        Dim updatedString = (Regex.Replace(s1, "\^.*", ""))
                        customerData(row, col) = updatedString
                    Else
                        customerData(row, col) = "INVALID"
                    End If
                Next
            Next
        End If
        Return customerData
    End Function

    Public Function ParseAndCheckMiddleName(ByVal customerData As Object, lowerBound As Integer, upperBound As Integer) As Object(,)
        For row As Integer = 1 To upperBound
            For col As Integer = 1 To lowerBound
                If (customerData(row, col) IsNot "INVALID") Then
                    Dim checkMiddleName = Regex.Match(customerData(row, col).Split(" ")(1), "[A-Z]\.")
                    If (checkMiddleName.Success) Then
                        customerData(row, col) = ContainsMiddleIntial(customerData(row, col))

                    Else
                        customerData(row, col) = NoMiddleInitial(customerData(row, col))
                    End If
                End If
            Next
        Next
        Return customerData
    End Function

    Public Function ContainsMiddleIntial(ByVal customerInfo As String) As String
        customerInfo = If(customerInfo.Split(" ").Length > 2, "Insured: " + customerInfo.Split(" ")(0) + " " + customerInfo.Split(" ")(1) + " " +
           customerInfo.Split(" ")(2) + Environment.NewLine + "Policy_Num: " + customerInfo.Split(" ")(3) + Environment.NewLine + "Amount: " +
         customerInfo.Split(" ")(4) + Environment.NewLine + "Effective Date: " + customerInfo.Split(" ")(5) + customerInfo.Split(" ")(6) + Environment.NewLine +
          "More Info: " + customerInfo.Split(" ")(7) + Environment.NewLine + "Reason: " + customerInfo.Split(" ")(8) + Environment.NewLine, Nothing)
        Return customerInfo
    End Function

    Public Function NoMiddleInitial(ByVal customerInfo As String) As String
        customerInfo = If(customerInfo.Split(" ").Length > 2, "Insured: " + customerInfo.Split(" ")(0) + " " + customerInfo.Split(" ")(1) + Environment.NewLine +
                                 "Policy_Num: " + customerInfo.Split(" ")(2) + Environment.NewLine + "Amount: " + customerInfo.Split(" ")(3) + Environment.NewLine + "Effective Date: " +
                                customerInfo.Split(" ")(4) + customerInfo.Split(" ")(5) + Environment.NewLine + "More Info: " + customerInfo.Split(" ")(6) + Environment.NewLine +
                                 "Reason: " + customerInfo.Split(" ")(7) + Environment.NewLine, Nothing)
        Return customerInfo
    End Function

    Sub PrintCustomerData(ByVal customerData As Object, lowerBound As Integer, upperBound As Integer)
        For row As Integer = 1 To upperBound
            For col As Integer = 1 To lowerBound
                If (customerData(row, col) IsNot "INVALID") Then
                    Console.WriteLine(customerData(row, col))
                End If
            Next
        Next
    End Sub

    Public Function ConvertToCustomer(ByVal customerData As Object, hasMiddleName As Boolean) As Customer
        Dim cust = New Customer
        If (customerData IsNot "INVALID") Then
            With cust
                If (hasMiddleName) Then
                    .GetSetName = customerData.Split(" ")(0) + " " + customerData.Split(" ")(1) + customerData.Split(" ")(2)
                    .GetSetAccountNumber = customerData.Split(" ")(3)
                    .GetSetAmount = Convert.ToDouble(customerData.Split(" ")(4))
                    .GetSetTransactionDate = Convert.ToDateTime(customerData.Split(" ")(5) + " " + customerData.Split(" ")(6))
                    .GetSetMoreInfo = customerData.Split(" ")(7)
                    .GetSetCustomerReason = customerData.Split(" ")(8)
                Else
                    .GetSetName = customerData.Split(" ")(0) + " " + customerData.Split(" ")(1)
                    .GetSetAccountNumber = customerData.Split(" ")(2)
                    .GetSetAmount = Convert.ToDouble(customerData.Split(" ")(3))
                    .GetSetTransactionDate = Convert.ToDateTime(customerData.Split(" ")(4) + " " + customerData.Split(" ")(5))
                    .GetSetMoreInfo = customerData.Split(" ")(6)
                    .GetSetCustomerReason = customerData.Split(" ")(7)
                End If
            End With
        End If
        Return cust
    End Function


    Public Function AddToCustomerList(ByVal customerList As List(Of Customer), customer As Customer) As List(Of Customer)
        customerList.Add(customer)
        Return customerList
    End Function


    Public Function CreateAllCustomers(ByVal customerData As Object, customerList As List(Of Customer), lowerBound As Integer, upperBound As Integer)
        For row As Integer = 1 To upperBound
            For col As Integer = 1 To lowerBound
                If (customerData(row, col) IsNot "INVALID") Then
                    Dim localCust = New Customer
                    Dim checkMiddleName = Regex.Match(customerData(row, col).Split(" ")(1), "[A-Z]\.")
                    If (checkMiddleName.Success) Then
                        localCust = ConvertToCustomer(customerData(row, col), hasMiddleName:=True)
                        customerList = AddToCustomerList(customerList, localCust)
                    Else
                        localCust = ConvertToCustomer(customerData(row, col), hasMiddleName:=False)
                        customerList = AddToCustomerList(customerList, localCust)
                    End If
                End If
            Next
        Next
        Return customerList
    End Function

    Public Function SortCustomersByValue(customerList As List(Of Customer)) As List(Of Customer)
        customerList.Sort(AddressOf CompareCustomerByAccountAmount)
        Return customerList
    End Function

    Public Function CompareCustomerByAccountAmount(x As Customer, y As Customer) As Double
        Return x.GetSetAmount.CompareTo(y.GetSetAmount)
    End Function

    Sub PrintCustomerDataFromClass(ByVal customerList As List(Of Customer))
        For i As Integer = 0 To customerList.Count - 1
            Console.Write("Insured: ")
            Console.WriteLine(customerList(i).GetSetName)
            Console.Write("Amount: ")
            Console.WriteLine(customerList(i).GetSetAmount)
            Console.WriteLine()
        Next
    End Sub

    Public Class Customer

        Private name As String
        Property GetSetName() As String
            Get
                Return name
            End Get
            Set(ByVal Value As String)
                name = Value
            End Set
        End Property

        Private AccountNumber As String
        Property GetSetAccountNumber() As String
            Get
                Return AccountNumber
            End Get

            Set(ByVal Value As String)
                AccountNumber = Value
            End Set
        End Property

        Private Amount As Double
        Property GetSetAmount() As Double
            Get
                Return Amount

            End Get
            Set(ByVal Value As Double)
                Amount = Value
            End Set
        End Property

        Private TransactionDate As DateTime
        Property GetSetTransactionDate() As DateTime
            Get
                Return TransactionDate
            End Get
            Set(ByVal Value As DateTime)
                TransactionDate = Value
            End Set
        End Property

        Private CustomerReason As String
        Property GetSetCustomerReason() As String
            Get
                Return CustomerReason
            End Get
            Set(ByVal Value As String)
                CustomerReason = Value
            End Set
        End Property

        Private MoreInfo As Char
        Property GetSetMoreInfo() As Char
            Get
                Return MoreInfo
            End Get
            Set(ByVal Value As Char)
                MoreInfo = Value
            End Set
        End Property

    End Class




End Module


