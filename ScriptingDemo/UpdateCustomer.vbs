Sub UpdateCustomer(num)
   Output.PrintLine("Changing " & Customers(num).ContactName & " to Bob Hope")
   Customers(num).ContactName = "Bob Hope"
End Sub