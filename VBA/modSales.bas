Attribute VB_Name = "modSales"
Option Compare Database
Option Explicit

' Example posting routine for one sales line.
' Build the full POS form around this pattern; post header, lines, stock and payments
' inside ONE DAO transaction in the production form.
Public Sub ValidateSaleLine(ByVal shopID As Long, ByVal productID As Long, ByVal qty As Long)
    If shopID <= 0 Or productID <= 0 Or qty <= 0 Then
        Err.Raise vbObjectError + 301, , "Select a shop/product and enter a quantity greater than zero."
    End If
    If GetShopStock(shopID, productID) < qty Then
        Err.Raise vbObjectError + 302, , "Insufficient stock for this sale."
    End If
End Sub

Public Function NextInvoiceNumber(ByVal shopID As Long) As String
    ' Unique index on InvoiceNo remains the final collision safeguard.
    NextInvoiceNumber = "S" & Format(Date, "yymmdd") & "-" & Format(shopID, "000") & "-" & _
        Format(CLng(Timer * 10) Mod 100000, "00000")
End Function

Public Function CalculateLineTotal(ByVal qty As Long, ByVal unitPrice As Currency, _
    Optional ByVal discount As Currency = 0) As Currency
    If qty < 1 Or unitPrice < 0 Or discount < 0 Then
        Err.Raise vbObjectError + 303, , "Invalid sale line values."
    End If
    CalculateLineTotal = (qty * unitPrice) - discount
    If CalculateLineTotal < 0 Then Err.Raise vbObjectError + 304, , "Discount cannot exceed line amount."
End Function
