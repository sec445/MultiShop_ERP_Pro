Attribute VB_Name = "modSetup"
Option Compare Database
Option Explicit

' Run this in a NEW, EMPTY backend database.
Public Sub BuildERPDatabase()
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Set db = CurrentDb

    CreateTable db, "tblShops", _
        "ShopID COUNTER CONSTRAINT pk_tblShops PRIMARY KEY, ShopCode TEXT(20), ShopName TEXT(100), Phone TEXT(30), Address TEXT(255), IsActive YESNO"
    CreateTable db, "tblUsers", _
        "UserID COUNTER CONSTRAINT pk_tblUsers PRIMARY KEY, UserName TEXT(50), PasswordHash TEXT(255), FullName TEXT(100), RoleName TEXT(20), DefaultShopID LONG, IsActive YESNO"
    CreateTable db, "tblCategories", _
        "CategoryID COUNTER CONSTRAINT pk_tblCategories PRIMARY KEY, CategoryName TEXT(80), IsActive YESNO"
    CreateTable db, "tblProducts", _
        "ProductID COUNTER CONSTRAINT pk_tblProducts PRIMARY KEY, SKU TEXT(40), Barcode TEXT(60), ProductName TEXT(150), CategoryID LONG, Brand TEXT(80), Model TEXT(80), UnitName TEXT(20), CostPrice CURRENCY, SalePrice CURRENCY, ReorderLevel LONG, TrackIMEI YESNO, IsActive YESNO"
    CreateTable db, "tblCustomers", _
        "CustomerID COUNTER CONSTRAINT pk_tblCustomers PRIMARY KEY, CustomerName TEXT(120), Phone TEXT(30), Address TEXT(255), CreditLimit CURRENCY, IsActive YESNO"
    CreateTable db, "tblSuppliers", _
        "SupplierID COUNTER CONSTRAINT pk_tblSuppliers PRIMARY KEY, SupplierName TEXT(120), Phone TEXT(30), Address TEXT(255), IsActive YESNO"
    CreateTable db, "tblStock", _
        "StockID COUNTER CONSTRAINT pk_tblStock PRIMARY KEY, ShopID LONG, ProductID LONG, QtyOnHand LONG, LastUpdated DATETIME"
    CreateTable db, "tblPurchases", _
        "PurchaseID COUNTER CONSTRAINT pk_tblPurchases PRIMARY KEY, PurchaseNo TEXT(30), ShopID LONG, SupplierID LONG, PurchaseDate DATETIME, SubTotal CURRENCY, DiscountAmount CURRENCY, NetTotal CURRENCY, PaidAmount CURRENCY, PurchaseStatus TEXT(20), Notes TEXT(255)"
    CreateTable db, "tblPurchaseDetails", _
        "PurchaseDetailID COUNTER CONSTRAINT pk_tblPurchaseDetails PRIMARY KEY, PurchaseID LONG, ProductID LONG, Qty LONG, UnitCost CURRENCY, LineTotal CURRENCY"
    CreateTable db, "tblSales", _
        "SaleID COUNTER CONSTRAINT pk_tblSales PRIMARY KEY, InvoiceNo TEXT(30), ShopID LONG, CustomerID LONG, SaleDate DATETIME, SubTotal CURRENCY, DiscountAmount CURRENCY, TaxAmount CURRENCY, NetTotal CURRENCY, PaidAmount CURRENCY, PaymentMethod TEXT(20), SaleStatus TEXT(20), UserID LONG"
    CreateTable db, "tblSaleDetails", _
        "SaleDetailID COUNTER CONSTRAINT pk_tblSaleDetails PRIMARY KEY, SaleID LONG, ProductID LONG, Qty LONG, UnitPrice CURRENCY, UnitCost CURRENCY, LineDiscount CURRENCY, LineTotal CURRENCY, IMEI TEXT(30)"
    CreateTable db, "tblStockMovements", _
        "MovementID COUNTER CONSTRAINT pk_tblStockMovements PRIMARY KEY, ShopID LONG, ProductID LONG, MovementDate DATETIME, MovementType TEXT(30), QtyIn LONG, QtyOut LONG, ReferenceType TEXT(30), ReferenceID LONG, UserID LONG, Notes TEXT(255)"
    CreateTable db, "tblStockTransfers", _
        "TransferID COUNTER CONSTRAINT pk_tblStockTransfers PRIMARY KEY, TransferNo TEXT(30), FromShopID LONG, ToShopID LONG, TransferDate DATETIME, TransferStatus TEXT(20), CreatedBy LONG, ReceivedBy LONG, Notes TEXT(255)"
    CreateTable db, "tblStockTransferDetails", _
        "TransferDetailID COUNTER CONSTRAINT pk_tblStockTransferDetails PRIMARY KEY, TransferID LONG, ProductID LONG, Qty LONG"
    CreateTable db, "tblIMEI", _
        "IMEIID COUNTER CONSTRAINT pk_tblIMEI PRIMARY KEY, IMEI TEXT(30), ProductID LONG, ShopID LONG, PurchaseDetailID LONG, SaleDetailID LONG, IMEIStatus TEXT(20), WarrantyMonths LONG, WarrantyEnd DATETIME"
    CreateTable db, "tblPayments", _
        "PaymentID COUNTER CONSTRAINT pk_tblPayments PRIMARY KEY, ShopID LONG, PaymentDate DATETIME, PartyType TEXT(20), PartyID LONG, ReferenceType TEXT(20), ReferenceID LONG, Amount CURRENCY, PaymentMethod TEXT(20), Direction TEXT(10), Notes TEXT(255), UserID LONG"
    CreateTable db, "tblExpenses", _
        "ExpenseID COUNTER CONSTRAINT pk_tblExpenses PRIMARY KEY, ShopID LONG, ExpenseDate DATETIME, ExpenseCategory TEXT(60), Description TEXT(255), Amount CURRENCY, PaidBy TEXT(20), UserID LONG"
    CreateTable db, "tblAuditLog", _
        "AuditID COUNTER CONSTRAINT pk_tblAuditLog PRIMARY KEY, EventDate DATETIME, UserID LONG, ActionName TEXT(60), TableName TEXT(60), RecordID LONG, Details TEXT(255)"

    AddIndex db, "tblProducts", "ux_ProductSKU", "SKU", True
    AddIndex db, "tblProducts", "ux_ProductBarcode", "Barcode", True
    AddIndex db, "tblStock", "ux_StockShopProduct", "ShopID, ProductID", True
    AddIndex db, "tblIMEI", "ux_IMEI", "IMEI", True
    AddIndex db, "tblSales", "ux_InvoiceNo", "InvoiceNo", True
    AddIndex db, "tblUsers", "ux_UserName", "UserName", True

    MsgBox "Core ERP tables and indexes created successfully.", vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Setup stopped: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

Private Sub CreateTable(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldDDL As String)
    On Error Resume Next
    db.Execute "CREATE TABLE [" & tableName & "] (" & fieldDDL & ")", dbFailOnError
    If Err.Number <> 0 Then
        Dim msg As String
        msg = Err.Description
        Err.Clear
        If Not TableExists(db, tableName) Then Err.Raise vbObjectError + 100, , "Could not create " & tableName & ": " & msg
    End If
    On Error GoTo 0
End Sub

Private Function TableExists(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
    Dim t As DAO.TableDef
    For Each t In db.TableDefs
        If StrComp(t.Name, tableName, vbTextCompare) = 0 Then TableExists = True: Exit Function
    Next
End Function

Private Sub AddIndex(ByVal db As DAO.Database, ByVal tableName As String, ByVal indexName As String, ByVal fields As String, ByVal isUnique As Boolean)
    On Error Resume Next
    Dim sql As String
    sql = "CREATE " & IIf(isUnique, "UNIQUE ", "") & "INDEX [" & indexName & "] ON [" & tableName & "] (" & fields & ")"
    db.Execute sql, dbFailOnError
    Err.Clear
    On Error GoTo 0
End Sub
