Attribute VB_Name = "modInventory"
Option Compare Database
Option Explicit

Public Function GetShopStock(ByVal shopID As Long, ByVal productID As Long) As Long
    GetShopStock = Nz(DLookup("QtyOnHand", "tblStock", "ShopID=" & shopID & " AND ProductID=" & productID), 0)
End Function

Public Sub EnsureStockRow(ByVal shopID As Long, ByVal productID As Long)
    If DCount("*", "tblStock", "ShopID=" & shopID & " AND ProductID=" & productID) = 0 Then
        CurrentDb.Execute "INSERT INTO tblStock (ShopID, ProductID, QtyOnHand, LastUpdated) VALUES (" & _
            shopID & "," & productID & ",0,Now())", dbFailOnError
    End If
End Sub

Public Sub PostStockMovement(ByVal shopID As Long, ByVal productID As Long, _
    ByVal qtyIn As Long, ByVal qtyOut As Long, ByVal movementType As String, _
    ByVal refType As String, ByVal refID As Long)
    On Error GoTo ErrHandler
    Dim ws As DAO.Workspace, db As DAO.Database, currentQty As Long
    If qtyIn < 0 Or qtyOut < 0 Or (qtyIn > 0 And qtyOut > 0) Then
        Err.Raise vbObjectError + 201, , "Invalid movement quantities."
    End If
    Set ws = DBEngine.Workspaces(0)
    Set db = CurrentDb
    ws.BeginTrans
    EnsureStockRow shopID, productID
    currentQty = GetShopStock(shopID, productID)
    If qtyOut > currentQty Then Err.Raise vbObjectError + 202, , "Insufficient stock."
    db.Execute "UPDATE tblStock SET QtyOnHand=QtyOnHand+" & qtyIn & "-" & qtyOut & _
       ", LastUpdated=Now() WHERE ShopID=" & shopID & " AND ProductID=" & productID, dbFailOnError
    db.Execute "INSERT INTO tblStockMovements (ShopID,ProductID,MovementDate,MovementType,QtyIn,QtyOut,ReferenceType,ReferenceID) " & _
       "VALUES (" & shopID & "," & productID & ",Now(),'" & Replace(movementType, "'", "''") & "'," & qtyIn & "," & qtyOut & ",'" & _
       Replace(refType, "'", "''") & "'," & refID & ")", dbFailOnError
    ws.CommitTrans
    Exit Sub
ErrHandler:
    On Error Resume Next
    ws.Rollback
    MsgBox "Stock movement failed: " & Err.Description, vbExclamation
End Sub
