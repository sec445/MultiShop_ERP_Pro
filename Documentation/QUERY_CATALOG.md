# Saved Query Catalog

Create these in Access SQL View after tables are installed.

## qryStockByShop
```sql
SELECT S.ShopID, S.ProductID, P.SKU, P.ProductName,
       S.QtyOnHand, P.ReorderLevel, P.CostPrice,
       S.QtyOnHand * P.CostPrice AS StockValue
FROM tblStock AS S INNER JOIN tblProducts AS P
ON S.ProductID = P.ProductID;
```

## qryLowStock
```sql
SELECT S.ShopID, P.SKU, P.ProductName, S.QtyOnHand, P.ReorderLevel
FROM tblStock AS S INNER JOIN tblProducts AS P
ON S.ProductID = P.ProductID
WHERE S.QtyOnHand <= P.ReorderLevel;
```

## qrySalesByDay
```sql
SELECT DateValue(SaleDate) AS SaleDay, ShopID,
       Sum(NetTotal) AS NetSales, Sum(PaidAmount) AS Collected
FROM tblSales
WHERE SaleStatus='POSTED'
GROUP BY DateValue(SaleDate), ShopID;
```

## qryIMEIAvailable
```sql
SELECT I.IMEI, P.ProductName, I.ShopID, I.IMEIStatus
FROM tblIMEI AS I INNER JOIN tblProducts AS P
ON I.ProductID=P.ProductID
WHERE I.IMEIStatus='IN_STOCK';
```

Note: Access SQL syntax and reserved words can vary by version. If a query reports a syntax error, bracket field names and verify table/field names in the installed database.
