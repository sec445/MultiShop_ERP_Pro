# MultiShop ERP Pro – Test Plan

Run tests in a disposable copy of the backend first.

| ID | Test | Expected result |
|---|---|---|
| T01 | Create two shops | Both shops saved with unique IDs |
| T02 | Add same product to both shops | One product master; separate shop stock rows |
| T03 | Receive purchase of 5 units | Shop stock increases by 5 |
| T04 | Sell 1 unit | Stock decreases by 1; sale and lines saved |
| T05 | Attempt sale above available stock | Posting blocked; no partial transaction |
| T06 | Scan unknown barcode | Clear not-found message; no invoice posted |
| T07 | Register duplicate IMEI | Duplicate rejected |
| T08 | Transfer 2 units from Shop A to B | Source decreases and destination increases after receive |
| T09 | Cancel unposted invoice | No stock or ledger impact |
| T10 | Reprint posted invoice | Same invoice number and totals |
| T11 | Cashier opens user admin | Access denied |
| T12 | Two PCs post sales simultaneously | Both transactions saved; no duplicate invoice number |
| T13 | Backend unavailable | Friendly connection error; no false success |
| T14 | Backup and restore | Restored database opens and reconciles |
| T15 | Reconcile stock | Opening + receipts + transfers in - sales - transfers out = on-hand |

## Go-live checklist
- [ ] Two or more workstation test completed
- [ ] Permissions reviewed
- [ ] Backup scheduled and restore tested
- [ ] Invoice layout approved
- [ ] Opening stock reconciled
- [ ] User roles tested
- [ ] Accountant reviews tax/accounting configuration
