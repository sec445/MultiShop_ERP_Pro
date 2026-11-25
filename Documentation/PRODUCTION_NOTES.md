# Production completion items
This starter intentionally does not claim to be a finished retail accounting product. Before live use, implement and verify:
1. Full frmLogin password hashing (never store plain-text passwords), lockout and session handling.
2. Complete POS transaction routine posting invoice header, detail rows, IMEI assignment, stock movements and payment atomically.
3. Purchase posting, returns, sales returns, void/reversal workflow and transfer dispatch/receive workflows.
4. Unique invoice numbering using a dedicated counter table and transaction-safe increment (do not rely on Timer-based sample number).
5. Referential integrity, validation rules, required fields and duplicate prevention.
6. Tax configuration appropriate to the business and jurisdiction.
7. Form-level role checks plus Windows folder permissions.
8. Error logging, concurrency tests and backup restore drills.
