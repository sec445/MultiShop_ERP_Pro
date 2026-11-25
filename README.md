# MultiShop ERP Pro – Microsoft Access Starter Build
**Developer credit:** Coded by Mr Sabaz Ali Khan  
**Currency:** PKR (Rs.)  
**Target:** Microsoft Access desktop, split multi-user deployment

## Important delivery note
This package is a source-and-setup starter kit, not a compiled `.accdb` application. Microsoft Access is required to create the actual database file, import the VBA modules, and build the forms/reports. The package includes a DAO table installer, core VBA examples, schema reference, setup guide, and test plan.

## Multi-user architecture
Use a split database:
- `MultiShop_Backend.accdb`: tables only, stored in a reliable wired-LAN shared folder.
- `MultiShop_Frontend.accdb`: forms, reports, queries, macros and VBA; install a separate local copy on each workstation.
- Each frontend links to the backend tables using UNC path, e.g. `\\SERVER\MultiShopData\MultiShop_Backend.accdb`.
- Do not put the backend on OneDrive, Dropbox, a USB drive, or a Wi-Fi-only share.
- Ensure users have Modify permissions on the backend folder because Access creates a locking file (`.laccdb`).
- Back up the backend when all users have exited, or use a tested backup process.
- For high concurrency, remote access, or many branches, use SQL Server as backend and Access as frontend.

## Quick start
1. Install Microsoft Access (32-bit or 64-bit; use matching Office VBA references).
2. Create a blank database named `MultiShop_Frontend.accdb`.
3. Press `Alt+F11` → Insert → Module. Import `VBA/modSetup.bas`, `VBA/modUtilities.bas`, `VBA/modInventory.bas`, and `VBA/modSales.bas` using File → Import File.
4. Create a second blank database named `MultiShop_Backend.accdb`.
5. In the backend VBA editor, import `modSetup.bas`, open `BuildERPDatabase` and run it. This creates the core tables and indexes in the current database.
6. In the frontend, create a module with `LinkBackendTables` from `modUtilities.bas`; set `BACKEND_PATH` to the backend UNC path and run it.
7. Create the forms listed in `Documentation/FORM_REPORT_BUILD_GUIDE.md`. Set each form's Record Source / Control Source as specified.
8. Add a startup form named `frmDashboard` and set it as the Display Form in Access Options.
9. Create one admin login manually in `tblUsers` (see security note below), then create shop and product records.
10. Test using `Documentation/TEST_PLAN.md` before live use.

## Core design
- All transactions are shop-specific via `ShopID`.
- Inventory is stored per shop in `tblStock`.
- Mobile devices can be tracked by IMEI in `tblIMEI`.
- Sales header and lines are separated.
- Stock movements are recorded in `tblStockMovements`.
- Transfer headers and lines support branch-to-branch stock movement.

## Security limitations
Access file-level security is not equivalent to a server database. This starter uses role fields and application-level checks as a convenience, not as strong protection against a knowledgeable user with direct file access. Restrict network-folder permissions. For stronger authentication, auditing and concurrent multi-branch use, migrate tables to SQL Server and use Windows/Entra or SQL authentication.

## Accounting note
The included schema is an operational retail foundation. Before using it for statutory accounting or tax filing, have a qualified accountant review tax rules, chart of accounts, returns, and financial statements applicable to your jurisdiction.
