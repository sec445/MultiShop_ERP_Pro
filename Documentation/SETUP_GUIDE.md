# Setup Guide

## A. Create backend
1. Open Access → Blank database → save as `MultiShop_Backend.accdb`.
2. Press Alt+F11.
3. Import `VBA/modSetup.bas`.
4. Run `BuildERPDatabase` from the VBA Immediate Window (`Ctrl+G`) or assign it to a temporary button.
5. Confirm tables appear in the Navigation Pane.

## B. Create frontend
1. Create `MultiShop_Frontend.accdb`.
2. Import all VBA modules from the `VBA` folder.
3. In `modUtilities.bas`, edit `BACKEND_PATH` to the actual UNC path.
4. Run `LinkBackendTables`.
5. Create forms using the form guide.
6. Create saved queries from `Database/QUERY_CATALOG.md`.
7. Create reports using the report guide.

## C. Enable VBA
Access → File → Options → Trust Center → Trust Center Settings → Trusted Locations. Add the local application folder. Do not broadly enable all macros.

## D. Network deployment
1. Put backend in a shared folder on a stable wired-LAN Windows PC/server.
2. Give each authorized user Modify permission on that folder.
3. Copy the frontend to each workstation's local drive, e.g. `C:\\MultiShopERP\\`.
4. Link all frontends to the same backend.
5. Test two simultaneous users before entering real data.
6. Keep daily verified backups and periodically test restore.

## E. First-time data
1. Add shops to `tblShops`.
2. Add user records to `tblUsers`; assign RoleName = Admin to the owner.
3. Add categories and products.
4. Add suppliers and customers.
5. Enter opening inventory using a controlled opening-stock process. Do not directly edit stock totals after transactions begin.
6. Set purchase cost and selling price for each product.

## F. Form creation essentials
Use bound forms for master data, and unbound or carefully controlled forms for transaction entry. Use transactions / DAO workspaces for multi-table posting so a failed invoice cannot leave half-posted stock.
