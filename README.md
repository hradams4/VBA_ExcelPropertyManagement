# VBA_ExcelPropertyManagement
###A VBA project to streamline and secure property management tasks and data.

*The project objective is to automate Excel Property Management's property ledger using a secure web-based solution that minimizes user error and ensures proper authorization privileges. It will provide access to ledger data on any device, implement best IT practices, and include manual auditing features, with exclusive administrator access to logging data, customer data and other restricted data stored in a dedicated Access database.

To enhance security and usability, the system must present as an Excel deployed VBA menu-driven application. Users, apart from the owner/admin (Sparky), should not have the ability to access the raw spreadsheet data. Upon opening, the system must authenticate users and grant role-specific access.*


###Project Requirements

<strong>REQ-EPM001 (Phase 0)</strong>
The ledger system shall support one administrative account, one business analyst account, and  support staff accounts equal to the amount of project managers employed.

<strong>REQ-EPM002 (Phase 0)</strong>
The ledger system shall allow only the admin account to manually access and manipulate the ledger spreadsheet in its raw form.

<strong>REQ-EPM003 (Phase 0)</strong>
The ledger system shall be deployed in Excel utilizing the VBA language for menus.

<strong>REQ-EPM004 (Phase 0)</strong>
The ledger system shall appear as a menu-driven application to staff and business analyst accounts.

<strong>REQ-EPM005 (Phase 0)</strong>
The ledger system shall authenticate the user by their unique login credentials and allow access to only the permissions granted for their respective account type.

<strong>REQ-EPM006 (Phase 0)</strong>
The ledger system shall support a legacy mode that shall direct only the authenticated admin user to the ledger spreadsheet for manual access and manipulation.

<strong>REQ-EPM007 (Phase 0)</strong>
The ledger system shall prompt the authenticated admin user to either choose legacy mode or the standard menu mode before proceeding through the system.

<strong>REQ-EPM008 (Phase 0)</strong>
When the authenticated admin user selects menu mode, the ledger system shall present a warning for all properties that have a lease expiring within 90 days.

<strong>REQ-EPM009 (Phase 0)</strong>
The ledger system shall prompt for any ledger items to be entered based on the date of entry.

<strong>REQ-EPM010 (Phase 0)</strong>
The ledger system shall prompt the user to enter a ledger entry for a late fee, based on property metrics, to be applied if rent is 3-5 days past the rent anniversary.

<strong>REQ-EPM011 (Phase 0)</strong>
The ledger system shall allow both admin and support staff account holders to enter tenant credit.

<strong>REQ-EPM012 (Phase 0)</strong>
The ledger system shall allow only the admin account holder to add and close a tenant from a property, add a new property, add and change user roles, and preset property groups for reports and access rights of the support staff. 

<strong>REQ-EPM013 (Phase 0)</strong>
The ledger system shall allow support staff accounts access to only a subset of rentals grouped and assigned by the admin account. 

<strong>REQ-EPM014 (Phase 0)</strong>
The ledger system shall allow support staff accounts to add new ledger entries with a comment that can include either rent received, a charge for damages, or credit to the tenant for rent received or tenant purchased materials.

<strong>REQ-EPM015 (Phase 0)</strong>
The ledger system shall allow support staff accounts to view the full ledger data while not allowing the modification of past entries.

<strong>REQ-EPM016 (Phase 0)</strong>
The ledger system shall allow support staff accounts to access tenants phone numbers, property address, email address(es), due dates and late fee grace period.

<strong>REQ-EPM017 (Phase 0)</strong>
The ledger system shall allow the business analyst account only the ability to access and generate reports of the various properties gross rent collected by quarter or by year, individually, grouped, or as a whole through rollup numbers in the ledger.

<strong>REQ-EPM018 (Phase 0)</strong>
The ledger system shall not allow the business analyst account access to any customer data or ledger entries.

<strong>REQ-EPM019(Phase 0)</strong>
The ledger system shall store customer data locally in an access database.

<strong>REQ-EPM020 (Phase 1)</strong>
The ledger system shall be accessible by all web enabled devices by utilizing user authentication through a secure login. 

<strong>REQ-EPM021 (Phase 1)</strong>
The ledger system shall be hosted by a dedicated server with a dedicated database.

<strong>REQ-EPM022 (Phase 1)</strong>
The ledger system data shall be TLS encrypted to protect customer information.

<strong>REQ-EPM023 (Phase 1)</strong>
The ledger system shall create a backup of its data every 7 days.

<strong>REQ-EPM024 (Phase 1)</strong>
The ledger system shall track user activity which includes user logins, ledger additions and modifications that shall be viewable only by the admin account holder.

