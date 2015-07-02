# oPos to FacturaSol

The oPos to FacturaSol "OT-Fs" is a simple web script that allows to sync the clients and invoices from Unicenta oPost to the Factusol software, it generates excel / odt files to export all the cientes, invoices and invoice lines to Factura Sol (2015EV).

  - Exports Clients, Invoices and Invoice lines
  - Fast Script (Up to 400 Invoices / Second)
  - Easy to fork and modify

This small script uses the PHPExcel library, so you can read the requeriments of this library in project's page: https://phpexcel.codeplex.com

> Ensure that the script have access to write in the output directory, the generated files will be copied here

### Version
1.0.1


### Installation

```sh
To configure the script simply open the index.php and define the database constants (First lines of the script)
```
### Todo's

* Refactor the code
* Sync also providers and providers factures

License
----

LGPL V2.1
