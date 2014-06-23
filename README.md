excel-etl
=========

**excel-etl** is an ETL library for executing SQL in Excel. Features include:

* Only requires Microsoft Office 2007 or above; no other database / business intelligence tools are needed
* Whole interface is Microsoft Excel; no learning curse for business users
* Processing based on database engine, much faster than tailor made native Excel 
* Support all data processing functionalities provided by Microsoft Access Database Engine, including joining multiple tables, aggregation, multiple ordering and arbitrary formula
* Support formulas output that reference other cells of the table
* Support formatting of the output cell
* Logics are separated from Excel File; all table definition and transformation are written in highly readeable XML format
* Log and stacktrace implemented for easier debugging

Why excel-etl?
---------

Traditionally there are 3 layers of data processing in reporting teams.

1. Enterprise ETL solutions, no matter source-type-insensitive like Informatica PowerCenter or source-type-specific like Oracle, are designed for huge amount of data and sophisticate ETL logics. Typical use cases include the daily / monthly updates of data warehouse and data marts.
2. Specialized Business Intelligence solutions like Business Objects, QlikView are more business-user-friendly solutions for reporting generating and data validation.
3. Spreadsheet applications like Microsoft Excel are great tools for quickly drafting reports.

But in practice, the solutions above do not always cover all needs for reporting. This is because:

1. Enterprise ETL and specialized BI tools are expensive, and they are not always justified the cost when the ETL process does not contain large amoung of data and complicated logics.
2. In the environment where auditing requirement is not high but agility is more of its concern, business users want to have a user friendly interface to dynamically update the data in the intermediate steps. Both ETL or BI tools do not provide such functionality.
3. Although it is possible to do manual ETL, it is very error-prone and repetative.


Although as it is universally agreed that vba is a crappy programming language, it doesn't prevent Microsoft Excel to be the de facto most popular Business Intelligence tool and exchange format in enterprise market. User would want to use Microsoft Excel as a viewer and editor.

**excel-etl** is thus developed to counter the challenges above.

How to use?
--------
The *main.xlsm* contains the whole library as well as an example demonstrating how to join 2 tables, aggregation and more. More examples to be added when requested.
