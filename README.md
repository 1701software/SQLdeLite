SQLdeLite 
by 1701 Software, Inc. 
http://www.1701software.com
=======================

SQLdeLite is an open source library that allows you to speed up your database development with Xojo.

Highlights:

- Single drop in module that speeds up your development.
- Automatically uses SQL prepared statements mitigating SQL injection attacks and speeding up database performance.
- Quickly and easily create SQL queries by using variables representing the properties of your objects. These variables are converted to the bounded parameters in prepared statements.
- No more string concatenation for your SQL!
- Dynamic objects that allow for any number of properties without having to define each one in the IDE. Now it's much faster to handle your query parameters and results.
- Built on top of the new Xojo framework. 
- Full support for iOSSQLiteDatabase on iOS! You can use the same business logic between projects despite using different database backends.
- Full support for all Xojo supported databases. Enable databases that require plugins by setting the appropriate constant to True (example: PLUGIN_MSSQL_ENABLED).
- Full support for cubeSQL. Make sure to enable support by changing the PLUGIN_CUBESQL_ENABLED constant to True.
- Valentina database is also supported by virtue of the SQLdeLite.ParameterizeSQL() method. This converts your SQL query into a Valentina compatible query with bound parameters.

The library does not expose any new database classes nor requires the usage of custom database adapters (as in "Active Record" from BKeeney Software).
The module extends the built-in Xojo databases classes and provides you with two additional methods. These methods replace the default SQLSelect() and 
SQLExecute() methods for your database of choice. 

Why use over ActiveRecord?

- ActiveRecord can only load a record via it's primary key which is forced to be an integer. It has the ability to load an object from a RecordSet which SQLdeLite can also do automatically (see the Advanced Features topic below).
- ActiveRecord is not available on iOS. SQLdeLite runs everywhere Xojo runs: Console, Desktop, Web, and iOS without any modifications. 
- ActiveRecord requires you to use their database specific adapters. SQLdeLite extends the Xojo native databases .
- ActiveRecord requires code generation using the commercial ARGen product or hand building your database classes. SQLdeLite can use classes or dynamic objects via SQLdeLite.Record.
- SQLdeLite is HALF the size contained inside a single module.
- SQLdeLite is built on top of the new Xojo framework and ready for the future.

Methods: 

- SQLdeLiteSelect(sqlStatementAsText, SQLdeLiteRecordObject)
- SQLdeLiteExecute(sqlStatementAsText, SQLdeLiteRecordObject)
- CreateInsertStatement(databaseObject, TableNameAsText, TableAndFieldNamesQuotedAsBoolean)

What is the "SQLdeLite.Record" class represented above as the SQLdeLiteRecordObject?

Good question! Have you ever built a large library of classes that supports your business logic and thus you have properties mapped to your database model? 
Isn't it frustruating managing the lifecycle of those objects? For instance you may initialize an instance of the object and fill some of its properties prior to doing 
a look up to a database. You might fill most of the properties after loading the data from the database and thus the object is less useful before being loaded. Which
properties should be available before/after you interact with the database?

Or how about during development when you just want to create a SQL statement using a number of variables. Whether you store those variables as individual variables
in your method or they are properties of an object you end up with some string concatenation gore looking like:

Dim sql As Text

sql = "SELECT * FROM Table WHERE Field = '" + variable1 + "' AND Field2 = " + variable2.ToText() + " AND Field3 = '" + variable3 + "';"

Some of you might do it the slightly faster way with an array and joining it at the end. Regardless this is dangerous for a number of reasons:

- The database engine does not benefit from query optimizations made possible with prepared statements and binding parameters.
- Easy to make mistakes as the developer as you try to concatenate the strings together properly. 
- Easy to include more fields than necessary in the statement or possibly fields you do not have valid values for.
- Your SQL statement is vulnerable to SQL injection because you are not properly escaping quotations characters.

Introducing the SQLdeLite.Record class. You can initialize an instance of it or sub-class it and use as needed. With SQLdeLite.Record you can create dynamic objects 
by filling the properties as you see fit without actually creating and building an object. Behind the scenes when you pass your instance of SQLdeLite.Record to 
the engine automatically converts all of your dynamic properties to SQL parameters. It then binds those parameters to a prepared statement appropriate for the 
database engine you are currently using. PostgreSQL, Oracle, and cubeSQL all handle parameter binding in different ways and SQLdeLite abstracts those differences away.

Building SQL Statements:

So in order to use SQLdeLite.Record and parameterize your SQL statement you can do the following.

----------

Dim row As New SQLdeLite.Record

row.Name = "Phillip Zedalis"

row.Title = "Managing Developer"

row.Company = "1701 Software, Inc."


Dim sql As Text

sql = "SELECT * FROM Users WHERE Name = $Name AND Title = $Title AND Company = $Company"

Dim rs As RecordSet

rs = db.SQLdeLiteSelect(sql, row)

----------

What happened behind the scenes is your new instantly created dynamic object was used to convert the SQL with $variables into a executable query for your 
database engine. In order to use a property of your SQLdeLite.Record object you simply pass in the case-sensitive name of the property preceeded by a $ symbol.

Advanced Features:

The SQLdeLiteSelect method also supports filling the results of the RecordSet back to your SQLdeLite.Record object. You pass True as the last parameter AND your 
query must return only one result. Assuming both factors are true your SQLdeLite.Record object will gain new dynamic properties representing the values of every 
column in the RecordSet. For example if we use the same "row" object as in the code example above and call the SQLdeLiteSelect method as so:

rs = db.SQLdeLiteSelect(sql, row, True)

The code above will actually loop through all the columns of your record and create dynamic properties in the row object. So despite never defining a "PhoneNumber"
property if the record included it then you can now access it via:

MsgBox(row.PhoneNumber)

No looping through your fields and binding the values or creating an object for every possible query you may want to run.

Valentina Support:

The Valentina database engine is a fantastic database that I use in many projects. Unfortunately the VDatabase object does not inherit from the Xojo Database object 
and thus the SQLdeLite extension methods are not available. However this turns out to be okay because Valentina has several different ways to query the engine/server 
that vary depending on your needs. 

Instead of using SQLdeLite to execute the queries you can simply use it to create your queries along with parameterized arrays suitable for Valentina. SQLdeLite is aware 
of the Valentina specific way of binding SQL parameters and returns to you everything you need to execute your queries against Valentina safely.


