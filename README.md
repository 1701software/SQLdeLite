SQLdeLite 1.0.0 beta by 1701 Software

What is SQLdeLite?

SQLdeLite is a low level SQL database abstraction layer for the XOJO language being actively developed here, at 1701 software.
It's primarily focused around the things that make managing and querying your databases difficult or time consuming.
It is designed around our experience with various database systems and our belief that heavy ORM's are not good. They tend to lock you in to a very particular way of using your databases and generally are not well optimized.
It’s designed to be a drop in replacement for the existing SQLiteDatabase and CubeSQLServer classes. All existing properties and methods should work out of the box. You will not have to change any of your existing code to start using SQLdeLite. MySQL and PostgreSQL are on the works.

What it can do?

Prepared Statements:

Currently if you wish to properly query a database you need to parameterize the query using prepared statements. Prepared statements provide both speed optimizations and protect you from SQL injection attacks. However, actually doing this for each database is a chore. SQLdeLite provides query helpers in the form of database methods (SQLdeLiteSelect and SQLdeLiteExecute) that rapidly speed up your development. You can now do something like:
myDb.SQLdeLiteSelect("SELECT * FROM Users WHERE Username = #?# AND Password = #?#", txtUsername.Text, txtPassword.Text)
This would return a RecordSet from your database. Behind the scenes it used prepared SQL statements appropriate for that database server to return the data.

Easier table schemas management:

It will assist you with creating and managing your table schemas through database migrations. By marking classes as an instance of a SQLdeLite.Table object you gain immediate schema capabilities. Public non-computed properties that are strings/integers/doubles will be compared with the table schema. If the table does not exist a "CreateTableSchema" event is raised in the database object.

If the table is missing particular fields or the data types do not match then a "UpdateTableSchema" event is raised. You can choose to update your schema or just return True. Again, the goal is flexibility and power, it does not force you to do anything. The table and class schemas are hashed and stored in the database object so subsequent queries do not involve introspection.

There is an example provided of using more complex types like dates. However the intent is to keep it low level. You should save Date.SQLDateTime as opposed to just Date. Perhaps use these classes as super's for more advanced classes with computed properties? Or use them in conjunction with your existing classes solely for the purpose of schema migrations and management.

Easier Inserts:

If you spend a lot of time inserting data into your database and have an object mapped to a SQLdeLite.Table then inserting is super easy. For instance you can do:

Dim newUser As New User(myDb)
newUser.Username = "testuser"
newUser.Password = "hashedPassword"
newUser.Insert()

This is a nice little helper to save you some time. SQL Insert statements are largely the same and theres no point wasting time writing methods to insert every known possibility.

Easier, standardized cross-db Prepared Statements syntax:

You don’t need multiple parameter insertion points in your SQL (those “?”) and a separated list of variables, the object can do it at once as:

newUser.Execute("UPDATE Users SET Password = #password# WHERE Username = #username#")

So while we don't like full on conventional ORM's we can use the schema info we already have to enhance the experience.
We do this without boxing you into any corners, paradigms, etc.

SQLdeLite does not write your queries for you. It just makes using SQL significantly more enjoyable.

Enjoy!

1701 Software - http://www.1701software.com/
