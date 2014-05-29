#tag Class
Protected Class DatabaseSQLite
Inherits SQLdeLite.DatabaseCore
Implements i_Database
	#tag Event
		Sub DoClose()
		  // Verify we are connected to the SQLite database
		  If (p_isConnected = True) Then
		    
		    p_database.Close()
		    p_isConnected = False
		    
		  Else
		    Dim connectError As New SQLdeLiteException
		    connectError.Message = "Not connected to a SQLdeLite.DatabaseSQLite"
		    Raise connectError
		  End If
		End Sub
	#tag EndEvent

	#tag Event
		Sub DoCommit()
		  // Verify we are connected to the SQLite database
		  If (p_isConnected = True) Then
		    
		    p_database.Commit()
		    
		  Else
		    Dim connectError As New SQLdeLiteException
		    connectError.Message = "Not connected to a SQLdeLite.DatabaseSQLite"
		    Raise connectError
		  End If
		End Sub
	#tag EndEvent

	#tag Event
		Function DoConnect() As Boolean
		  // Make sure they provided a valid FolderItem
		  If (p_database.DatabaseFile = Nil) Then
		    Dim error As New SQLdeLiteException
		    error.Message = "You have not provided a valid SQLdeLite.DatabaseSQLite.DatabaseFile"
		    Raise error
		    Return False
		  End If
		  
		  // Make sure the FolderItem exists
		  If (p_database.DatabaseFile.Exists() = False) Then
		    If (p_database.CreateDatabaseFile() = False) Then
		      Dim error As New SQLdeLiteException
		      error.Message = "Unable to create DatabaseSQLite at '" + p_database.DatabaseFile.NativePath + "'"
		      Raise error
		    End If
		  End If
		  
		  If (p_database.Connect() = True) Then
		    p_isConnected = True
		    Return True
		  Else
		    Return False
		  End If
		End Function
	#tag EndEvent

	#tag Event
		Sub DoCreateTable(TableName As String, PrimaryKeyColumnName As String, BindType As Integer = -1)
		  // Verify we are connected to the SQLite database
		  If (p_isConnected = True) Then
		    
		    // Find cached table information
		    Dim cachedTable As c_TableCache
		    cachedTable = m_getCachedTable(TableName)
		    
		    // Determine if user passed in a custom BindType for Primary Key
		    Dim createColumnType As Integer
		    If (BindType = -1) Then
		      // Determine the data type of the primary key
		      For Each prop As c_PropertyCache In cachedTable.Properties
		        // Determine if this property name is the same as the PrimaryKeyColumnName identified by the developer.
		        If (prop.PropertyName = PrimaryKeyColumnName) Then
		          // Determine what type of binding this is.
		          createColumnType = m_ClassTypeToBindType(prop.PropertyType)
		        End If
		      Next
		    Else
		      // User passed in a custom SQL BindType
		      createColumnType = BindType
		    End If
		    
		    // Determine the english variant of the primary key bind type
		    Dim createColumnTypeString As String
		    createColumnTypeString = m_BindTypeToSQLType(createColumnType)
		    If (createColumnTypeString = "") Then
		      Dim createPrimaryKeyError As New SQLdeLiteException
		      createPrimaryKeyError.Message = "Invalid primary key type."
		      Raise createPrimaryKeyError
		      Return
		    End If
		    
		    // Create the SQL for creating the table
		    Dim createTableSQL As String
		    createTableSQL = "CREATE TABLE " + TableName + " (" + PrimaryKeyColumnName + " " + createColumnTypeString + " PRIMARY KEY);"
		    SQLExecute(createTableSQL)
		    
		  Else
		    Dim connectError As New SQLdeLiteException
		    connectError.Message = "Not connected to a SQLdeLite.DatabaseSQLite"
		    Raise connectError
		  End If
		  
		End Sub
	#tag EndEvent

	#tag Event
		Sub DoInsertRecord(TableName As String, Data As DatabaseRecord)
		  // Verify we are connected to the SQLite database
		  If (p_isConnected = True) Then
		    
		    p_database.InsertRecord(TableName, Data)
		    
		  Else
		    Dim connectError As New SQLdeLiteException
		    connectError.Message = "Not connected to a SQLdeLite.DatabaseSQLite"
		    Raise connectError
		  End If
		End Sub
	#tag EndEvent

	#tag Event
		Function DoPrepare(Statement As String) As PreparedSQLStatement
		  // Verify we are connected to the SQLite database
		  If (p_isConnected = True) Then
		    
		    Return p_database.Prepare(Statement)
		    
		  Else
		    Dim connectError As New SQLdeLiteException
		    connectError.Message = "Not connected to a SQLdeLite.DatabaseSQLite"
		    Raise connectError
		  End If
		End Function
	#tag EndEvent

	#tag Event
		Sub DoRollback()
		  // Verify we are connected to the SQLite database
		  If (p_isConnected = True) Then
		    
		    p_database.Rollback()
		    
		  Else
		    Dim connectError As New SQLdeLiteException
		    connectError.Message = "Not connected to a SQLdeLite.DatabaseSQLite"
		    Raise connectError
		  End If
		End Sub
	#tag EndEvent

	#tag Event
		Function DoSQLdeLiteParams(Query As String, Values() As Parameter, ReturnRecordSet As Boolean = False) As RecordSet
		  // Verify we are connected to the SQLite database
		  If (p_isConnected = True) Then
		    
		    Return m_buildAndQuery(Query, Values, ReturnRecordSet)
		    
		  Else
		    Dim connectError As New SQLdeLiteException
		    connectError.Message = "Not connected to a SQLdeLite.DatabaseSQLite"
		    Raise connectError
		  End If
		End Function
	#tag EndEvent

	#tag Event
		Function DoSQLdeLiteVariants(Query As String, Values() As Variant, ReturnRecordSet As Boolean = False) As RecordSet
		  // Verify we are connected to the SQLite database
		  If (p_isConnected = True) Then
		    
		    Dim props() As Parameter
		    props = m_convertVariantsToParams(Values)
		    
		    Return m_buildAndQuery(Query, props, ReturnRecordSet)
		    
		  Else
		    Dim connectError As New SQLdeLiteException
		    connectError.Message = "Not connected to a SQLdeLite.DatabaseSQLite"
		    Raise connectError
		  End If
		End Function
	#tag EndEvent

	#tag Event
		Function DoSQLSelect(SQL As String) As RecordSet
		  // Verify we are connected to the SQLite database
		  If (p_isConnected = True) Then
		    
		    Return p_database.SQLSelect(SQL)
		    
		  Else
		    Dim connectError As New SQLdeLiteException
		    connectError.Message = "Not connected to a SQLdeLite.DatabaseSQLite"
		    Raise connectError
		    Return Nil
		  End If
		End Function
	#tag EndEvent

	#tag Event
		Function DoVariantsToParameters(Values() As Variant) As Parameter()
		  // We need to convert the values to useful types.
		  Dim props() As Parameter
		  For Each Value As Variant In Values
		    If (Value.Type = Variant.TypeDouble) Then // Double
		      Dim param As New Parameter(Value, m_classTypeToBindType("Double"))
		      props.Append(param)
		    ElseIf (Value.Type = Variant.TypeInteger) Then // Integer
		      Dim param As New Parameter(Value, m_classTypeToBindType("Int32"))
		      props.Append(param)
		    ElseIf (Value.Type = Variant.TypeString) Then // String
		      Dim param As New Parameter(Value, m_classTypeToBindType("String"))
		      props.Append(param)
		    End If
		  Next
		  
		  Return props
		End Function
	#tag EndEvent

	#tag Event
		Sub OnCreateTableColumn(TableName As String, ColumnName As String, BindType As Integer = -1)
		  // Verify we are connected to the SQLite database
		  If (p_isConnected = True) Then
		    
		    // Find cached table information
		    Dim cachedTable As c_TableCache
		    cachedTable = m_getCachedTable(TableName)
		    
		    // Determine if user passed in a custom BindType for this column.
		    // Determine the data type of the column
		    Dim createColumnType As Integer
		    If (BindType = -1) Then
		      For Each prop As c_PropertyCache In cachedTable.Properties
		        // Determine if this property name is the same as the PrimaryKeyColumnName identified by the developer.
		        If (prop.PropertyName = ColumnName) Then
		          // Determine what type of binding this is.
		          createColumnType = m_ClassTypeToBindType(prop.PropertyType)
		        End If
		      Next
		    Else
		      // User passed in a custom SQL BindType
		      createColumnType = BindType
		    End If
		    
		    // Determine the english variant of the primary key bind type
		    Dim createColumnTypeString As String
		    createColumnTypeString = m_BindTypeToSQLType(createColumnType)
		    If (createColumnTypeString = "") Then
		      Dim createPrimaryKeyError As New SQLdeLiteException
		      createPrimaryKeyError.Message = "Invalid property type."
		      Raise createPrimaryKeyError
		      Return
		    End If
		    
		    // Create the SQL for creating the table
		    Dim createTableSQL As String
		    createTableSQL = "ALTER TABLE " + TableName + " ADD " + ColumnName + " " + createColumnTypeString
		    SQLExecute(createTableSQL)
		    
		  Else
		    Dim connectError As New SQLdeLiteException
		    connectError.Message = "Not connected to a SQLdeLite.DatabaseSQLite"
		    Raise connectError
		  End If
		End Sub
	#tag EndEvent

	#tag Event
		Sub OnSQLExecute(SQL As String)
		  // Verify we are connected to the SQLite database
		  If (p_isConnected = True) Then
		    
		    p_database.SQLExecute(SQL)
		    
		  Else
		    Dim connectError As New SQLdeLiteException
		    connectError.Message = "Not connected to a SQLdeLite.DatabaseSQLite"
		    Raise connectError
		    Return
		  End If
		  
		  
		End Sub
	#tag EndEvent


	#tag Method, Flags = &h0
		Sub Backup(Destination As SQLiteDatabase, callBackHandler As SQLiteBackupInterface = Nil, sleepTimeInMilliseconds As Integer = 10)
		  // Verify we are connected to the SQLite database
		  If (p_isConnected = True) Then
		    
		    p_database.BackUp(Destination, callBackHandler, sleepTimeInMilliseconds)
		    
		  Else
		    Dim connectError As New SQLdeLiteException
		    connectError.Message = "Not connected to a SQLdeLite.DatabaseSQLite"
		    Raise connectError
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor()
		  // Call Super Constructor
		  Super.Constructor()
		  
		  // Initialize variables
		  p_database = New SQLiteDatabase()
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CreateBlob(TableName As String, ColumnName As String, Row As UInt64, Length As UInt64, DatabaseName As String) As SQLiteBlob
		  // Verify we are connected to the SQLite database
		  If (p_isConnected = True) Then
		    
		    Return p_database.CreateBlob(TableName, ColumnName, Row, Length, DatabaseName)
		    
		  Else
		    Dim connectError As New SQLdeLiteException
		    connectError.Message = "Not connected to a SQLdeLite.DatabaseSQLite"
		    Raise connectError
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CreateDatabaseFile() As Boolean
		  Return p_database.CreateDatabaseFile()
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Decrypt()
		  p_database.Decrypt()
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Encrypt(Password As String)
		  p_database.Encrypt(Password)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub InsertRecord(TableName As String, Data As DatabaseRecord)
		  // Verify we are connected to the SQLite database
		  If (p_isConnected = True) Then
		    
		    p_database.InsertRecord(TableName, Data)
		    
		  Else
		    Dim connectError As New SQLdeLiteException
		    connectError.Message = "Not connected to a SQLdeLite.DatabaseSQLite"
		    Raise connectError
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function LastRowID() As Int64
		  Return p_database.LastRowID
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function m_bindTypeToSQLType(BindType As Integer) As String
		  If (BindType = SQLitePreparedStatement.SQLITE_INTEGER) Then // Integer
		    Return "INTEGER"
		  ElseIf (BindType = SQLitePreparedStatement.SQLITE_TEXT) Then // String
		    Return "TEXT"
		  ElseIf (BindType = SQLitePreparedStatement.SQLITE_BLOB) Then // Blob
		    Return "BLOB"
		  ElseIf (BindType = SQLitePreparedStatement.SQLITE_NULL) Then // NULL
		    Return "NULL"
		  ElseIf (BindType = SQLitePreparedStatement.SQLITE_DOUBLE) Then // Double
		    Return "REAL"
		  Else
		    Return ""
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub m_buildAndInsertObject(ClassName As String, Query As String, Values As Dictionary)
		  // Verify we are connected to the SQLite database
		  If (p_isConnected = True) Then
		    
		    // Create SQLite Prepared Statement
		    Dim ps As SQLitePreparedStatement
		    ps = p_database.Prepare(Query)
		    
		    // Bind all the properties to their bind types.
		    Dim propCount As Integer = 0
		    For Each prop1 As Variant In Values.Keys
		      // Determine the property type
		      Dim tableCache As c_TableCache
		      Dim propType As Integer
		      tableCache = p_zTableCache.Value(ClassName)
		      For Each prop As c_PropertyCache In tableCache.Properties
		        If (prop.PropertyName = prop1) Then
		          propType = m_classTypeToBindType(prop.PropertyType)
		        End If
		      Next
		      // Set the bind type
		      ps.BindType(propCount, propType)
		      // Set the bind value
		      If (propType = SQLitePreparedStatement.SQLITE_BLOB) Then
		        // Not implemented yet.
		      ElseIf (propType = SQLitePreparedStatement.SQLITE_DOUBLE) Then
		        ps.Bind(propCount, Values.Value(prop1).DoubleValue)
		      ElseIf (propType = SQLitePreparedStatement.SQLITE_INTEGER) Then
		        ps.Bind(propCount, Values.Value(prop1).IntegerValue)
		      ElseIf (propType = SQLitePreparedStatement.SQLITE_NULL) Then
		        // Not implemented yet.
		      ElseIf (propType = SQLitePreparedStatement.SQLITE_TEXT) Then
		        ps.Bind(propCount, Values.Value(prop1).StringValue)
		      End If
		      // Increment the propCount
		      propCount = propCount + 1
		    Next
		    
		    // Execute the given prepared statement
		    ps.SQLExecute()
		    
		    // Update Error Codes
		    m_updateErrorCodes()
		    
		  Else
		    Dim connectError As New SQLdeLiteException
		    connectError.Message = "Not connected to a SQLdeLite.DatabaseSQLite"
		    Raise connectError
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function m_buildAndQuery(Query As String, Values() As Parameter, ReturnRecordSet As Boolean) As RecordSet
		  // Verify we are connected to the SQLite database
		  If (p_isConnected = True) Then
		    
		    // Create array of SQL parameters
		    Dim params() As Parameter
		    
		    // Create new SQL query
		    Dim newQuery As String
		    newQuery = Query
		    
		    // Loop through the values passed to us that aren't property related.
		    Dim valueCount As Integer = 0
		    While InStr(0, newQuery, "#?#") > 0
		      // Find where the variable is in the query.
		      Dim location As Integer = InStr(0, newQuery, "#?#")
		      newQuery = Replace(newQuery, "#?#", "?")
		      // Add nil values to array
		      While params.Ubound < location
		        params.Append(Nil)
		      Wend
		      // Add parameter to array
		      params.Insert(location, Values(valueCount))
		      // Increment count so we look at that right position in the Values() array
		      valueCount = valueCount + 1
		    Wend
		    
		    // Create SQLite Prepared Statement
		    Dim ps As SQLitePreparedStatement
		    ps = p_database.Prepare(newQuery)
		    
		    // Bind all the properties to their bind types in order that they are listed in the query.
		    Dim paramCount As Integer = 0
		    For X As Integer = 0 To params.Ubound
		      // Make sure there is an actual zSQLParameter here.
		      If (params(X) <> Nil) Then
		        // Get the zSQLParameter
		        Dim tempSQL As Parameter
		        tempSQL = params(X)
		        // Create binding for this parameter
		        ps.BindType(paramCount, tempSQL.BindType)
		        // Set the bind value
		        If (tempSQL.BindType = SQLitePreparedStatement.SQLITE_BLOB) Then
		          // Not implemented yet.
		        ElseIf (tempSQL.BindType = SQLitePreparedStatement.SQLITE_DOUBLE) Then
		          ps.Bind(paramCount, tempSQL.Value.DoubleValue)
		        ElseIf (tempSQL.BindType = SQLitePreparedStatement.SQLITE_INTEGER) Then
		          ps.Bind(paramCount, tempSQL.Value.IntegerValue)
		        ElseIf (tempSQL.BindType = SQLitePreparedStatement.SQLITE_NULL) Then
		          // Not implemented yet.
		        ElseIf (tempSQL.BindType = SQLitePreparedStatement.SQLITE_TEXT) Then
		          ps.Bind(paramCount, tempSQL.Value.StringValue)
		        End If
		        // Increment paramCount
		        paramCount = paramCount + 1
		      End If
		    Next
		    
		    // Execute the given prepared statement
		    If (ReturnRecordSet = True) Then
		      Return ps.SQLSelect()
		    Else
		      ps.SQLExecute()
		      Return Nil
		    End If
		    
		    // Update Error Codes
		    m_updateErrorCodes()
		    
		    
		  Else
		    Dim connectError As New SQLdeLiteException
		    connectError.Message = "Not connected to a SQLdeLite.DatabaseSQLite"
		    Raise connectError
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub m_buildAndUpdateObject(ClassName As String, TableName As String, Query As String, Values() As Parameter, PropValues As Dictionary)
		  // Verify we are connected to the SQLite database
		  If (p_isConnected = True) Then
		    
		    // Let's fetch the table cache information
		    Dim tableCache As c_TableCache
		    tableCache = p_zTableCache.Value(className)
		    
		    // Create array of SQL parameters
		    Dim params() As Parameter
		    
		    // Create new SQL query
		    Dim newQuery As String
		    newQuery = Query
		    
		    // Loop through class properties to determine if the value needs to be added to the array.
		    Dim addCount As Integer
		    addCount = 0
		    For Each prop As c_PropertyCache In tableCache.Properties
		      // Create property bindtype
		      Dim propType As Integer
		      // Create replacement variable. ie: if property = Name, then create #name# so we can determine if it exists.
		      Dim propString As String
		      propString = "#" + Lowercase(prop.PropertyName) + "#"
		      // Find all instances of the property
		      While InStr(0, newQuery, propString) > 0
		        // Find where the variable is in the query.
		        Dim location As Integer = InStr(0, newQuery, propString)
		        newQuery = Replace(newQuery, propString, "?")
		        // Add nil values to array
		        While params.Ubound < location
		          params.Append(Nil)
		        Wend
		        propType = m_classTypeToBindType(prop.PropertyType)
		        // Create parameter and add to array
		        Dim tempParam As New Parameter(PropValues.Value(prop.PropertyName), propType)
		        params.Insert(location, tempParam)
		      Wend
		    Next
		    
		    // Loop through the values passed to us that aren't property related.
		    Dim valueCount As Integer = 0
		    While InStr(0, newQuery, "#?#") > 0
		      // Find where the variable is in the query.
		      Dim location As Integer = InStr(0, newQuery, "#?#")
		      newQuery = Replace(newQuery, "#?#", "?")
		      // Add nil values to array
		      While params.Ubound < location
		        params.Append(Nil)
		      Wend
		      // Add parameter to array
		      params.Insert(location, Values(valueCount))
		      // Increment count so we look at that right position in the Values() array
		      valueCount = valueCount + 1
		    Wend
		    
		    // Create SQLite Prepared Statement
		    Dim ps As SQLitePreparedStatement
		    ps = p_database.Prepare(newQuery)
		    
		    // Bind all the properties to their bind types in order that they are listed in the query.
		    Dim paramCount As Integer = 0
		    For X As Integer = 0 To params.Ubound
		      // Make sure there is an actual zSQLParameter here.
		      If (params(X) <> Nil) Then
		        // Get the zSQLParameter
		        Dim tempSQL As Parameter
		        tempSQL = params(X)
		        // Create binding for this parameter
		        ps.BindType(paramCount, tempSQL.BindType)
		        // Set the bind value
		        If (tempSQL.BindType = SQLitePreparedStatement.SQLITE_BLOB) Then
		          // Not implemented yet.
		        ElseIf (tempSQL.BindType = SQLitePreparedStatement.SQLITE_DOUBLE) Then
		          ps.Bind(paramCount, tempSQL.Value.DoubleValue)
		        ElseIf (tempSQL.BindType = SQLitePreparedStatement.SQLITE_INTEGER) Then
		          ps.Bind(paramCount, tempSQL.Value.IntegerValue)
		        ElseIf (tempSQL.BindType = SQLitePreparedStatement.SQLITE_NULL) Then
		          // Not implemented yet.
		        ElseIf (tempSQL.BindType = SQLitePreparedStatement.SQLITE_TEXT) Then
		          ps.Bind(paramCount, tempSQL.Value.StringValue)
		        End If
		        // Increment paramCount
		        paramCount = paramCount + 1
		      End If
		    Next
		    
		    // Execute the given prepared statement
		    ps.SQLExecute()
		    
		    // Update Error Codes
		    m_updateErrorCodes()
		    
		  Else
		    Dim connectError As New SQLdeLiteException
		    connectError.Message = "Not connected to a SQLdeLite.DatabaseSQLite"
		    Raise connectError
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function m_checkTableExists(TableName As String) As Boolean
		  // Verify we are connected to the SQLite database
		  If (p_isConnected = True) Then
		    
		    // Check to determine if a specific table exists in the SQLite Database
		    Dim verifyTablePS As SQLitePreparedStatement
		    verifyTablePS = p_database.Prepare("SELECT * FROM sqlite_master WHERE TYPE = 'table' AND tbl_name = ?")
		    verifyTablePS.BindType(0, SQLitePreparedStatement.SQLITE_TEXT)
		    verifyTablePS.Bind(0, TableName)
		    
		    Dim verifyTableRS As RecordSet
		    verifyTableRS = verifyTablePS.SQLSelect()
		    
		    If (verifyTableRS = Nil Or verifyTableRS.RecordCount = 0) Then
		      Return False
		    Else
		      Return True
		    End If
		    
		  Else
		    Dim connectError As New SQLdeLiteException
		    connectError.Message = "Not connected to a SQLdeLite.DatabaseSQLite"
		    Raise connectError
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function m_classTypeToBindType(PropertyType As String) As Integer
		  // This method helps the Table constructor create a hash representation of the classes
		  // properties and types. It looks up the generic Xojo types in order identify what the
		  // default SQLite property type should be.
		  
		  // Determine PropertyType for SQLiteDatabase
		  If (PropertyType = "Date") Then
		    Return SQLitePreparedStatement.SQLITE_TEXT
		  ElseIf (PropertyType = "Double") Then
		    Return SQLitePreparedStatement.SQLITE_DOUBLE
		  ElseIf (PropertyType = "Int32") Then
		    Return SQLitePreparedStatement.SQLITE_INTEGER
		  ElseIf (PropertyType = "String") Then
		    Return SQLitePreparedStatement.SQLITE_TEXT
		  Else
		    Dim bindTypeError As New SQLdeLiteException
		    bindTypeError.Message = "No mapping for column property type to SQLite"
		    Raise bindTypeError
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub m_insertRow(ClassName As String, TableName As String, Values As Dictionary)
		  // Verify we are connected to the SQLite database
		  If (p_isConnected = True) Then
		    
		    // Fetch the primary key column name
		    Dim primaryKey As String
		    primaryKey = c_TableCache(p_zTableCache.Value(className)).PrimaryKeyColumnName
		    
		    // Let's remove the primary key from the array of values
		    Dim hasPrimary As Boolean = False
		    For Each prop1 As Variant In Values.Keys
		      If (prop1 = primaryKey) Then
		        hasPrimary = True
		      End If
		    Next
		    If (hasPrimary = True) Then
		      Values.Remove(primaryKey)
		    End If
		    
		    // Create the SQL statement to insert the values
		    Dim insertSQL() As String
		    insertSQL.Append("INSERT INTO ")
		    insertSQL.Append(TableName)
		    insertSQL.Append(" (")
		    For Each prop1 As Variant In Values.Keys
		      insertSQL.Append(prop1)
		      insertSQL.Append(",")
		    Next
		    insertSQL.Remove(insertSQL.Ubound) // Remove the last comma
		    insertSQL.Append(") VALUES (")
		    // Add appropriate number of question marks
		    For X As Integer = 0 To Values.Count - 1
		      insertSQL.Append("?")
		      insertSQL.Append(",")
		    Next
		    insertSQL.Remove(insertSQL.Ubound) // Remove the last comma
		    insertSQL.Append(")")
		    
		    // Build and execute query via prepared statement
		    m_buildAndInsertObject(ClassName, Join(insertSQL, ""), Values)
		    
		  Else
		    Dim connectError As New SQLdeLiteException
		    connectError.Message = "Not connected to a SQLdeLite.DatabaseSQLite"
		    Raise connectError
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function m_SQLTypeToBindType(SQLType As String) As Integer
		  If (SQLType = "INTEGER") Then
		    Return SQLitePreparedStatement.SQLITE_INTEGER
		  ElseIf (SQLType = "TEXT") Then
		    Return SQLitePreparedStatement.SQLITE_TEXT
		  ElseIf (SQLType = "REAL") Then
		    Return SQLitePreparedStatement.SQLITE_DOUBLE
		  ElseIf (SQLType = "BLOB") Then
		    Return SQLitePreparedStatement.SQLITE_BLOB
		  ElseIf (SQLType = "NULL") Then
		    Return SQLitePreparedStatement.SQLITE_NULL
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub m_updateErrorCodes()
		  Error = p_database.Error
		  ErrorCode = p_database.ErrorCode
		  ErrorMessage = p_database.ErrorMessage
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub m_verifyTableSchema(className As String)
		  // Verify we are connected to the SQLite database
		  If (p_isConnected = True) Then
		    
		    // Let's verify that the table exists in the database.
		    Dim verifyTableCache As c_TableCache
		    verifyTableCache = p_zTableCache.Value(className)
		    
		    If (m_CheckTableExists(verifyTableCache.TableName) = False) Then
		      Dim userHandledCreateTable As Boolean = False
		      userHandledCreateTable = CreateTableSchema(verifyTableCache.TableName)
		      If (userHandledCreateTable = False) Then
		        // User did not handle the CreateTable event.
		        Dim verifyTableError As New SQLdeLiteException
		        verifyTableError.Message = "You must handle the CreateTableSchema event for your DatabaseSQLite object."
		        Raise verifyTableError
		      ElseIf (userHandledCreateTable = True) Then
		        // User supposedly handled the event. Let's verify...
		        If (m_CheckTableExists(verifyTableCache.TableName) = False) Then
		          Dim verifyTableError As New SQLdeLiteException
		          verifyTableError.Message = "You must handle the CreateTableSchema event for your DatabaseSQLite object."
		          Raise verifyTableError
		        End If
		      End If
		    End If
		    
		    // Let's determine if the table is already verified.
		    If (verifyTableCache.Verified = False) Then
		      
		      // Let's verify that the database schema matches the class schema.
		      // This is run before we have asked the user what the bindtypes are.
		      Dim tableSchemaHash As String
		      
		      Dim tableSchemaSQL As String
		      tableSchemaSQL = "PRAGMA table_info('" + verifyTableCache.TableName + "');"
		      Dim tableSchemaRS As RecordSet
		      tableSchemaRS = p_database.SQLSelect(tableSchemaSQL)
		      
		      // Loop through table and fetch table names.
		      Dim columnNames() As String
		      While Not tableSchemaRS.EOF
		        columnNames.Append(Lowercase(tableSchemaRS.Field("name").StringValue))
		        tableSchemaRS.MoveNext()
		      Wend
		      columnNames.Sort()
		      
		      // Create JSONItem representing the table schema
		      Dim tableSchemaPropertiesJSON As New JSONItem
		      
		      // Loop through table names and add values to table schema
		      For Each col As String In columnNames
		        // Move to first column
		        tableSchemaRS.MoveFirst()
		        
		        While Not tableSchemaRS.EOF
		          If (Lowercase(tableSchemaRS.Field("name").StringValue) = col) Then
		            // Add Properties to JSON cache
		            Dim propJSON As New JSONItem
		            
		            Dim propName As String
		            Dim propType As Integer
		            propName = Lowercase(tableSchemaRS.Field("name").StringValue)
		            propType = m_SQLTypeToBindType(tableSchemaRS.Field("type").StringValue)
		            
		            propJSON.Value(propName) = propType
		            
		            // Determine if this property is the primary key.
		            If (tableSchemaRS.Field("pk").IntegerValue = 1) Then
		              verifyTableCache.PrimaryKeyColumnName = tableSchemaRS.Field("name").StringValue
		            End If
		            
		            tableSchemaPropertiesJSON.Append(propJSON)
		          End If
		          
		          // Move to next column.
		          tableSchemaRS.MoveNext()
		        Wend
		      Next
		      
		      // Hash the JSON results of the table
		      tableSchemaHash = EncodeHex(MD5(tableSchemaPropertiesJSON.ToString()))
		      
		      // Clear columnNames array
		      Redim columnNames(-1)
		      
		      // Loop through class and fetch table names.
		      For P As Integer = 0 To verifyTableCache.Properties.Ubound
		        columnNames.Append(Lowercase(verifyTableCache.Properties(P).PropertyName))
		      Next
		      columnNames.Sort()
		      
		      // Let's guess what the initial bind types should be for the zTableCache object.
		      Dim classSchemaHash As String
		      
		      // Create JSONItem representing the class schema
		      Dim classSchemaPropertiesJSON As New JSONItem
		      
		      For X As Integer = 0 To columnNames.Ubound
		        // Loop through properties to see if it matches the one in this array
		        For P As Integer = 0 To verifyTableCache.Properties.Ubound
		          If (columnNames(X) = Lowercase(verifyTableCache.Properties(P).PropertyName)) Then
		            Dim propJSON As New JSONItem
		            
		            Dim propName As String
		            Dim propType As Integer
		            propName = Lowercase(verifyTableCache.Properties(P).PropertyName)
		            propType = m_ClassTypeToBindType(verifyTableCache.Properties(P).PropertyType)
		            
		            propJSON.Value(propName) = propType
		            
		            classSchemaPropertiesJSON.Append(propJSON)
		          End If
		        Next
		      Next
		      
		      // Hash the JSON results of the class
		      classSchemaHash = EncodeHex(MD5(classSchemaPropertiesJSON.ToString()))
		      
		      // Compare the JSON results of the class and table. Raise the UpdateTableSchema if it does not match.
		      If (tableSchemaHash <> classSchemaHash) Then
		        Dim userHandledUpdateTable As Boolean = False
		        userHandledUpdateTable = UpdateTableSchema(verifyTableCache.TableName)
		        If (userHandledUpdateTable = False) Then
		          // User did not handle the CreateTable event.
		          Dim verifyTableError As New SQLdeLiteException
		          verifyTableError.Message = "You must handle the UpdateTableSchema event for your DatabaseSQLite object."
		          Raise verifyTableError
		        End If
		      End If
		      
		      // Mark this table as verified
		      verifyTableCache.Verified = True
		      
		    End If
		    
		  Else
		    Dim connectError As New SQLdeLiteException
		    connectError.Message = "Not connected to a SQLdeLite.DatabaseSQLite"
		    Raise connectError
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function OpenBlob(TableName As String, ColumnName As String, Row As UInt64, ReadWrite As Boolean, DatabaseName As String = "") As SQLiteBlob
		  // Verify we are connected to the SQLite database
		  If (p_isConnected = True) Then
		    
		    Return p_database.OpenBlob(TableName, ColumnName, Row, ReadWrite, DatabaseName)
		    
		  Else
		    Dim connectError As New SQLdeLiteException
		    connectError.Message = "Not connected to a SQLdeLite.DatabaseSQLite"
		    Raise connectError
		  End If
		End Function
	#tag EndMethod


	#tag Hook, Flags = &h0
		Event CreateTableSchema(TableName As String) As Boolean
	#tag EndHook

	#tag Hook, Flags = &h0
		Event UpdateTableSchema(TableName As String) As Boolean
	#tag EndHook


	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  Return p_database.DatabaseFile
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  p_database.DatabaseFile = value
			End Set
		#tag EndSetter
		DatabaseFile As FolderItem
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  Return p_database.EncryptionKey
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  p_database.EncryptionKey = value
			End Set
		#tag EndSetter
		EncryptionKey As String
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  return p_database.LibraryVersion
			End Get
		#tag EndGetter
		LibraryVersion As String
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  Return p_database.MultiUser
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  p_database.MultiUser = value
			End Set
		#tag EndSetter
		MultiUser As Boolean
	#tag EndComputedProperty

	#tag Property, Flags = &h21
		Private p_database As SQLiteDatabase
	#tag EndProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  Return p_database.ShortColumnNames
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  p_database.ShortColumnNames = value
			End Set
		#tag EndSetter
		ShortColumnNames As Boolean
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  Return p_database.ThreadYieldInterval
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  p_database.ThreadYieldInterval = value
			End Set
		#tag EndSetter
		ThreadYieldInterval As Integer
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  Return p_database.Timeout
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  p_database.Timeout = value
			End Set
		#tag EndSetter
		Timeout As Double
	#tag EndComputedProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="EncryptionKey"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Error"
			Group="Behavior"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ErrorCode"
			Group="Behavior"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ErrorMessage"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Group="Position"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="LibraryVersion"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="MultiUser"
			Group="Behavior"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ShortColumnNames"
			Group="Behavior"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ThreadYieldInterval"
			Group="Behavior"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Timeout"
			Group="Behavior"
			Type="Double"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Group="Position"
			Type="Integer"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
