#tag Module
Protected Module SQLdeLite
	#tag Method, Flags = &h21
		Private Function mQuery(db As Database, SelectText As Text, Parameters() As Auto, ExecuteOnly As Boolean = False) As RecordSet
		  // Determine if we had any parameters subbed in. If not then call the underlying database SQLSelect method as there is nothing to process.
		  If (Parameters.Ubound = -1) Then
		    Return db.SQLSelect(SelectText)
		  End If
		  
		  // Create instance of PreparedSQLStatement
		  Dim _ps As PreparedSQLStatement
		  #If SQLDeLite.PLUGIN_CUBESQL_ENABLED = True Then
		    Dim _psCube As CubeSQLVM
		  #EndIf
		  
		  // Determine what database engine we are on.
		  Dim _dbInfo As Xojo.Introspection.TypeInfo
		  _dbInfo = Xojo.Introspection.GetType(db)
		  
		  // Create a prepared SQL statement appropriate for the current database engine.
		  If (_dbInfo.FullName = "CubeSQLServer") Then
		    #If SQLdeLite.PLUGIN_CUBESQL_ENABLED = True Then
		      _psCube = CubeSQLServer(db).VMPrepare(SelectText)
		    #EndIf
		  Else
		    _ps = db.Prepare(SelectText)
		  End If
		  
		  // Loop through the parameters and determine the appropriate database types for binding to the prepared statement.
		  For _count As Integer = 0 To Parameters.Ubound
		    
		    // SQLiteDatabase
		    If (_dbInfo.FullName = "SQLiteDatabase") Then
		      
		      // Determine what type of field this is.
		      Dim __parameterInfo As Xojo.Introspection.TypeInfo
		      __parameterInfo = Xojo.Introspection.GetType(Parameters(_count))
		      
		      If (__parameterInfo.FullName = "Boolean") Then
		        _ps.BindType(_count, SQLitePreparedStatement.SQLITE_BOOLEAN)
		      ElseIf (__parameterInfo.FullName = "Double") Then
		        _ps.BindType(_count, SQLitePreparedStatement.SQLITE_DOUBLE)
		      ElseIf (__parameterInfo.FullName = "Int64") Then
		        _ps.BindType(_count, SQLitePreparedStatement.SQLITE_INT64)
		      ElseIf (__parameterInfo.FullName = "Int32") Then
		        _ps.BindType(_count, SQLitePreparedStatement.SQLITE_INTEGER)
		      ElseIf (__parameterInfo.FullName = "Text") Then
		        _ps.BindType(_count, SQLitePreparedStatement.SQLITE_TEXT)
		      End If
		      
		    End If
		    
		    // cubeSQLServer
		    #If SQLdeLite.PLUGIN_CUBESQL_ENABLED = True Then
		      
		      If (_dbInfo.FullName = "CubeSQLServer") Then
		        
		        // Remember that the CubeSQLVM object is 1-based for binding. So it's "count + 1".
		        
		        // Determine what type of field this is.
		        Dim __parameterInfo As Xojo.Introspection.TypeInfo
		        __parameterInfo = Xojo.Introspection.GetType(Parameters(_count))
		        
		        If (__parameterInfo.FullName = "Double") Then
		          _psCube.BindDouble(_count + 1, Parameters(_count))
		        ElseIf (__parameterInfo.FullName = "Int64") Then
		          _psCube.BindInt64(_count + 1, Parameters(_count))
		        ElseIf (__parameterInfo.FullName = "Int32") Then
		          _psCube.BindInt(_count + 1, Parameters(_count))
		        ElseIf (__parameterInfo.FullName = "Text") Then
		          _psCube.BindText(_count + 1, Parameters(_count))
		        End If
		        
		      End If
		      
		    #EndIf
		    
		    // MySQLCommunityServer
		    #If SQLDeLite.PLUGIN_MYSQL_ENABLED = True Then
		      
		      If (_dbInfo.FullName = "MySQLCommunityServer") Then
		        
		        // Determine what type of field this is.
		        Dim __parameterInfo As Xojo.Introspection.TypeInfo
		        __parameterInfo = Xojo.Introspection.GetType(Parameters(_count))
		        
		        If (__parameterInfo.FullName = "Date") Then
		          _ps.BindType(_count, MySQLPreparedStatement.MYSQL_TYPE_DATE)
		        ElseIf (__parameterInfo.FullName = "Double") Then
		          _ps.BindType(_count, MySQLPreparedStatement.MYSQL_TYPE_DOUBLE)
		        ElseIf (__parameterInfo.FullName = "Int64") Then
		          _ps.BindType(_count, MySQLPreparedStatement.MYSQL_TYPE_LONGLONG)
		        ElseIf (__parameterInfo.FullName = "Int32") Then
		          _ps.BindType(_count, MySQLPreparedStatement.MYSQL_TYPE_LONG)
		        ElseIf (__parameterInfo.FullName = "Text") Then
		          _ps.BindType(_count, MySQLPreparedStatement.MYSQL_TYPE_STRING)
		        End If
		        
		      End If
		      
		    #EndIf
		    
		    // PostgreSQLDatabase
		    #If SQLDeLite.PLUGIN_POSTGRESQL_ENABLED = True Then
		      
		      If (_dbInfo.FullName = "PostgreSQLDatabase") Then
		        
		        _ps.Bind(_count, Parameters(_count))
		        
		      End If
		      
		    #EndIf
		    
		    // ODBCDatabase
		    #If SQLDeLite.PLUGIN_ODBC_ENABLED = True Then
		      
		      If (_dbInfo.FullName = "ODBCDatabase") Then
		        
		        // Determine what type of field this is.
		        Dim __parameterInfo As Xojo.Introspection.TypeInfo
		        __parameterInfo = Xojo.Introspection.GetType(Parameters(_count))
		        
		        If (__parameterInfo.FullName = "Date") Then
		          _ps.BindType(_count, ODBCPreparedStatement.ODBC_TYPE_DATE)
		        ElseIf (__parameterInfo.FullName = "Double") Then
		          _ps.BindType(_count, ODBCPreparedStatement.ODBC_TYPE_DOUBLE)
		        ElseIf (__parameterInfo.FullName = "Int64") Then
		          _ps.BindType(_count, ODBCPreparedStatement.ODBC_TYPE_BIGINT)
		        ElseIf (__parameterInfo.FullName = "Int32") Then
		          _ps.BindType(_count, ODBCPreparedStatement.ODBC_TYPE_INTEGER)
		        ElseIf (__parameterInfo.FullName = "Text") Then
		          _ps.BindType(_count, ODBCPreparedStatement.ODBC_TYPE_STRING)
		        End If
		        
		      End If
		      
		    #EndIf
		    
		    // MSSQLServerDatabase (Only on Windows)
		    #If TargetWindows = True And SQLDeLite.PLUGIN_MSSQL_ENABLED = True Then
		      
		      If (_dbInfo.FullName = "MSSQLServerDatabase") Then
		        
		        // Determine what type of field this is.
		        Dim __parameterInfo As Xojo.Introspection.TypeInfo
		        __parameterInfo = Xojo.Introspection.GetType(Parameters(_count))
		        
		        If (__parameterInfo.FullName = "Date") Then
		          _ps.BindType(_count, MSSQLServerPreparedStatement.MSSQLSERVER_TYPE_DATE)
		        ElseIf (__parameterInfo.FullName = "Double") Then
		          _ps.BindType(_count, MSSQLServerPreparedStatement.MSSQLSERVER_TYPE_DOUBLE)
		        ElseIf (__parameterInfo.FullName = "Int64") Then
		          _ps.BindType(_count, MSSQLServerPreparedStatement.MSSQLSERVER_TYPE_BIGINT)
		        ElseIf (__parameterInfo.FullName = "Int32") Then
		          _ps.BindType(_count, MSSQLServerPreparedStatement.MSSQLSERVER_TYPE_INT)
		        ElseIf (__parameterInfo.FullName = "Text") Then
		          _ps.BindType(_count, MSSQLServerPreparedStatement.MSSQLSERVER_TYPE_STRING)
		        End If
		        
		      End If
		      
		    #EndIf
		    
		    // OracleDatabase
		    #If SQLDeLite.PLUGIN_ORACLE_ENABLED = True Then
		      
		      If (_dbInfo.FullName = "OracleDatabase") Then
		        
		        // Determine what type of field this is.
		        Dim __parameterInfo As Xojo.Introspection.TypeInfo
		        __parameterInfo = Xojo.Introspection.GetType(Parameters(_count))
		        
		        If (__parameterInfo.FullName = "Date") Then
		          _ps.BindType(_count, OracleSQLPreparedStatement.SQL_TYPE_DATE)
		        ElseIf (__parameterInfo.FullName = "Double") Then
		          _ps.BindType(_count, OracleSQLPreparedStatement.SQL_TYPE_FLOAT)
		        ElseIf (__parameterInfo.FullName = "Int64") Then
		          _ps.BindType(_count, OracleSQLPreparedStatement.SQL_TYPE_INTEGER)
		        ElseIf (__parameterInfo.FullName = "Int32") Then
		          _ps.BindType(_count, OracleSQLPreparedStatement.SQL_TYPE_INTEGER)
		        ElseIf (__parameterInfo.FullName = "Text") Then
		          _ps.BindType(_count, OracleSQLPreparedStatement.SQL_TYPE_STRING)
		        End If
		        
		      End If
		      
		    #EndIf
		    
		  Next
		  
		  // Loop through the parameters and bind them to the prepared statement. Not applicable to CubeSQLServer or PostgreSQLDatabase.
		  If (_dbInfo.FullName <> "CubeSQLServer" And _dbInfo.FullName <> "PostgreSQLDatabase") Then
		    
		    For _count As Integer = 0 To Parameters.Ubound
		      
		      _ps.Bind(_count, Parameters(_count))
		      
		    Next
		    
		  End If
		  
		  // Determine we are selecting or just executing.
		  If (ExecuteOnly = True) Then
		    
		    // Call the database SQLSelect method now that we have bound all the parameters.
		    If (_dbInfo.FullName = "CubeSQLServer") Then
		      #If SQLdeLite.PLUGIN_CUBESQL_ENABLED = True Then
		        _psCube.VMExecute()
		      #Else
		        _ps.SQLExecute()
		      #EndIf
		    Else
		      _ps.SQLExecute()
		    End If
		    
		  Else
		    
		    // Call the database SQLSelect method now that we have bound all the parameters.
		    If (_dbInfo.FullName = "CubeSQLServer") Then
		      #If SQLdeLite.PLUGIN_CUBESQL_ENABLED = True Then
		        Return _psCube.VMSelect()
		      #Else
		        Return _ps.SQLSelect()
		      #EndIf
		    Else
		      Return _ps.SQLSelect()
		    End If
		    
		  End If
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ParameterizeSQL(db As Object, ByRef SQLText As Text, Record As SQLdeLite.Record) As Auto()
		  // Create array of bound parameters
		  Dim _parameters() As Auto
		  
		  // Create a new SelectText variable
		  Dim _selectText As Text
		  _selectText = SQLText
		  
		  // Determine what database engine we are on.
		  Dim _dbInfo As Xojo.Introspection.TypeInfo
		  _dbInfo = Xojo.Introspection.GetType(db)
		  
		  // Create replacement variable. Some databases use a different variable.
		  Dim _replacement As Text = "?"
		  Dim _replacementCount As Integer = 1
		  
		  If (_dbInfo.FullName = "PostgreSQLDatabase") Then
		    _replacement = "$"
		  End If
		  If (_dbInfo.FullName = "OracleDatabase") Then
		    _replacement = ":"
		  End If
		  If (_dbInfo.FullName = "VDatabase") Then
		    _replacement = ":"
		  End If
		  
		  // Loop through the properties of Record
		  For Each _entry As Xojo.Core.DictionaryEntry In Record.GetIterator()
		    
		    // Create variable representing the variable that would be found in the SQL statement.
		    Dim __field As Text
		    __field = "$" + _entry.Key
		    
		    // Replace the variable with a question mark (for prepared statements) and add the value to the collection of bound parameters.
		    If (_selectText.IndexOf(__field) > 0) Then
		      
		      // PostgreSQLDatabase requires a parameter number (1-based)
		      If (_dbInfo.FullName = "PostgreSQLDatabase") Then
		        _replacement = "$" + _replacementCount.ToText()
		        _replacementCount = _replacementCount + 1
		      End If
		      
		      // OracleDatabase requires a parameter name behind a colon (:name)
		      If (_dbInfo.FullName = "OracleDatabase") Then
		        Dim __key As Text = _entry.Key
		        _replacement = ":" + __key
		      End If
		      
		      // VDatabase (Valentina) requires a parameter number (1-based)
		      If (_dbInfo.FullName = "VDatabase") Then
		        _replacement = ":" + _replacementCount.ToText()
		        _replacementCount = _replacementCount + 1
		      End If
		      
		      _selectText = _selectText.ReplaceAll(__field, _replacement)
		      
		      _parameters.Append(_entry.Value)
		      
		    End If
		    
		  Next
		  
		  // Loop through the public properties of the Record object (potential sub-class) to bind any properties.
		  Dim _recordInfo As Xojo.Introspection.TypeInfo
		  _recordInfo = Xojo.Introspection.GetType(Record)
		  
		  For Each _property As Xojo.Introspection.PropertyInfo In _recordInfo.Properties
		    
		    // Determine if the property is public.
		    If (_property.IsPublic = True) Then
		      
		      // Create variable representing the variable that would be found in the SQL statement.
		      Dim __field As Text
		      __field = "$" + _property.Name
		      
		      // Replace the variable with a question mark (for prepared statements) and add the value to the collection of bound parameters.
		      If (_selectText.IndexOf(__field) > 0) Then
		        
		        // PostgreSQLDatabase requires a parameter number (1-based)
		        If (_dbInfo.FullName = "PostgreSQLDatabase") Then
		          _replacement = "$" + _replacementCount.ToText()
		          _replacementCount = _replacementCount + 1
		        End If
		        
		        // OracleDatabase requires a parameter name behind a colon (:name)
		        If (_dbInfo.FullName = "OracleDatabase") Then
		          _replacement = ":" + _property.Name
		        End If
		        
		        // VDatabase (Valentina) requires a parameter number (1-based)
		        If (_dbInfo.FullName = "VDatabase") Then
		          _replacement = ":" + _replacementCount.ToText()
		          _replacementCount = _replacementCount + 1
		        End If
		        
		        _selectText = _selectText.ReplaceAll(__field, _replacement)
		        
		        _parameters.Append(_property.Value(Record))
		        
		      End If
		      
		    End If
		    
		  Next
		  
		  // Update the SQLText property
		  SQLText = _selectText
		  
		  // Return the parameter array
		  Return _parameters
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, CompatibilityFlags = (TargetConsole and (Target32Bit or Target64Bit)) or  (TargetWeb and (Target32Bit or Target64Bit)) or  (TargetDesktop and (Target32Bit or Target64Bit))
		Function ParameterizeSQL_Variant(db As Object, ByRef SQLText As Text, Record As SQLdeLite.Record) As Variant()
		  // Create array of bound parameters
		  Dim _parameters() As Variant
		  Dim _parametersAuto() As Auto
		  
		  // Call the ParameterizeSQL Auto version
		  _parametersAuto = ParameterizeSQL(db, SQLText, Record)
		  
		  // Convert Auto to Variant
		  For Each _parameter As Auto In _parametersAuto
		    Dim _parameterVariant As Variant
		    _parameterVariant = _parameter
		    _parameters.Append(_parameterVariant)
		  Next
		  
		  // Return the parameter array
		  Return _parameters
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SQLdeLiteExecute(Extends db As Database, SQLText As Text, Record As SQLdeLite.Record)
		  // Verify that the Record object is not Nil. If it's Nil then just call the underlying database SQLSelect method as there is nothing to process.
		  If (Record = Nil) Then
		    db.SQLExecute(SQLText)
		  End If
		  
		  // Create array of bound parameters
		  Dim _parameters() As Auto
		  
		  // Create the parameters array
		  _parameters = ParameterizeSQL(db, SQLText, Record)
		  
		  // Determine if we had any parameters subbed in. If not then call the underlying database SQLExecute method as there is nothing to process.
		  If (_parameters.Ubound = -1) Then
		    db.SQLExecute(SQLText)
		  Else
		    
		    // Since we now have an array of all our parameters let's call the unified mQuery method.
		    Call mQuery(db, SQLText, _parameters, True)
		    
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function SQLdeLiteSelect(Extends db As Database, SQLText As Text, Record As SQLdeLite.Record, FillValuesWhenSingleResult As Boolean = False) As RecordSet
		  // Verify that the Record object is not Nil. If it's Nil then just call the underlying database SQLSelect method as there is nothing to process.
		  If (Record = Nil) Then
		    Return db.SQLSelect(SQLText)
		  End If
		  
		  // Create array of bound parameters
		  Dim _parameters() As Auto
		  
		  // Create the parameters array
		  _parameters = ParameterizeSQL(db, SQLText, Record)
		  
		  // Create RecordSet object to hold result of our SQLdeLite query
		  Dim _rs As RecordSet
		  
		  // Determine if we had any parameters subbed in. If not then call the underlying database SQLSelect method as there is nothing to process.
		  If (_parameters.Ubound = -1) Then
		    _rs = db.SQLSelect(SQLText)
		  Else
		    
		    // Since we now have an array of all our parameters let's call the unified mQuery method.
		    _rs = mQuery(db, SQLText, _parameters)
		    
		  End If
		  
		  // Make sure the user has not overrided FillValuesWhenSingleResult which determines if the RecordSet should write back the values when there is a single result.
		  If (FillValuesWhenSingleResult = True And _rs <> Nil) Then
		    
		    // If the RecordSet only has one row then we should assign the column values back to the original Record object.
		    If (_rs.RecordCount = 1) Then
		      
		      // Loop through the fields in the RecordSet (1-based)
		      For  __count As Integer = 1 To _rs.FieldCount
		        
		        // Get DatabaseField object for this field.
		        Dim __field As DatabaseField
		        __field = _rs.IdxField(__count)
		        
		        // Update the value of the Record object
		        Record.SetProperty(DefineEncoding(__field.Name, Encodings.UTF8).ToText(), __field.Value)
		        
		      Next
		      
		    End If
		    
		  End If
		  
		  // Return the RecordSet
		  Return _rs
		End Function
	#tag EndMethod


	#tag Note, Name = Version
		
		SQLdeLite Version 2.1609.90
		======================
		
		
	#tag EndNote


	#tag Constant, Name = PLUGIN_CUBESQL_ENABLED, Type = Boolean, Dynamic = False, Default = \"False", Scope = Public
	#tag EndConstant

	#tag Constant, Name = PLUGIN_MSSQL_ENABLED, Type = Boolean, Dynamic = False, Default = \"False", Scope = Public
	#tag EndConstant

	#tag Constant, Name = PLUGIN_MYSQL_ENABLED, Type = Boolean, Dynamic = False, Default = \"True", Scope = Public
	#tag EndConstant

	#tag Constant, Name = PLUGIN_ODBC_ENABLED, Type = Boolean, Dynamic = False, Default = \"False", Scope = Public
	#tag EndConstant

	#tag Constant, Name = PLUGIN_ORACLE_ENABLED, Type = Boolean, Dynamic = False, Default = \"False", Scope = Public
	#tag EndConstant

	#tag Constant, Name = PLUGIN_POSTGRESQL_ENABLED, Type = Boolean, Dynamic = False, Default = \"True", Scope = Public
	#tag EndConstant


	#tag ViewBehavior
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
	#tag EndViewBehavior
End Module
#tag EndModule
