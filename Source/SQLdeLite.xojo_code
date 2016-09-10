#tag Module
Protected Module SQLdeLite
	#tag Method, Flags = &h21, CompatibilityFlags = (TargetConsole and (Target32Bit or Target64Bit)) or  (TargetWeb and (Target32Bit or Target64Bit)) or  (TargetDesktop and (Target32Bit or Target64Bit))
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

	#tag Method, Flags = &h21, CompatibilityFlags = (TargetIOS and (Target32Bit or Target64Bit))
		Private Function mQuery(db As iOSSQLiteDatabase, SelectText As Text, Parameters() As Auto, ExecuteOnly As Boolean = False) As iOSSQLiteRecordSet
		  // Determine if we had any parameters subbed in. If not then call the underlying database SQLSelect method as there is nothing to process.
		  If (Parameters.Ubound = -1) Then
		    Return db.SQLSelect(SelectText)
		  End If
		  
		  // Determine we are selecting or just executing.
		  If (ExecuteOnly = True) Then
		    
		    // Call the database SQLSelect method now that we have bound all the parameters.
		    db.SQLExecute(SelectText, Parameters)
		    
		  Else
		    
		    // Call the database SQLSelect method now that we have bound all the parameters. On iOS you cannot pass an array of parameters so we manually call the function depending on the number of parameters.
		    If (Parameters.Ubound = 0) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0))
		      
		    ElseIf (Parameters.Ubound = 1) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1))
		      
		    ElseIf (Parameters.Ubound = 2) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2))
		      
		    ElseIf (Parameters.Ubound = 3) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3))
		      
		    ElseIf (Parameters.Ubound = 4) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4))
		      
		    ElseIf (Parameters.Ubound = 5) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5))
		      
		    ElseIf (Parameters.Ubound = 6) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6))
		      
		    ElseIf (Parameters.Ubound = 7) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7))
		      
		    ElseIf (Parameters.Ubound = 8) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8))
		      
		    ElseIf (Parameters.Ubound = 9) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9))
		      
		    ElseIf (Parameters.Ubound = 10) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10))
		      
		    ElseIf (Parameters.Ubound = 11) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11))
		      
		    ElseIf (Parameters.Ubound = 12) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12))
		      
		    ElseIf (Parameters.Ubound = 13) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13))
		      
		    ElseIf (Parameters.Ubound = 14) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14))
		      
		    ElseIf (Parameters.Ubound = 15) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15))
		      
		    ElseIf (Parameters.Ubound = 16) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16))
		      
		    ElseIf (Parameters.Ubound = 17) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17))
		      
		    ElseIf (Parameters.Ubound = 18) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18))
		      
		    ElseIf (Parameters.Ubound = 19) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19))
		      
		    ElseIf (Parameters.Ubound = 20) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20))
		      
		    ElseIf (Parameters.Ubound = 21) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21))
		      
		    ElseIf (Parameters.Ubound = 22) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22))
		      
		    ElseIf (Parameters.Ubound = 23) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23))
		      
		    ElseIf (Parameters.Ubound = 24) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24))
		      
		    ElseIf (Parameters.Ubound = 25) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25))
		      
		    ElseIf (Parameters.Ubound = 26) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26))
		      
		    ElseIf (Parameters.Ubound = 27) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27))
		      
		    ElseIf (Parameters.Ubound = 28) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28))
		      
		    ElseIf (Parameters.Ubound = 29) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29))
		      
		    ElseIf (Parameters.Ubound = 30) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30))
		      
		    ElseIf (Parameters.Ubound = 31) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31))
		      
		    ElseIf (Parameters.Ubound = 32) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32))
		      
		    ElseIf (Parameters.Ubound = 33) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33))
		      
		    ElseIf (Parameters.Ubound = 34) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34))
		      
		    ElseIf (Parameters.Ubound = 35) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35))
		      
		    ElseIf (Parameters.Ubound = 36) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36))
		      
		    ElseIf (Parameters.Ubound = 37) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37))
		      
		    ElseIf (Parameters.Ubound = 38) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38))
		      
		    ElseIf (Parameters.Ubound = 39) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39))
		      
		    ElseIf (Parameters.Ubound = 40) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40))
		      
		    ElseIf (Parameters.Ubound = 41) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41))
		      
		    ElseIf (Parameters.Ubound = 42) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42))
		      
		    ElseIf (Parameters.Ubound = 43) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43))
		      
		    ElseIf (Parameters.Ubound = 44) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44))
		      
		    ElseIf (Parameters.Ubound = 45) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45))
		      
		    ElseIf (Parameters.Ubound = 46) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46))
		      
		    ElseIf (Parameters.Ubound = 47) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47))
		      
		    ElseIf (Parameters.Ubound = 48) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48))
		      
		    ElseIf (Parameters.Ubound = 49) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49))
		      
		    ElseIf (Parameters.Ubound = 50) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50))
		      
		    ElseIf (Parameters.Ubound = 51) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51))
		      
		    ElseIf (Parameters.Ubound = 52) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52))
		      
		    ElseIf (Parameters.Ubound = 53) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53))
		      
		    ElseIf (Parameters.Ubound = 54) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54))
		      
		    ElseIf (Parameters.Ubound = 55) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55))
		      
		    ElseIf (Parameters.Ubound = 56) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56))
		      
		    ElseIf (Parameters.Ubound = 57) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57))
		      
		    ElseIf (Parameters.Ubound = 58) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58))
		      
		    ElseIf (Parameters.Ubound = 59) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59))
		      
		    ElseIf (Parameters.Ubound = 60) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60))
		      
		    ElseIf (Parameters.Ubound = 61) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61))
		      
		    ElseIf (Parameters.Ubound = 62) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62))
		      
		    ElseIf (Parameters.Ubound = 63) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63))
		      
		    ElseIf (Parameters.Ubound = 64) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64))
		      
		    ElseIf (Parameters.Ubound = 65) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65))
		      
		    ElseIf (Parameters.Ubound = 66) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66))
		      
		    ElseIf (Parameters.Ubound = 67) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67))
		      
		    ElseIf (Parameters.Ubound = 68) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68))
		      
		    ElseIf (Parameters.Ubound = 69) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68), Parameters(69))
		      
		    ElseIf (Parameters.Ubound = 70) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68), Parameters(69), Parameters(70))
		      
		    ElseIf (Parameters.Ubound = 71) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68), Parameters(69), Parameters(70), Parameters(71))
		      
		    ElseIf (Parameters.Ubound = 72) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68), Parameters(69), Parameters(70), Parameters(71), Parameters(72))
		      
		    ElseIf (Parameters.Ubound = 73) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68), Parameters(69), Parameters(70), Parameters(71), Parameters(72), Parameters(73))
		      
		    ElseIf (Parameters.Ubound = 74) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68), Parameters(69), Parameters(70), Parameters(71), Parameters(72), Parameters(73), Parameters(74))
		      
		    ElseIf (Parameters.Ubound = 75) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68), Parameters(69), Parameters(70), Parameters(71), Parameters(72), Parameters(73), Parameters(74), Parameters(75))
		      
		    ElseIf (Parameters.Ubound = 76) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68), Parameters(69), Parameters(70), Parameters(71), Parameters(72), Parameters(73), Parameters(74), Parameters(75), Parameters(76))
		      
		    ElseIf (Parameters.Ubound = 77) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68), Parameters(69), Parameters(70), Parameters(71), Parameters(72), Parameters(73), Parameters(74), Parameters(75), Parameters(76), Parameters(77))
		      
		    ElseIf (Parameters.Ubound = 78) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68), Parameters(69), Parameters(70), Parameters(71), Parameters(72), Parameters(73), Parameters(74), Parameters(75), Parameters(76), Parameters(77), Parameters(78))
		      
		    ElseIf (Parameters.Ubound = 79) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68), Parameters(69), Parameters(70), Parameters(71), Parameters(72), Parameters(73), Parameters(74), Parameters(75), Parameters(76), Parameters(77), Parameters(78), Parameters(79))
		      
		    ElseIf (Parameters.Ubound = 80) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68), Parameters(69), Parameters(70), Parameters(71), Parameters(72), Parameters(73), Parameters(74), Parameters(75), Parameters(76), Parameters(77), Parameters(78), Parameters(79), Parameters(80))
		      
		    ElseIf (Parameters.Ubound = 81) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68), Parameters(69), Parameters(70), Parameters(71), Parameters(72), Parameters(73), Parameters(74), Parameters(75), Parameters(76), Parameters(77), Parameters(78), Parameters(79), Parameters(80), Parameters(81))
		      
		    ElseIf (Parameters.Ubound = 82) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68), Parameters(69), Parameters(70), Parameters(71), Parameters(72), Parameters(73), Parameters(74), Parameters(75), Parameters(76), Parameters(77), Parameters(78), Parameters(79), Parameters(80), Parameters(81), Parameters(82))
		      
		    ElseIf (Parameters.Ubound = 83) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68), Parameters(69), Parameters(70), Parameters(71), Parameters(72), Parameters(73), Parameters(74), Parameters(75), Parameters(76), Parameters(77), Parameters(78), Parameters(79), Parameters(80), Parameters(81), Parameters(82), Parameters(83))
		      
		    ElseIf (Parameters.Ubound = 84) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68), Parameters(69), Parameters(70), Parameters(71), Parameters(72), Parameters(73), Parameters(74), Parameters(75), Parameters(76), Parameters(77), Parameters(78), Parameters(79), Parameters(80), Parameters(81), Parameters(82), Parameters(83), Parameters(84))
		      
		    ElseIf (Parameters.Ubound = 85) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68), Parameters(69), Parameters(70), Parameters(71), Parameters(72), Parameters(73), Parameters(74), Parameters(75), Parameters(76), Parameters(77), Parameters(78), Parameters(79), Parameters(80), Parameters(81), Parameters(82), Parameters(83), Parameters(84), Parameters(85))
		      
		    ElseIf (Parameters.Ubound = 86) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68), Parameters(69), Parameters(70), Parameters(71), Parameters(72), Parameters(73), Parameters(74), Parameters(75), Parameters(76), Parameters(77), Parameters(78), Parameters(79), Parameters(80), Parameters(81), Parameters(82), Parameters(83), Parameters(84), Parameters(85), Parameters(86))
		      
		    ElseIf (Parameters.Ubound = 87) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68), Parameters(69), Parameters(70), Parameters(71), Parameters(72), Parameters(73), Parameters(74), Parameters(75), Parameters(76), Parameters(77), Parameters(78), Parameters(79), Parameters(80), Parameters(81), Parameters(82), Parameters(83), Parameters(84), Parameters(85), Parameters(86), Parameters(87))
		      
		    ElseIf (Parameters.Ubound = 88) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68), Parameters(69), Parameters(70), Parameters(71), Parameters(72), Parameters(73), Parameters(74), Parameters(75), Parameters(76), Parameters(77), Parameters(78), Parameters(79), Parameters(80), Parameters(81), Parameters(82), Parameters(83), Parameters(84), Parameters(85), Parameters(86), Parameters(87), Parameters(88))
		      
		    ElseIf (Parameters.Ubound = 89) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68), Parameters(69), Parameters(70), Parameters(71), Parameters(72), Parameters(73), Parameters(74), Parameters(75), Parameters(76), Parameters(77), Parameters(78), Parameters(79), Parameters(80), Parameters(81), Parameters(82), Parameters(83), Parameters(84), Parameters(85), Parameters(86), Parameters(87), Parameters(88), Parameters(89))
		      
		    ElseIf (Parameters.Ubound = 90) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68), Parameters(69), Parameters(70), Parameters(71), Parameters(72), Parameters(73), Parameters(74), Parameters(75), Parameters(76), Parameters(77), Parameters(78), Parameters(79), Parameters(80), Parameters(81), Parameters(82), Parameters(83), Parameters(84), Parameters(85), Parameters(86), Parameters(87), Parameters(88), Parameters(89), Parameters(90))
		      
		    ElseIf (Parameters.Ubound = 91) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68), Parameters(69), Parameters(70), Parameters(71), Parameters(72), Parameters(73), Parameters(74), Parameters(75), Parameters(76), Parameters(77), Parameters(78), Parameters(79), Parameters(80), Parameters(81), Parameters(82), Parameters(83), Parameters(84), Parameters(85), Parameters(86), Parameters(87), Parameters(88), Parameters(89), Parameters(90), Parameters(91))
		      
		    ElseIf (Parameters.Ubound = 92) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68), Parameters(69), Parameters(70), Parameters(71), Parameters(72), Parameters(73), Parameters(74), Parameters(75), Parameters(76), Parameters(77), Parameters(78), Parameters(79), Parameters(80), Parameters(81), Parameters(82), Parameters(83), Parameters(84), Parameters(85), Parameters(86), Parameters(87), Parameters(88), Parameters(89), Parameters(90), Parameters(91), Parameters(92))
		      
		    ElseIf (Parameters.Ubound = 93) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68), Parameters(69), Parameters(70), Parameters(71), Parameters(72), Parameters(73), Parameters(74), Parameters(75), Parameters(76), Parameters(77), Parameters(78), Parameters(79), Parameters(80), Parameters(81), Parameters(82), Parameters(83), Parameters(84), Parameters(85), Parameters(86), Parameters(87), Parameters(88), Parameters(89), Parameters(90), Parameters(91), Parameters(92), Parameters(93))
		      
		    ElseIf (Parameters.Ubound = 94) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68), Parameters(69), Parameters(70), Parameters(71), Parameters(72), Parameters(73), Parameters(74), Parameters(75), Parameters(76), Parameters(77), Parameters(78), Parameters(79), Parameters(80), Parameters(81), Parameters(82), Parameters(83), Parameters(84), Parameters(85), Parameters(86), Parameters(87), Parameters(88), Parameters(89), Parameters(90), Parameters(91), Parameters(92), Parameters(93), Parameters(94))
		      
		    ElseIf (Parameters.Ubound = 95) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68), Parameters(69), Parameters(70), Parameters(71), Parameters(72), Parameters(73), Parameters(74), Parameters(75), Parameters(76), Parameters(77), Parameters(78), Parameters(79), Parameters(80), Parameters(81), Parameters(82), Parameters(83), Parameters(84), Parameters(85), Parameters(86), Parameters(87), Parameters(88), Parameters(89), Parameters(90), Parameters(91), Parameters(92), Parameters(93), Parameters(94), Parameters(95))
		      
		    ElseIf (Parameters.Ubound = 96) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68), Parameters(69), Parameters(70), Parameters(71), Parameters(72), Parameters(73), Parameters(74), Parameters(75), Parameters(76), Parameters(77), Parameters(78), Parameters(79), Parameters(80), Parameters(81), Parameters(82), Parameters(83), Parameters(84), Parameters(85), Parameters(86), Parameters(87), Parameters(88), Parameters(89), Parameters(90), Parameters(91), Parameters(92), Parameters(93), Parameters(94), Parameters(95), Parameters(96))
		      
		    ElseIf (Parameters.Ubound = 97) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68), Parameters(69), Parameters(70), Parameters(71), Parameters(72), Parameters(73), Parameters(74), Parameters(75), Parameters(76), Parameters(77), Parameters(78), Parameters(79), Parameters(80), Parameters(81), Parameters(82), Parameters(83), Parameters(84), Parameters(85), Parameters(86), Parameters(87), Parameters(88), Parameters(89), Parameters(90), Parameters(91), Parameters(92), Parameters(93), Parameters(94), Parameters(95), Parameters(96), Parameters(97))
		      
		    ElseIf (Parameters.Ubound = 98) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68), Parameters(69), Parameters(70), Parameters(71), Parameters(72), Parameters(73), Parameters(74), Parameters(75), Parameters(76), Parameters(77), Parameters(78), Parameters(79), Parameters(80), Parameters(81), Parameters(82), Parameters(83), Parameters(84), Parameters(85), Parameters(86), Parameters(87), Parameters(88), Parameters(89), Parameters(90), Parameters(91), Parameters(92), Parameters(93), Parameters(94), Parameters(95), Parameters(96), Parameters(97), Parameters(98))
		      
		    ElseIf (Parameters.Ubound = 99) Then
		      
		      Return db.SQLSelect(SelectText, Parameters(0), Parameters(1), Parameters(2), Parameters(3), Parameters(4), Parameters(5), Parameters(6), Parameters(7), Parameters(8), Parameters(9), Parameters(10), Parameters(11), Parameters(12), Parameters(13), Parameters(14), Parameters(15), Parameters(16), Parameters(17), Parameters(18), Parameters(19), Parameters(20), Parameters(21), Parameters(22), Parameters(23), Parameters(24), Parameters(25), Parameters(26), Parameters(27), Parameters(28), Parameters(29), Parameters(30), Parameters(31), Parameters(32), Parameters(33), Parameters(34), Parameters(35), Parameters(36), Parameters(37), Parameters(38), Parameters(39), Parameters(40), Parameters(41), Parameters(42), Parameters(43), Parameters(44), Parameters(45), Parameters(46), Parameters(47), Parameters(48), Parameters(49), Parameters(50), Parameters(51), Parameters(52), Parameters(53), Parameters(54), Parameters(55), Parameters(56), Parameters(57), Parameters(58), Parameters(59), Parameters(60), Parameters(61), Parameters(62), Parameters(63), Parameters(64), Parameters(65), Parameters(66), Parameters(67), Parameters(68), Parameters(69), Parameters(70), Parameters(71), Parameters(72), Parameters(73), Parameters(74), Parameters(75), Parameters(76), Parameters(77), Parameters(78), Parameters(79), Parameters(80), Parameters(81), Parameters(82), Parameters(83), Parameters(84), Parameters(85), Parameters(86), Parameters(87), Parameters(88), Parameters(89), Parameters(90), Parameters(91), Parameters(92), Parameters(93), Parameters(94), Parameters(95), Parameters(96), Parameters(97), Parameters(98), Parameters(99))
		      
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
		  
		  #If TargetIOS <> True Then
		    
		    If (_dbInfo.FullName = "PostgreSQLDatabase") Then
		      _replacement = "$"
		    End If
		    If (_dbInfo.FullName = "OracleDatabase") Then
		      _replacement = ":"
		    End If
		    If (_dbInfo.FullName = "VDatabase") Then
		      _replacement = ":"
		    End If
		    
		  #EndIf
		  
		  // Loop through the properties of Record
		  For Each _entry As Xojo.Core.DictionaryEntry In Record.GetIterator()
		    
		    // Create variable representing the variable that would be found in the SQL statement.
		    Dim __field As Text
		    __field = "$" + _entry.Key
		    
		    // Replace the variable with a question mark (for prepared statements) and add the value to the collection of bound parameters.
		    If (_selectText.IndexOf(__field) > 0) Then
		      
		      #If TargetIOS <> True Then
		        
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
		        
		      #EndIf
		      
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
		        
		        #If TargetIOS <> True Then
		          
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
		          
		        #EndIf
		        
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

	#tag Method, Flags = &h0, CompatibilityFlags = (TargetConsole and (Target32Bit or Target64Bit)) or  (TargetWeb and (Target32Bit or Target64Bit)) or  (TargetDesktop and (Target32Bit or Target64Bit))
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

	#tag Method, Flags = &h0, CompatibilityFlags = (TargetIOS and (Target32Bit or Target64Bit))
		Sub SQLdeLiteExecute(Extends db As iOSSQLiteDatabase, SQLText As Text, Record As SQLdeLite.Record)
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

	#tag Method, Flags = &h0, CompatibilityFlags = (TargetConsole and (Target32Bit or Target64Bit)) or  (TargetWeb and (Target32Bit or Target64Bit)) or  (TargetDesktop and (Target32Bit or Target64Bit))
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

	#tag Method, Flags = &h0, CompatibilityFlags = (TargetIOS and (Target32Bit or Target64Bit))
		Function SQLdeLiteSelect(Extends db As iOSSqliteDatabase, SQLText As Text, Record As SQLdeLite.Record, FillValuesWhenSingleResult As Boolean = False) As iOSSQLiteRecordSet
		  // Verify that the Record object is not Nil. If it's Nil then just call the underlying database SQLSelect method as there is nothing to process.
		  If (Record = Nil) Then
		    Return db.SQLSelect(SQLText)
		  End If
		  
		  // Create array of bound parameters
		  Dim _parameters() As Auto
		  
		  // Create the parameters array
		  _parameters = ParameterizeSQL(db, SQLText, Record)
		  
		  // Create RecordSet object to hold result of our SQLdeLite query
		  Dim _rs As iOSSQLiteRecordSet
		  
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
		      
		      // Loop through the fields in the RecordSet (0-based on iOS)
		      For  __count As Integer = 0 To _rs.FieldCount - 1
		        
		        // Get DatabaseField object for this field.
		        Dim __field As iOSSQLiteDatabaseField
		        __field = _rs.IdxField(__count)
		        
		        // Update the value of the Record object
		        Record.SetProperty(__field.Name, __field.Value)
		        
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
