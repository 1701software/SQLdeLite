#tag Class
Protected Class DatabaseCore
	#tag Method, Flags = &h0
		Sub Close()
		  DoClose()
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Commit()
		  DoCommit()
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Connect() As Boolean
		  Return DoConnect()
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub Constructor()
		  // Verify proper license exists.
		  #IF DebugBuild = FALSE THEN
		    MsgBox("SQLdeLite is not licensed for production yet.")
		    Quit
		  #ENDIF
		  
		  // Initialize variables
		  p_zTableCache = New Dictionary
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub CreateTable(TableName As String, PrimaryKeyColumnName As String, BindType As Integer = - 1)
		  DoCreateTable(TableName, PrimaryKeyColumnName, BindType)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub CreateTableColumn(TableName As String, ColumnName As String, BindType As Integer = - 1)
		  OnCreateTableColumn(TableName, ColumnName, BindType)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub InsertRecord(TableName As String, Data As DatabaseRecord)
		  DoInsertRecord(TableName, Data)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub m_cacheTable(tableCache As c_TableCache)
		  p_zTableCache.Value(tableCache.ClassName) = tableCache
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function m_convertVariantsToParams(Values() As Variant) As Parameter()
		  Return DoVariantsToParameters(Values)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function m_getCachedTable(TableName As String) As c_TableCache
		  // Find cached table information
		  Dim cachedTable As c_TableCache
		  For X As Integer = 0 To p_zTableCache.Keys.Ubound
		    Dim tempTableClass As String
		    tempTableClass = p_zTableCache.Key(X).StringValue
		    If (c_TableCache(p_zTableCache.Value(tempTableClass)).TableName = TableName) Then
		      cachedTable = p_zTableCache.Value(tempTableClass)
		    End If
		  Next
		  Return cachedTable
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function m_hasCachedTableSchema(className As String) As Boolean
		  // Determine if this database has already cached the information for the provided table.
		  #pragma BreakOnExceptions Off
		  Dim tableCache As c_TableCache
		  Try
		    tableCache = p_zTableCache.Value(className)
		    Return True
		  Catch oob As KeyNotFoundException
		    Return False
		  End Try
		  #pragma BreakOnExceptions Default
		  
		  // If we got this far then the table has not been cached.
		  Return False
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Prepare(Statement As String) As PreparedSQLStatement
		  Return DoPrepare(Statement)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Rollback()
		  DoRollback()
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SQLdeLiteExecute(SQLQuery As String, Values() As Parameter)
		  Dim emptyRecordSet As RecordSet
		  emptyRecordSet = DoSQLdeLiteParams(SQLQuery, Values, False)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SQLdeLiteExecute(SQLQuery As String, ParamArray Values As Variant)
		  // Create new Variant array
		  Dim variantArray() As Variant
		  For Each Value As Variant In Values
		    variantArray.Append(Value)
		  Next
		  
		  Dim emptyRecordSet As RecordSet
		  emptyRecordSet = DoSQLdeLiteVariants(SQLQuery, variantArray, False)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function SQLdeLiteSelect(SQLQuery As String, Values() As Parameter) As RecordSet
		  Return DoSQLdeLiteParams(SQLQuery, Values, True)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function SQLdeLiteSelect(SQLQuery As String, ParamArray Values As Variant) As RecordSet
		  // Create new Variant array
		  Dim variantArray() As Variant
		  For Each Value As Variant In Values
		    variantArray.Append(Value)
		  Next
		  
		  Return DoSQLdeLiteVariants(SQLQuery, variantArray, True)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SQLExecute(SQL As String)
		  OnSQLExecute(SQL)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function SQLSelect(SQL As String) As RecordSet
		  Return DoSQLSelect(SQL)
		End Function
	#tag EndMethod


	#tag Hook, Flags = &h0
		Event DoClose()
	#tag EndHook

	#tag Hook, Flags = &h0
		Event DoCommit()
	#tag EndHook

	#tag Hook, Flags = &h0
		Event DoConnect() As Boolean
	#tag EndHook

	#tag Hook, Flags = &h0
		Event DoCreateTable(TableName As String, PrimaryKeyColumnName As String, BindType As Integer = -1)
	#tag EndHook

	#tag Hook, Flags = &h0
		Event DoInsertRecord(TableName As String, Data As DatabaseRecord)
	#tag EndHook

	#tag Hook, Flags = &h0
		Event DoPrepare(Statement As String) As PreparedSQLStatement
	#tag EndHook

	#tag Hook, Flags = &h0
		Event DoRollback()
	#tag EndHook

	#tag Hook, Flags = &h0
		Event DoSQLdeLiteParams(Query As String, Values() As Parameter, ReturnRecordSet As Boolean = False) As RecordSet
	#tag EndHook

	#tag Hook, Flags = &h0
		Event DoSQLdeLiteVariants(Query As String, Values() As Variant, ReturnRecordSet As Boolean = False) As RecordSet
	#tag EndHook

	#tag Hook, Flags = &h0
		Event DoSQLSelect(SQL As String) As RecordSet
	#tag EndHook

	#tag Hook, Flags = &h0
		Event DoVariantsToParameters(Values() As Variant) As Parameter()
	#tag EndHook

	#tag Hook, Flags = &h0
		Event OnCreateTableColumn(TableName As String, ColumnName As String, BindType As Integer = -1)
	#tag EndHook

	#tag Hook, Flags = &h0
		Event OnSQLExecute(SQL As String)
	#tag EndHook


	#tag Property, Flags = &h0
		Error As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		ErrorCode As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		ErrorMessage As String
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected p_isConnected As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected p_zTableCache As Dictionary
	#tag EndProperty


	#tag ViewBehavior
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
End Class
#tag EndClass
