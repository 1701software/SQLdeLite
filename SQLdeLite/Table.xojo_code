#tag Class
Protected Class Table
	#tag Method, Flags = &h0
		Sub Constructor(tableDatabase As i_Database, tableName As String = "")
		  // Initialize p_tableDatabase
		  p_tableDatabase = tableDatabase
		  p_stagingProperties = New Dictionary()
		  
		  // Determine the current className
		  Dim tempTable As Introspection.TypeInfo
		  tempTable = Introspection.GetType(self)
		  p_className = tempTable.FullName
		  
		  // Determine if the user has passed in a tableName to override introspection discovery.
		  If (tableName = "") Then
		    // Determine what the tableName is automatically using introspection result.
		    Dim tempTableName As String
		    tempTableName = p_className
		    //Dim tableNameLength As Integer = 0
		    Dim X As Integer = 0
		    While X < Len(p_className)
		      // Capture the right side of the class name with this length
		      Dim tempTempTableName As String
		      tempTempTableName = Right(p_className, X)
		      // Determine if there is a period in it, otherwise just keep moving to the beginning.
		      // If there's a period then we want to use the final class name in the hierarchy as the table name.
		      // If there's no period in it then it will just end up being the class name since its not in a module.
		      If (InStr(0, tempTempTableName, ".") > 0) Then
		        // Found the last word in class heirarchy.
		        tempTableName = Right(tempTableName, X)
		        X = Len(p_className) // End the While since we found the class name in the heirarchy
		        // If this was in a hierarchy then it probably has a leading period (the one we found), lets remove that.
		        If (Left(tempTableName, 1) = ".") Then
		          tempTableName = Replace(tempTableName, ".", "")
		        End If
		        tableName = tempTableName
		      Else
		        // Need to dig deeper
		        X = X + 1
		      End If
		    Wend
		  End If
		  
		  // Set local version of tableName
		  p_tableName = tableName
		  
		  // Ask our desired database if it has cached the results of the current table schema.
		  If (tableDatabase.m_HasCachedTableSchema(p_className) = False) Then
		    // We need to add this table's schema to the database cache.
		    Dim tableCache As New c_TableCache
		    tableCache.ClassName = p_className
		    tableCache.TableName = tableName
		    // Loop through properties to add to tableCache
		    Dim i As Introspection.TypeInfo
		    i = Introspection.GetType(self)
		    Dim p() As Introspection.PropertyInfo = i.GetProperties()
		    For x As Integer = 0 to p.Ubound
		      // Make sure we only cache the public properties. This filters out private framework properties
		      // and properties that are just used for methods, etc.
		      If (p(x).IsPrivate = False) Then
		        // Cache property values
		        Dim propCache As New c_PropertyCache
		        propCache.PropertyName = p(x).Name
		        propCache.PropertyType = p(x).PropertyType.FullName
		        // Initialize XojoScript object for querying property values at run-time.
		        propCache.PropertyScript = New XojoScript
		        // Add property information to PropertyScriptSource
		        Dim tempScript() As String
		        tempScript.Append("Sub SetValue (propertyName As String)")
		        tempScript.Append(EndOfLine)
		        tempScript.Append("m_setPropertyValue(propertyName")
		        // Let's determine if this is a Date or other complex object (XojoScript cannot pass objects so we need to compensate)
		        If (p(x).PropertyType.Name = "Double" Or p(x).PropertyType.Name = "Int32" Or p(x).PropertyType.Name = "String") Then
		          // This is a string or integer
		          tempScript.Append(", ")
		          tempScript.Append(p(x).Name)
		        Else // This is a complex object, do nothing and let m_setPropertyValue figure out what event to raise based on type.
		        End If
		        tempScript.Append(")")
		        tempscript.Append(EndOfLine)
		        tempScript.Append("End Sub")
		        tempScript.Append(EndOfLine)
		        tempScript.Append("SetValue(""")
		        tempScript.Append(p(x).Name)
		        tempScript.Append(""")")
		        propCache.PropertyScript.Source = Join(tempScript, "") + EndOfLine
		        tableCache.Properties.Append(propCache)
		      End If
		    Next
		    // Register completed tableCache to database.
		    p_tableDatabase.m_CacheTable(tableCache)
		  End If
		  
		  // Let's verify the schema against the database of choice.
		  p_tableDatabase.m_verifyTableSchema(p_ClassName)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Execute(Query As String, ParamArray Values As Parameter)
		  // Update cache of current values
		  m_updateAllValues()
		  
		  // We just need to pass along the appropriate information to the database object.
		  p_tableDatabase.m_buildAndUpdateObject(p_className, p_tableName, Query, Values, p_stagingProperties)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Execute(Query As String, ParamArray Values As Variant)
		  // Update cache of current values
		  m_updateAllValues()
		  
		  // We need to convert the values to useful types.
		  Dim props() As Parameter
		  Dim variantArray() As Variant
		  For Each Value As Variant In Values
		    variantArray.Append(Value)
		  Next
		  props = p_tableDatabase.m_convertVariantsToParams(variantArray)
		  
		  // We just need to pass along the appropriate information to the database object.
		  p_tableDatabase.m_buildAndUpdateObject(p_className, p_tableName, Query, props, p_stagingProperties)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Insert()
		  // Update cache of current values
		  m_updateAllValues()
		  
		  // Tell database to insert a row with these values in dictionary
		  p_tableDatabase.m_insertRow(p_className, p_tableName, p_stagingProperties)
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub m_setPropertyValue(PropertyName As String)
		  // Acquire existing TableCache held by the database.
		  Dim tableCache As c_TableCache
		  tableCache = p_tableDatabase.m_GetCachedTable(p_tableName)
		  
		  // Loop through properties setting the stagingProperties Dictionary
		  For Each p As c_PropertyCache In tableCache.Properties
		    // Determine if this property is the one we are trying to set.
		    If (p.PropertyName = PropertyName) Then
		      
		      // Determine if the property is a Date object
		      If (p.PropertyType = "Date") Then
		        // Set the staging property value for this property.
		        Dim tempDateString As String
		        tempDateString = SetDateString(PropertyName)
		        If (tempDateString <> "") Then
		          p_lastPropertyValue = tempDateString
		        End If
		      End If
		      
		    End If
		  Next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub m_setPropertyValue(PropertyName As String, Value As Double)
		  p_lastPropertyValue = Value
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub m_setPropertyValue(PropertyName As String, Value As Integer)
		  p_lastPropertyValue = Value
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub m_setPropertyValue(PropertyName As String, Value As String)
		  p_lastPropertyValue = Value
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub m_updateAllValues()
		  // Acquire existing TableCache held by the database.
		  Dim tableCache As c_TableCache
		  tableCache = p_tableDatabase.m_GetCachedTable(p_tableName)
		  
		  // Reset the temporary Dictionary we are using to cache property values.
		  p_stagingProperties.Clear()
		  
		  // Loop through properties setting the stagingProperties Dictionary
		  For Each p As c_PropertyCache In tableCache.Properties
		    p.PropertyScript.Context = self
		    p.PropertyScript.Run
		    p.PropertyScript.Context = Nil
		    // We need to make sure that the last property value gets set back to Nil so properties don't overlap.
		    If (p_lastPropertyValue <> Nil) Then
		      p_stagingProperties.Value(p.PropertyName) = p_lastPropertyValue
		      p_lastPropertyValue = Nil
		    End If
		  Next
		End Sub
	#tag EndMethod


	#tag Hook, Flags = &h0
		Event SetDateString(PropertyName As String) As String
	#tag EndHook


	#tag Property, Flags = &h21
		Private p_className As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private p_lastPropertyValue As Variant
	#tag EndProperty

	#tag Property, Flags = &h21
		Private p_stagingProperties As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h21
		Private p_tableDatabase As i_Database
	#tag EndProperty

	#tag Property, Flags = &h21
		Private p_tableName As String
	#tag EndProperty


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
End Class
#tag EndClass
