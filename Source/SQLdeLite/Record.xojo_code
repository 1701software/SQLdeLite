#tag Class
Protected Class Record
	#tag Method, Flags = &h0
		Sub Constructor()
		  // Initialize pDictionary_Properties
		  pDictionary_Properties = New Xojo.Core.Dictionary()
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CreateInsertStatement(db As Object, TableName As Text, TableAndFieldNamesQuoted As Boolean = True) As Text
		  // Determine what database engine we are on.
		  Dim _info As Xojo.Introspection.TypeInfo
		  _info = Xojo.Introspection.GetType(me)
		  
		  // Create array to hold INSERT statement
		  Dim _sql() As Text
		  
		  _sql.Append("INSERT INTO ")
		  If (TableAndFieldNamesQuoted = True) Then
		    _sql.Append("""")
		  End If
		  _sql.Append(TableName)
		  If (TableAndFieldNamesQuoted = True) Then
		    _sql.Append("""")
		  End If
		  _sql.Append(" (")
		  
		  // Loop through the properties of Record
		  For Each _entry As Xojo.Core.DictionaryEntry In Record.GetIterator()
		    
		    If (TableAndFieldNamesQuoted = True) Then
		      _sql.Append("""")
		    End If
		    _sql.Append(_entry.Key)
		    If (TableAndFieldNamesQuoted = True) Then
		      _sql.Append("""")
		    End If
		    _sql.Append(", ")
		    
		  Next
		  
		  // Remove trailing comma.
		  _sql.Remove(_sql.Ubound)
		  
		  // Write out the values.
		  _sql.Append(") VALUES (")
		  
		  // Loop through the properties of Record
		  For Each _entry As Xojo.Core.DictionaryEntry In Record.GetIterator()
		    
		    // Verify the entry has a value.
		    If (_entry.Value = Nil) Then
		      
		      _sql.Append("NULL")
		      
		    Else
		      
		      // Find the type of the value
		      Dim __entryInfo As Xojo.Introspection.TypeInfo
		      __entryInfo = Xojo.Introspection.GetType(_entry.Value)
		      
		      If (__entryInfo.FullName = "Int32") Then
		        Dim __temp As Int32
		        __temp = _entry.Value
		        _sql.Append(__temp.ToText())
		      ElseIf (__entryInfo.FullName = "Int64") Then
		        Dim __temp As Int64
		        __temp = _entry.Value
		        _sql.Append(__temp.ToText())
		      ElseIf (__entryInfo.FullName = "Integer") Then
		        Dim __temp As Integer
		        __temp = _entry.Value
		        _sql.Append(__temp.ToText())
		      ElseIf (__entryInfo.FullName = "Double") Then
		        Dim __temp As Double
		        __temp = _entry.Value
		        _sql.Append(__temp.ToText())
		      ElseIf (__entryInfo.FullName = "String") Then
		        #If TargetIOS = False Then
		          Dim __temp As String
		          __temp = _entry.Value
		          _sql.Append("'")
		          _sql.Append(DefineEncoding(__temp, Encodings.UTF8).ToText())
		          _sql.Append("'")
		        #EndIf
		      ElseIf (__entryInfo.FullName = "Text") Then
		        Dim __temp As Text
		        __temp = _entry.Value
		        _sql.Append("'")
		        _sql.Append(__temp)
		        _sql.Append("'")
		      End If
		      
		    End If
		    
		    _sql.Append(", ")
		    
		  Next
		  
		  // Loop through the public properties of the Record object (potential sub-class) to bind any properties.
		  Dim _recordInfo As Xojo.Introspection.TypeInfo
		  _recordInfo = Xojo.Introspection.GetType(me)
		  
		  For Each _property As Xojo.Introspection.PropertyInfo In _recordInfo.Properties
		    
		    // Determine if the property is public.
		    If (_property.IsPublic = True) Then
		      
		      // Find the type of the value
		      Dim __entryInfo As Xojo.Introspection.TypeInfo
		      __entryInfo = Xojo.Introspection.GetType(_property)
		      
		      If (__entryInfo.FullName = "Int32") Then
		        Dim __temp As Int32
		        __temp = _property.Value(me)
		        _sql.Append(__temp.ToText())
		      ElseIf (__entryInfo.FullName = "Int64") Then
		        Dim __temp As Int64
		        __temp = _property.Value(me)
		        _sql.Append(__temp.ToText())
		      ElseIf (__entryInfo.FullName = "Integer") Then
		        Dim __temp As Integer
		        __temp = _property.Value(me)
		        _sql.Append(__temp.ToText())
		      ElseIf (__entryInfo.FullName = "Double") Then
		        Dim __temp As Double
		        __temp = _property.Value(me)
		        _sql.Append(__temp.ToText())
		      ElseIf (__entryInfo.FullName = "String") Then
		        #If TargetIOS = False Then
		          Dim __temp As String
		          __temp = _property.Value(me)
		          _sql.Append(DefineEncoding(__temp, Encodings.UTF8).ToText().ReplaceAll("'", "''"))
		        #EndIf
		      ElseIf (__entryInfo.FullName = "Text") Then
		        Dim __temp As Text
		        __temp = _property.Value(me)
		        _sql.Append(__temp.ReplaceAll("'", "''"))
		      End If
		      _sql.Append(", ")
		      
		    End If
		    
		  Next
		  
		  // Remove trailing comma.
		  _sql.Remove(_sql.Ubound)
		  
		  // Close the INSERT statement
		  _sql.Append(")")
		  
		  // Return the INSERT statement
		  Return Text.Join(_sql, "")
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetIterator() As Xojo.Core.Dictionary
		  Return pDictionary_Properties
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetProperty(Name As Text) As Auto
		  Return pDictionary_Properties.Value(Name)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Operator_Lookup(Name As Text) As Auto
		  If (pDictionary_Properties.HasKey(Name)) Then
		    
		    Return pDictionary_Properties.Value(Name)
		    
		  End If
		  
		  Return Nil
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Operator_Lookup(Name As Text, Assigns Value As Boolean)
		  pDictionary_Properties.Value(Name) = Value
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Operator_Lookup(Name As Text, Assigns Value As Date)
		  pDictionary_Properties.Value(Name) = Value
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Operator_Lookup(Name As Text, Assigns Value As Double)
		  pDictionary_Properties.Value(Name) = Value
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Operator_Lookup(Name As Text, Assigns Value As Int64)
		  pDictionary_Properties.Value(Name) = Value
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Operator_Lookup(Name As Text, Assigns Value As Integer)
		  pDictionary_Properties.Value(Name) = Value
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0, CompatibilityFlags = (TargetConsole and (Target32Bit or Target64Bit)) or  (TargetWeb and (Target32Bit or Target64Bit)) or  (TargetDesktop and (Target32Bit or Target64Bit))
		Sub Operator_Lookup(Name As Text, Assigns Value As String)
		  pDictionary_Properties.Value(Name) = Value
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Operator_Lookup(Name As Text, Assigns Value As Text)
		  pDictionary_Properties.Value(Name) = Value
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Operator_Lookup(Name As Text, Assigns Value As UInt64)
		  pDictionary_Properties.Value(Name) = Value
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetProperty(Name As Text, Value As Auto)
		  pDictionary_Properties.Value(Name) = Value
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h21
		Private pDictionary_Properties As Xojo.Core.Dictionary
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
