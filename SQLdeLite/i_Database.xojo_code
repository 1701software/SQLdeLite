#tag Interface
Protected Interface i_Database
	#tag Method, Flags = &h0
		Sub m_buildAndUpdateObject(ClassName As String, TableName As String, Query As String, Values() As Parameter, PropValues As Dictionary)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub m_cacheTable(tableCache As c_TableCache)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function m_convertVariantsToParams(Values() As Variant) As Parameter()
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function m_getCachedTable(TableName As String) As c_TableCache
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function m_hasCachedTableSchema(className As String) As Boolean
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub m_insertRow(ClassName As String, TableName As String, Values As Dictionary)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub m_verifyTableSchema(className As String)
		  
		End Sub
	#tag EndMethod


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
End Interface
#tag EndInterface
