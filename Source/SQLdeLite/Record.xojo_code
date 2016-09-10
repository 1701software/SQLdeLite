#tag Class
Protected Class Record
	#tag Method, Flags = &h0
		Sub Constructor()
		  // Initialize pDictionary_Properties
		  pDictionary_Properties = New Xojo.Core.Dictionary()
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetIterator() As Xojo.Core.Dictionary
		  Return pDictionary_Properties
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

	#tag Method, Flags = &h0
		Sub Operator_Lookup(Name As Text, Assigns Value As Text)
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
