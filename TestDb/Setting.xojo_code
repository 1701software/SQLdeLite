#tag Class
Protected Class Setting
Inherits SQLdeLite.Table
	#tag Event
		Function SetDateString(PropertyName As String) As String
		  // We need to catch the fact that the UpdateDate property is a complex object and can't be handled dynamically.
		  If (PropertyName = "UpdateDate") Then
		    If (UpdateDate <> Nil) Then
		      Return UpdateDate.SQLDateTime
		    Else
		      Return ""
		    End If
		  End If
		End Function
	#tag EndEvent


	#tag Property, Flags = &h0
		DoubleTest As Double
	#tag EndProperty

	#tag Property, Flags = &h0
		ID As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		Name As String
	#tag EndProperty

	#tag Property, Flags = &h0
		UpdateDate As Date
	#tag EndProperty

	#tag Property, Flags = &h0
		Value As String
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="DoubleTest"
			Group="Behavior"
			Type="Double"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ID"
			Group="Behavior"
			Type="Integer"
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
		#tag ViewProperty
			Name="Value"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
