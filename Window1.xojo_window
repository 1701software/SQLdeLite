#tag Window
Begin Window Window1
   BackColor       =   &cFFFFFF00
   Backdrop        =   0
   CloseButton     =   True
   Compatibility   =   ""
   Composite       =   False
   Frame           =   0
   FullScreen      =   False
   FullScreenButton=   False
   HasBackColor    =   False
   Height          =   400
   ImplicitInstance=   True
   LiveResize      =   True
   MacProcID       =   0
   MaxHeight       =   32000
   MaximizeButton  =   True
   MaxWidth        =   32000
   MenuBar         =   344995769
   MenuBarVisible  =   True
   MinHeight       =   64
   MinimizeButton  =   True
   MinWidth        =   64
   Placement       =   0
   Resizeable      =   True
   Title           =   "Untitled"
   Visible         =   True
   Width           =   600
   Begin SQLdeLite.DatabaseCubeSQL db
      AutoCommit      =   False
      DatabaseName    =   ""
      Enabled         =   True
      Encryption      =   0
      EndChunk        =   False
      Error           =   False
      ErrorCode       =   0
      ErrorMessage    =   ""
      Height          =   "32"
      Host            =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      IsEndChunk      =   False
      Left            =   40
      LockedInPosition=   False
      Password        =   ""
      PingFrequency   =   0
      Port            =   0
      Scope           =   0
      ServerVersion   =   ""
      TabIndex        =   0
      TabPanelIndex   =   0
      TabStop         =   True
      ThreadYieldInterval=   "0"
      Timeout         =   0
      Top             =   40
      UseREALServerProtocol=   False
      Username        =   ""
      Visible         =   True
      Width           =   "32"
   End
End
#tag EndWindow

#tag WindowCode
	#tag Event
		Sub Open()
		  // Setup Database
		  'db.DatabaseFile = GetFolderItem("/Users/phillipzedalis/Bitbucket/1701-sqldelite/test.sqlite", FolderItem.PathTypeNative)
		  db.Host = "localhost"
		  db.Port = 4430
		  db.DatabaseName = "Test"
		  db.Username = "admin"
		  db.Password = "admin"
		  
		  If (Window1.db.Connect() = True) Then
		    
		    db.AutoCommit = True
		    
		    // Insert some values into the database using the ORM-like features.
		    Dim setting As New TestDb.Setting(Window1.db)
		    setting.Name = "Double Test"
		    setting.Value = App.ShortVersion
		    setting.Insert()
		    
		    Dim setting2 As New TestDb.Setting(Window1.db)
		    setting2.Name = "Bob"
		    setting2.Value = "Villa"
		    setting2.DoubleTest = 31.76
		    setting2.Insert()
		    
		    // Let's insert a row with a date. Dates have a special event handler on the Setting object.
		    Dim setting5 As New TestDb.Setting(Window1.db)
		    setting5.Name = "DateTest"
		    setting5.Value = "TestValue"
		    setting5.UpdateDate = New Date
		    setting5.Insert()
		    
		    // Use properties to send an update statement with auto bindings that replace the #?#
		    Dim setting3 As New TestDb.Setting(Window1.db)
		    setting3.Name = "Phillip"
		    setting3.Value = "Programmer"
		    setting3.execute("update setting set name = #name#, value = #value# where id = #?# and name = #?#", 5, "Bob")
		    
		    // Use a combination of values and manual bindings for the #?#. Manual bindings skips the cache and runs a tad faster.
		    Dim setting4 As New TestDb.Setting(Window1.db)
		    setting4.Value = "Awesome"
		    setting4.execute("update setting set name = #?#, value = #value# where id = #?# and name = #?#", _
		    New SQLdeLite.Parameter("Phillip", SQLitePreparedStatement.SQLITE_TEXT), _
		    New SQLdeLite.Parameter(7, SQLitePreparedStatement.SQLITE_INTEGER), _
		    New SQLdeLite.Parameter("DatabaseVersion", SQLitePreparedStatement.SQLITE_TEXT))
		    
		    
		    // Query the database
		    Dim tempRS1 As RecordSet
		    tempRS1 = Window1.db.SQLdeLiteSelect("select * from Setting where id < #?# and id > #?#", 20, 16)
		    
		    // Execute a query against the database
		    Window1.db.SQLdelIteExecute("update setting set name = #?# where id = #?#", "TEST TEST", 1)
		    
		    // Query the database yet again
		    Dim tempRS2 As RecordSet
		    tempRS2 = Window1.db.SQLdeLiteSelect("select * from setting where id > #?# and id < #?#", 0, 3)
		    
		    MsgBox("Done")
		    
		  Else
		    
		    MsgBox("Not connected.")
		    
		  End If
		  
		  
		  
		End Sub
	#tag EndEvent


#tag EndWindowCode

#tag Events db
	#tag Event
		Function CreateTableSchema(TableName As String) As Boolean
		  // Let's add the columns for the Setting table
		  If (TableName = "Setting") Then
		    // Create Table
		    db.CreateTable(TableName, "ID")
		    // Add additional columns.
		    db.CreateTableColumn(TableName, "DoubleTest")
		    db.CreateTableColumn(TableName, "Name")
		    db.CreateTableColumn(TableName, "UpdateDate")
		    db.CreateTableColumn(TableName, "Value")
		    
		    // Return True that we handled the creation.
		    Return True
		  End If
		End Function
	#tag EndEvent
	#tag Event
		Function UpdateTableSchema(TableName As String) As Boolean
		  
		End Function
	#tag EndEvent
#tag EndEvents
#tag ViewBehavior
	#tag ViewProperty
		Name="BackColor"
		Visible=true
		Group="Appearance"
		InitialValue="&hFFFFFF"
		Type="Color"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Backdrop"
		Visible=true
		Group="Appearance"
		Type="Picture"
		EditorType="Picture"
	#tag EndViewProperty
	#tag ViewProperty
		Name="CloseButton"
		Visible=true
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Composite"
		Visible=true
		Group="Appearance"
		InitialValue="False"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Frame"
		Visible=true
		Group="Appearance"
		InitialValue="0"
		Type="Integer"
		EditorType="Enum"
		#tag EnumValues
			"0 - Document"
			"1 - Movable Modal"
			"2 - Modal Dialog"
			"3 - Floating Window"
			"4 - Plain Box"
			"5 - Shadowed Box"
			"6 - Rounded Window"
			"7 - Global Floating Window"
			"8 - Sheet Window"
			"9 - Metal Window"
			"10 - Drawer Window"
			"11 - Modeless Dialog"
		#tag EndEnumValues
	#tag EndViewProperty
	#tag ViewProperty
		Name="FullScreen"
		Group="Appearance"
		InitialValue="False"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="FullScreenButton"
		Visible=true
		Group="Appearance"
		InitialValue="False"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasBackColor"
		Visible=true
		Group="Appearance"
		InitialValue="False"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Height"
		Visible=true
		Group="Position"
		InitialValue="400"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="ImplicitInstance"
		Visible=true
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Interfaces"
		Visible=true
		Group="ID"
		Type="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="LiveResize"
		Visible=true
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MacProcID"
		Visible=true
		Group="Appearance"
		InitialValue="0"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaxHeight"
		Visible=true
		Group="Position"
		InitialValue="32000"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaximizeButton"
		Visible=true
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaxWidth"
		Visible=true
		Group="Position"
		InitialValue="32000"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MenuBar"
		Visible=true
		Group="Appearance"
		Type="MenuBar"
		EditorType="MenuBar"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MenuBarVisible"
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinHeight"
		Visible=true
		Group="Position"
		InitialValue="64"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinimizeButton"
		Visible=true
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinWidth"
		Visible=true
		Group="Position"
		InitialValue="64"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Name"
		Visible=true
		Group="ID"
		Type="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Placement"
		Visible=true
		Group="Position"
		InitialValue="0"
		Type="Integer"
		EditorType="Enum"
		#tag EnumValues
			"0 - Default"
			"1 - Parent Window"
			"2 - Main Screen"
			"3 - Parent Window Screen"
			"4 - Stagger"
		#tag EndEnumValues
	#tag EndViewProperty
	#tag ViewProperty
		Name="Resizeable"
		Visible=true
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Super"
		Visible=true
		Group="ID"
		Type="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Title"
		Visible=true
		Group="Appearance"
		InitialValue="Untitled"
		Type="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Visible"
		Visible=true
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Width"
		Visible=true
		Group="Position"
		InitialValue="600"
		Type="Integer"
	#tag EndViewProperty
#tag EndViewBehavior
