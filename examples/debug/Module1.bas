Attribute VB_Name = "Module1"
' Example VBA module for testing visiowings remote debugging
' This module demonstrates various debugging scenarios

Option Explicit

' Simple function with breakpoint opportunities
Public Function CalculateSum(a As Integer, b As Integer) As Integer
    Dim result As Integer
    
    ' Set breakpoint here to inspect input values
    result = a + b
    
    ' Set breakpoint here to inspect result
    CalculateSum = result
End Function

' Subroutine with multiple steps
Public Sub ProcessShapes()
    Dim shp As Visio.Shape
    Dim pg As Visio.Page
    Dim counter As Integer
    
    ' Set breakpoint here
    Set pg = ActivePage
    counter = 0
    
    ' Set breakpoint in loop to inspect each shape
    For Each shp In pg.Shapes
        counter = counter + 1
        Debug.Print "Processing shape: " & shp.Name
        
        ' Modify shape properties
        If shp.CellExists("User.Status", visExistsAnywhere) = 0 Then
            shp.AddNamedRow visSectionUser, "Status", visTagDefault
        End If
        
        ' Set breakpoint here to see shape modifications
        shp.Cells("User.Status").FormulaU = "\"Processed\""
    Next shp
    
    ' Set breakpoint here to inspect final count
    MsgBox "Processed " & counter & " shapes"
End Sub

' Function with conditional logic
Public Function GetShapeCategory(shp As Visio.Shape) As String
    Dim category As String
    
    ' Set breakpoint here
    If shp.Master Is Nothing Then
        category = "Custom"
    ElseIf InStr(shp.Master.Name, "Process") > 0 Then
        ' Set breakpoint here
        category = "Process"
    ElseIf InStr(shp.Master.Name, "Decision") > 0 Then
        category = "Decision"
    Else
        ' Set breakpoint here
        category = "Other"
    End If
    
    ' Set breakpoint here to inspect final category
    GetShapeCategory = category
End Function

' Subroutine with error handling
Public Sub SafeShapeOperation()
    On Error GoTo ErrorHandler
    
    Dim shp As Visio.Shape
    Dim pg As Visio.Page
    
    ' Set breakpoint here
    Set pg = ActivePage
    
    If pg.Shapes.Count = 0 Then
        ' Set breakpoint here to test empty page scenario
        MsgBox "No shapes found"
        Exit Sub
    End If
    
    ' Set breakpoint here
    Set shp = pg.Shapes(1)
    
    ' This might raise an error if cell doesn't exist
    ' Set breakpoint here to test error handling
    Debug.Print shp.Cells("User.CustomProperty").ResultStr("")
    
    Exit Sub
    
ErrorHandler:
    ' Set breakpoint here to inspect error
    Debug.Print "Error: " & Err.Description
    MsgBox "An error occurred: " & Err.Description
End Sub

' Nested function calls for call stack testing
Public Sub TestCallStack()
    Dim result As Integer
    
    ' Set breakpoint here - top of stack
    result = Level1Function(5)
    
    ' Set breakpoint here - after return
    MsgBox "Final result: " & result
End Sub

Private Function Level1Function(value As Integer) As Integer
    ' Set breakpoint here - one level deep
    Level1Function = Level2Function(value * 2)
End Function

Private Function Level2Function(value As Integer) As Integer
    ' Set breakpoint here - two levels deep
    Level2Function = Level3Function(value + 10)
End Function

Private Function Level3Function(value As Integer) As Integer
    ' Set breakpoint here - three levels deep (bottom of stack)
    Level3Function = value * value
End Function

' Variable inspection test
Public Sub TestVariableInspection()
    Dim intValue As Integer
    Dim strValue As String
    Dim dblValue As Double
    Dim boolValue As Boolean
    Dim objValue As Object
    
    ' Set breakpoint here to inspect initial values
    intValue = 42
    strValue = "Hello, Debugger!"
    dblValue = 3.14159
    boolValue = True
    Set objValue = ActiveDocument
    
    ' Set breakpoint here to inspect all variables
    Debug.Print "Integer: " & intValue
    Debug.Print "String: " & strValue
    Debug.Print "Double: " & dblValue
    Debug.Print "Boolean: " & boolValue
    Debug.Print "Object: " & objValue.Name
    
    ' Modify values
    intValue = intValue * 2
    strValue = strValue & " (modified)"
    
    ' Set breakpoint here to inspect modified values
    MsgBox "Variables inspected"
End Sub
