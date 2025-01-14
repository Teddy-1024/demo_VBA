' Author: Edward Middleton-Smith
' Precision And Research Technology Systems Limited

' Project:
' Technology:
' Feature:



' NOTES



' METHOD LAYOUT COMMENTS
' FUNCTION
' ARGUMENTS
' PROCESSING ACCELERATION
' CONSTANTS
' VARIABLE DECLARATION
' ARGUMENT VALIDATION
' VARIABLE INSTANTIATION
' METHODS
' RETURNS
' ERROR HANDLING
' PROCESSING DECELARATION



' MODULE INITIALISATION
' Set array start index to 1 to match spreadsheet indices
Option Base 1
' Forced Variable Declaration
Option Explicit


' Worksheet orientation
Enum orientation
    ColumnHeaders = 0
    RowHeaders = 1
End Enum


' Array search direction
Enum direction
    x = 1
    y = 2
    Z = 3
End Enum


' Different search direction
Enum dir_traverse
    FORWARDS = 1
    BACKWARDS = -1
End Enum


' Customer account status
Enum OverdueStatus
    UNDUE = 0
    OVERDUE = 1
    SUPEROVERDUE = 2
End Enum



