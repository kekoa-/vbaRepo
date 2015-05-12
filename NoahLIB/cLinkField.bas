VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLinkField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' this class represents a link between a key-value pair in the worksheet and a key-value pair in the DB.
'

' table name
Public tableName As String

' key info
' name of the key column in the db
Public keyColumnName As String
' type of the key column in the db
Public keyType_ As String

' name of the worksheet for the key
Public keyWorksheetName As String
' name of the range for the key
Public keyRangeName As String



'value info
' type of the value column in the db
Public type_ As String
' name of the value column in the db
Public columnName As String

' name of the worksheet for the value
Public WorksheetName As String
' name of the range for the value
Public RangeName As String


' multi-cell ranges are supported
Public linkType As String


' reference workbook
Public wb As Workbook

' key in the DB
Public linkID As Long
