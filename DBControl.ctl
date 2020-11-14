VERSION 5.00
Begin VB.UserControl DBControl 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   465
   ScaleWidth      =   480
   ToolboxBitmap   =   "DBControl.ctx":0000
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "DBControl.ctx":0182
      Stretch         =   -1  'True
      ToolTipText     =   "CSI Database Control"
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "DBControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' CSI Database Control  Version 1.0
'
' Created 3/16/1999, By Alex Rohr
' Modified 3/17/1999
'
' The CSI Database Control allows the user to perform
' multiple tasks on any Microsoft Access 97 Database.
' This control is for the use of CSI Corporation.
'
'*************** Features *******************
' 03/16/1999  Compact Database
' 03/16/1999  Repair Database
' 03/16/1999  Alter Table
' 03/17/1999  Create Index
' 03/17/1999  Create Primary Key Index
' 03/17/1999  Drop Index
' 03/17/1999  Create Table
' 03/17/1999  Drop Table
' 03/17/1999  Add Column
' 03/17/1999  Delete Column
'****************************************************
' Private variables
Private DBs As String                          ' Database path and file name
Private Database1 As Database        ' Database object
Private A1 As Object, A2 As Object  ' Database recordsets
Private DBOptions As String
Private DBLocale As String
Private DBPassword As String
Private DBs2 As String                        ' Temp database
Private DBTable As String
Private DBField As String
Private DBColumn As String
Private DBType As String
' Public Events
Public Event DBsChanged(ByVal DBs As String)  ' Database name passed by the user
'Public Encrypt As Boolean           ' Encrypt DB Yes/No
Option Explicit
Public Function DBCompact(ByVal DBName1 As String, Optional ByVal DBName2 As String = "C:\CSITmp.mdb", Optional ByVal DstLocale As String = "", Optional ByVal Options As String = dbEncrypt, Optional ByVal Password As String = ";pwd=")
     On Error GoTo CompactError
     ' set database variables to user variables
     DBs = DBName1: DBs2 = DBName2: DBLocale = DstLocale
     DBOptions = Options: DBPassword = Password
     If Dir(DBs) = "" Or DBs = "" Then
          MsgBox "Database Was Not Located.", vbExclamation, "DB Error"
     Else
               DBEngine.CompactDatabase DBs, DBs2, DBLocale, DBOptions, DBPassword
               Kill DBs
               DBEngine.CompactDatabase DBs2, DBs, DBLocale, DBOptions, DBPassword
               Kill DBs2
     End If
     Exit Function
CompactError:
     MsgBox "Error Compacting Database.  Make Sure Database Is Closed." & vbCrLf & vbCrLf & Err.Number, vbExclamation, "Compact Error"
End Function
Public Function DBRepair(ByVal DBName1 As String)
     ' Repair database function
     On Error GoTo RepairError
     ' set database name to users database
     DBs = DBName1
     If Dir(DBs) = "" Then
          MsgBox "Database Was Not Located.", vbExclamation, "DB Error"
     Else
          DBEngine.RepairDatabase DBs
     End If
     Exit Function
RepairError:
     MsgBox "Error Repairing Database.  Make Sure The Database Is Closed.", vbExclamation, "Repair Error"
End Function
Public Function CreateDB(ByVal DBName1 As String)
     ' Create Database
     '*********** not working at this time
     On Error GoTo CreateDBError
     DBs = DBName1
     If Dir(DBs) = "" Then
          MsgBox "Database Was Not Located.", vbExclamation, "DB Error"
     Else
          MsgBox "Function Not Complete. ", vbCritical, "NO GO"
          'Set Database1 = DBEngine.CreateDatabase(DBs)
     End If
     Exit Function
CreateDBError:
     MsgBox Err.Number & vbCrLf & Err.Description
End Function
Public Function DBCreateIndex(ByVal DBName1 As String, ByVal Table As String, ByVal Field As String, Optional ByVal Password As String = ";pwd=")
     ' Create Index
     On Error GoTo IndexError
     ' set database variables to user variables
     DBs = DBName1
     DBTable = Table
     DBField = Field
     DBPassword = Password
     If Dir(DBs) = "" Then
          MsgBox "Database Was Not Located.", vbExclamation, "DB Error"
     Else
          ' open database
          Set Database1 = OpenDatabase(DBs, False, False, Password)
          ' Check for brackets on the field name
          If Left(DBField, 1) = "[" And Right(DBField, 1) = "]" Then
               ' create index
               Database1.Execute "Create Index MyIndex1 ON " & DBTable & " (" & DBField & ")"
          Else
               ' create index
               Database1.Execute "Create Index MyIndex1 ON " & DBTable & " ([" & DBField & "])"
          End If
          ' close the open database
          Database1.Close
     End If
     Exit Function
IndexError:
     MsgBox "Error Creating Index. " & Err.Number, vbExclamation, "Index Error"
End Function
Public Function DBPrimaryIndex(ByVal DBName1 As String, ByVal Table As String, ByVal Field As String, Optional ByVal Password As String = ";pwd=")
     ' Create Primary Index
     On Error GoTo IndexError
     ' set database variables to user variables
     DBs = DBName1
     DBTable = Table
     DBField = Field
     DBPassword = Password
     If Dir(DBs) = "" Then
          MsgBox "Database Was Not Located.", vbExclamation, "DB Error"
     Else
          ' open database
          Set Database1 = OpenDatabase(DBs, False, False, Password)
          ' check for brackets [ ]
          If Left(DBField, 1) = "[" And Right(DBField, 1) = "]" Then
               ' create primary key
               Database1.Execute "Create Unique Index MyIndex1 ON " & DBTable & " (" & DBField & ")"
          Else
               ' create primary key
               Database1.Execute "Create Unique Index MyIndex1 ON " & DBTable & " ([" & DBField & "])"
          End If
          ' close open database
          Database1.Close
     End If
     Exit Function
IndexError:
     MsgBox "Error Creating Primary Index. " & Err.Number, vbExclamation, "Index Error"
End Function
Public Function DBAddColumn(ByVal DBName1 As String, ByVal Table As String, ByVal Column As String, Optional ByVal DBType1 As String = "Text", Optional ByVal Password As String)
     ' Add a specific column
     On Error GoTo AddColumnError
     ' set database variables to user variables
     DBs = DBName1
     DBTable = Table
     DBColumn = Column
     DBPassword = Password
     DBType = DBType1
     If Dir(DBs) = "" Then
          MsgBox "Database Was Not Located.", vbExclamation, "DB Error"
     Else
          ' open database
          Set Database1 = OpenDatabase(DBs, False, False, Password)
          ' add column to existing table
          Database1.Execute ("Alter Table " & DBTable & " Add Column " & DBColumn & " " & DBType)
          Database1.Close
     End If
     Exit Function
AddColumnError:
     MsgBox "Error Adding Column.  Make Sure The Table Name & Column Name Are Encolsed With Brackets. [ ]", vbExclamation, "Add Column Error"
End Function
Public Function DBDeleteColumn(ByVal DBName1 As String, ByVal Table As String, ByVal Column As String, Optional ByVal Password As String)
     ' Delete a specific column
     On Error GoTo DeleteColumnError
     ' set database variables to user variables
     DBs = DBName1
     DBTable = Table
     DBColumn = Column
     DBPassword = Password
     If Dir(DBs) = "" Then
          MsgBox "Database Was Not Located.", vbExclamation, "DB Error"
     Else
          ' open database
          Set Database1 = OpenDatabase(DBs, False, False, Password)
          ' Delete column from database table
          Database1.Execute ("Alter Table " & DBTable & " Drop Column " & DBColumn)
          Database1.Close
     End If
     Exit Function
DeleteColumnError:
     MsgBox "Error Deleting Column.  Make Sure The Table Name Is Encolsed With Brackets. [ ]", vbExclamation, "Delete Column Error"
End Function
Public Function DBCreateTable(ByVal DBName1 As String, ByVal Table As String, Optional ByVal Password As String = ";pwd=")
     ' Add database table
     On Error GoTo AddError
     ' set database name to users passed name
     DBs = DBName1
     DBPassword = Password
     If Dir(DBs) = "" Then
        MsgBox "Database Was Not Located.", vbExclamation, "DB Error"
     Else
          If Left(Table, 1) = "[" And Right(Table, 1) = "]" Then
               DBTable = Table
          Else
               DBTable = "[" & Table & "]"
          End If
          ' open database
          Set Database1 = OpenDatabase(DBs, False, False, Password)
          ' create new table with one field called Key
          Database1.Execute "Create Table " & DBTable & " ([Key] number)"
          Database1.Close
     End If
     Exit Function
AddError:
     MsgBox "Error Deleting Table " & DBTable & "." & vbCrLf & "Or Table Does Not Exist.", vbExclamation, "Drop Error"
End Function
Public Function DBDeleteTable(ByVal DBName1 As String, ByVal Table As String, Optional ByVal Password As String = ";pwd=")
     ' Drop database table
     On Error GoTo DropError
     ' set database name to users passed name
     DBs = DBName1
     DBTable = Table
     DBPassword = Password
     If Dir(DBs) = "" Then
          MsgBox "Database Was Not Located.", vbExclamation, "DB Error"
     Else
          ' open database
          Set Database1 = OpenDatabase(DBs, False, False, Password)
          ' delete table from database
          Database1.Execute "Drop Table " & DBTable
          Database1.Close
     End If
     Exit Function
DropError:
     MsgBox "Error Deleting Table " & DBTable & "." & vbCrLf & "Or Table Does Not Exist.", vbExclamation, "Drop Error"
End Function
Private Sub UserControl_ReadProperties(PB As PropertyBag)
     ' Read the database name from the property bag
     DataBaseName = PB.ReadProperty("Database Name", DataBaseName)
End Sub
Private Sub UserControl_WriteProperties(PB As PropertyBag)
     ' Write the database name to the property bag
     PB.WriteProperty "Database Name", DataBaseName
End Sub
Public Property Get DataBaseName() As String
     ' set the database name to the database object
     DataBaseName = DBs
End Property
Public Property Let DataBaseName(ByVal DBName1 As String)
     DBs = DBName1
     RaiseEvent DBsChanged(DBName1)
     UserControl.PropertyChanged (DataBaseName)
End Property
