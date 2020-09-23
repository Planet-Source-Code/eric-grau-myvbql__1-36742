VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MYSQL_RS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"MYSQL_FIELD"
Attribute VB_Ext_KEY = "Member0" ,"MYSQL_FIELD"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'###############################################################################
'###############################################################################
'
' Copyright (C) 2000,2001 icarz, Inc.
'
' This library is free software; you can redistribute it and/or
' modify it under the terms of the GNU Library General Public
' License as published by the Free Software Foundation; either
' version 2 of the License, or (at your option) any later version.
'
' This library is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
' Library General Public License for more details.
'
' You should have received a copy of the GNU Library General Public
' License along with this library; if not, write to the Free
' Software Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.
'
'###############################################################################
'###############################################################################
'
' Written by Eric Grau
'
' Please send questions, comments, and changes to mysql@icarz.com
'
'###############################################################################
'###############################################################################
'

Option Explicit

'===============================================================================================
'========= private variables ===================================================================
'===============================================================================================

Private MyCONN As MYSQL_CONNECTION                         'connection object that "owns" this rs object
Private MyFIELDS As MYSQL_FIELD                            'field object
Private mlCurrentRecord As Long                            'current record #
Private mlRecordCount As Long                              'number of records in result set
Private mlFieldCount As Long                               'number of fields in result set
Private mlAffectedRecords As Long                          'number of rows affected by query
Private msQueryString As String                            'last sql statement queried on this rs object
Private myState As MYSQL_RECORDSET_STATE                   'state of this rs object
Private mlMYSQL_RES As Long                                'pointer to mysql result set
Private mlMYSQL_ROW As Long                                'pointer to mysql row
Private mlMYSQL_FIELD_LENGTHS As Long                      'pointer to array of column lengths
Private MyAdd As MYSQL_UPDATE                              'used to hold added rows
Private MyReplace As MYSQL_UPDATE                          'used to hold replaced rows
Private MyUpdate As MYSQL_UPDATE                           'used to hold updated rows
Private MyDelete As MYSQL_UPDATE                           'used to hold deleted rows
Private mlUpdateRecord As Long                             'record # currently being editted
Private MyEditMode As MYSQL_EDIT_MODE                      'current edit mode

'===============================================================================================
'========= rs state enum =======================================================================
'===============================================================================================

Public Enum MYSQL_RECORDSET_STATE
    MY_RS_OPEN = 0
    MY_RS_CLOSED = 1
End Enum

'===============================================================================================
'========= edit mode enum ======================================================================
'===============================================================================================

Public Enum MYSQL_EDIT_MODE
    MY_EDIT_NONE = 0
    MY_EDIT_IN_PROGRESS = 1
    MY_EDIT_ADD = 2
    MY_EDIT_DELETE = 3
    'MY_EDIT_REPLACE = 4
End Enum

'===============================================================================================
'========= absolute position enum ==============================================================
'===============================================================================================

Public Enum MYSQL_ABSOLUTE_POSITION
    MY_POS_BOF = -3
    MY_POS_EOF = -2
    MY_POS_UNKNOWN = -1
End Enum

'===============================================================================================
'========= Find enum ===========================================================================
'===============================================================================================

Public Enum MYSQL_FIND
    MY_FIND_AND_RETURN_POSITION_ONLY = 1
    MY_FIND_GOTO_RECORD_AND_RETURN_POSITION = 2
End Enum

'===============================================================================================
'========= public properties ===================================================================
'===============================================================================================

Public Property Get State() As MYSQL_RECORDSET_STATE
Attribute State.VB_Description = "Returns the state of the rs object."
    State = myState
End Property

Public Property Get AffectedRecords() As Long
Attribute AffectedRecords.VB_Description = "Returns the number of records affected by the last query."
    AffectedRecords = mlAffectedRecords
End Property

Public Property Get QueryString() As String
Attribute QueryString.VB_Description = "Returns the last sql statement executed on this rs."
    QueryString = msQueryString
End Property

Public Property Get RecordCount() As Long
Attribute RecordCount.VB_Description = "Returns the number of rows in this result set."
    RecordCount = mlRecordCount
End Property

Public Property Get FieldCount() As Long
Attribute FieldCount.VB_Description = "Returns the number of fields in this result set."
    FieldCount = mlFieldCount
End Property

Public Property Get Fields(ByVal vIndexKey As Variant) As MYSQL_FIELD
Attribute Fields.VB_Description = "Returns the requested field object.  Field may be reference by name or by number (0 to FieldCount - 1)."
    'used when referencing an element in the rs
    If Me.BOF _
            Or Me.EOF Then
        'setup invalid properties ...
        MyFIELDS.SetProperties -1, 0, 0, 0, 0
    Else
        'setup properties for field object
        MyFIELDS.SetProperties vIndexKey, mlFieldCount, mlMYSQL_RES, mlMYSQL_ROW, mlMYSQL_FIELD_LENGTHS
    End If

    'return the field object
    Set Fields = MyFIELDS
End Property

Public Property Get BOF() As Boolean
Attribute BOF.VB_Description = "Returns True if rs is at the BOF pointer."
    BOF = (myState = MY_RS_CLOSED) Or (mlRecordCount = 0) Or (mlCurrentRecord = 0)
End Property

Public Property Get EOF() As Boolean
Attribute EOF.VB_Description = "Returns True if rs is at the EOF pointer."
    EOF = (myState = MY_RS_CLOSED) Or (mlRecordCount = 0) Or (mlCurrentRecord > mlRecordCount)
End Property

Public Property Get AbsolutePosition() As Long
Attribute AbsolutePosition.VB_Description = "Returns/set the result set position (1 to RecordCount).  Unknown, bof, or eof values may also be returned."
    If Me.State = MY_RS_CLOSED _
            Or Me.RecordCount = 0 Then
        AbsolutePosition = MY_POS_UNKNOWN
    ElseIf Me.BOF Then
        AbsolutePosition = MY_POS_BOF
    ElseIf Me.EOF Then
        AbsolutePosition = MY_POS_EOF
    Else
        AbsolutePosition = mlCurrentRecord
    End If
End Property

Public Property Let AbsolutePosition(ByVal lPosition As Long)
    If Me.State = MY_RS_CLOSED _
            Or Me.RecordCount = 0 Then
        'can't set position when rs is closed or rs has no records
        Exit Property
    ElseIf lPosition > 0 _
            And lPosition <= mlRecordCount Then
        'set the record #
        mlCurrentRecord = lPosition
        'reposition the record pointer
        GetRow
    End If
End Property

Public Property Get TableCount() As Long
Attribute TableCount.VB_Description = "Returns the number of tables in this result set."
    Dim lCnt As Long
    Dim vTbl As Variant
    vTbl = BuildTableArray

    If IsArray(vTbl) Then
        TableCount = UBound(vTbl) + 1
    Else
        TableCount = 0
    End If
End Property

Public Property Get Table() As String
Attribute Table.VB_Description = "Returns the names of the tables in this result set.  If there are more that one table, the name(s) will be separated by commas."
    Dim lCnt As Long
    Dim vTbl As Variant
    vTbl = BuildTableArray

    If IsArray(vTbl) Then
        lCnt = 0
        Do
            'when there's more than 1 table, send back a comma delimited list
            Table = Table & IIf(lCnt > 0, ",", "") & vTbl(lCnt)
            lCnt = lCnt + 1
        Loop While lCnt <= UBound(vTbl)
    Else
        Table = ""
    End If
End Property

Public Property Get EditMode() As MYSQL_EDIT_MODE
Attribute EditMode.VB_Description = "Returns the edit mode of the result set."
    Dim uDel As MYSQL_UPDATE_FIELDS

    If Not MyDelete Is Nothing Then
        'if this returns an object that's not nothing, this row's been deleted
        Set uDel = MyDelete.Item(Me.AbsolutePosition, True)
    End If

    EditMode = IIf(uDel Is Nothing, MyEditMode, MY_EDIT_DELETE)
End Property

'===============================================================================================
'========= friend properties ===================================================================
'===============================================================================================

Friend Property Let EditMode(my_ed As MYSQL_EDIT_MODE)
    MyEditMode = my_ed
    'when edit mode is set to none ... clear the update record #
    If MyEditMode = MY_EDIT_NONE Then mlUpdateRecord = 0
End Property

Friend Property Get Connection() As MYSQL_CONNECTION
    Set Connection = MyCONN
End Property

Friend Property Set Connection(cn As MYSQL_CONNECTION)
    Set MyCONN = cn
End Property

Friend Property Get MyUpdateObject() As MYSQL_UPDATE
    Set MyUpdateObject = MyUpdate
End Property

Friend Property Set MyUpdateObject(my_upd As MYSQL_UPDATE)
    Set MyUpdate = my_upd
End Property

Friend Property Get MyAddObject() As MYSQL_UPDATE
    Set MyAddObject = MyAdd
End Property

Friend Property Set MyAddObject(my_add As MYSQL_UPDATE)
    Set MyAdd = my_add
End Property

Friend Property Get UpdateRecordNum() As Long
    'make sure the update record # is 0 when edit mode is set to none
    If MyEditMode = MY_EDIT_NONE Then mlUpdateRecord = 0
    UpdateRecordNum = mlUpdateRecord
End Property

Friend Property Let UpdateRecordNum(lRec As Long)
    mlUpdateRecord = lRec
End Property

'===============================================================================================
'========= friend methods ======================================================================
'===============================================================================================

Friend Function OpenRecordset(ByRef sSQL As String, ByRef lMYSQL As Long, ByRef bGotError As Boolean) As MYSQL_RECORDSET_STATE
    Dim lRc As Long

    'in case we already have an open rs
    Me.CloseRecordset
    'default to zero affected records ...
    mlAffectedRecords = 0
    'start with a return value of rs state closed
    OpenRecordset = myState

    'Debug.Print "Sending query: " & sSQL

    'save the query string
    msQueryString = sSQL
    'run the query on the server
    lRc = mysql_query(lMYSQL, sSQL)
    'Debug.Print "Query return value = " & lRc

    If lRc <> 0 Then
        'query failed
        bGotError = True
        'Debug.Print "Query failed... have to check for errors."
        Exit Function
    End If

    'return the result pointer
    mlMYSQL_RES = mysql_store_result(lMYSQL)
    'Debug.Print "Pointer to result = " & mlMYSQL_RES

    If mlMYSQL_RES = 0 Then
        'check if we should have gotten data back
        lRc = mysql_field_count(lMYSQL)

        If lRc = 0 Then
            'no result set, no error
            mlAffectedRecords = mysql_affected_rows(lMYSQL)
            'Debug.Print "Query affected " & mlAffectedRows & " rows."
        Else
            'we have an error
            bGotError = True
            'Debug.Print "Query should have returned result... have to check for errors."
            Exit Function
        End If
    Else
        'we got data back ...
        'setup module wide variables
        mlAffectedRecords = mysql_affected_rows(lMYSQL)
        'Debug.Print "Number of affected rows = " & mlAffectedRecords
        mlRecordCount = mysql_num_rows(mlMYSQL_RES)
        'Debug.Print "Number of rows  = " & mlRecordCount
        mlFieldCount = mysql_num_fields(mlMYSQL_RES)
        'Debug.Print "Number of fields  = " & mlFieldCount

        'start by pointing to row #1
        mlCurrentRecord = 1
        'rs is now open
        myState = MY_RS_OPEN

        'reposition the record pointer
        GetRow
    End If
End Function

'===============================================================================================
'========= public methods ======================================================================
'===============================================================================================

Public Sub MoveFirst()
Attribute MoveFirst.VB_Description = "Move to the first row in the result set."
    If myState = MY_RS_CLOSED Then Exit Sub

    'point to first record
    mlCurrentRecord = 1
    'reposition the record pointer
    GetRow
End Sub

Public Sub MoveLast()
Attribute MoveLast.VB_Description = "Move to the last row in the result set."
    If myState = MY_RS_CLOSED Then Exit Sub

    'point to last record
    mlCurrentRecord = mlRecordCount
    'reposition the record pointer
    GetRow
End Sub

Public Sub MovePrevious()
Attribute MovePrevious.VB_Description = "Move to the previous row in the result set."
    If myState = MY_RS_CLOSED Then Exit Sub

    'check if we're already at BOF
    If mlCurrentRecord > 0 Then
        'point to previous record
        mlCurrentRecord = mlCurrentRecord - 1
        'reposition the record pointer
        GetRow
    End If
End Sub

Public Sub MoveNext()
Attribute MoveNext.VB_Description = "Move to the next row in the result set."
    If myState = MY_RS_CLOSED Then Exit Sub

    'check if we're already at EOF
    If mlCurrentRecord <= mlRecordCount Then
        'point to next record
        mlCurrentRecord = mlCurrentRecord + 1
        'reposition the record pointer
        GetRow
    End If
End Sub

Public Sub CloseRecordset()
Attribute CloseRecordset.VB_Description = "Close the record set."
    'close rs if it's currently open
    If myState = MY_RS_OPEN Then
        'check if we need to update first have to update first
        CheckAutoUpdate
        'now close it up
        RealClose
    End If
End Sub

Public Function FindFirst(ByVal vField As Variant, ByVal vData As Variant, ByVal FindMethod As MYSQL_FIND) As Long
Attribute FindFirst.VB_Description = "Find the first row where the given field matches the given value."
    Dim lOldPosition As Long
    Dim lCount As Long
    Dim bToggle As Boolean

    If Me.RecordCount = 0 Or myState = MY_RS_CLOSED Then
        ' get out of the function, since we have no records
        FindFirst = -1
        Exit Function
    End If

    If IsNumeric(vField) Then
        ' check if it's out of bounds
        If vField >= 0 _
                And vField < Me.FieldCount Then bToggle = True
    Else
        ' check if the field name exists
        For lCount = 0 To Me.FieldCount - 1
            If UCase(Trim(Me.Fields(lCount).Name)) = UCase(Trim(vField)) Then
                bToggle = True
                Exit For
            End If
        Next lCount
    End If

    If Not bToggle Then
        ' get out of the function since the field doesnt exist.
        FindFirst = -1
        Exit Function
    End If

    If FindMethod = MYSQL_FIND.MY_FIND_GOTO_RECORD_AND_RETURN_POSITION Then
        'auto udpate as needed ...
        CheckAutoUpdate
    End If

    ' get the current record in case we are going to return to it
    lOldPosition = mlCurrentRecord

    ' set this to 0 so our do/loop functions correctly
    mlCurrentRecord = 0

    ' reset this variable
    bToggle = False

    Do Until mlCurrentRecord = RecordCount
        ' go to next row.
        mlCurrentRecord = mlCurrentRecord + 1
        'reposition the record pointer
        GetRow True
        If Me.Fields(vField).Value = vData Then
            ' set the flag to true, we found it
            bToggle = True
            Exit Do
        End If
    Loop

    If Not bToggle Then
        ' couldnt find the record. exit with a -1, and return to
        ' the record we were at
        mlCurrentRecord = lOldPosition
        'reposition the record pointer
        GetRow True
        FindFirst = -1
        Exit Function
    End If

    ' if we are here, then the record was found.
    FindFirst = mlCurrentRecord

    Select Case FindMethod
        Case MYSQL_FIND.MY_FIND_AND_RETURN_POSITION_ONLY
            mlCurrentRecord = lOldPosition
            'reposition the record pointer
            GetRow True
        Case MYSQL_FIND.MY_FIND_GOTO_RECORD_AND_RETURN_POSITION
            ' we are already on the record we found ...
    End Select
End Function

Public Function FindLast(ByVal vField As Variant, ByVal vData As Variant, ByVal FindMethod As MYSQL_FIND) As Long
Attribute FindLast.VB_Description = "Find the last row where the given field matches the given value."
    Dim lOldPosition As Long
    Dim lCount As Long
    Dim bToggle As Boolean

    If Me.RecordCount = 0 Or myState = MY_RS_CLOSED Then
        ' get out of the function, since we have no records
        FindLast = -1
        Exit Function
    End If

    If IsNumeric(vField) Then
        ' check if it's out of bounds
        If vField >= 0 _
                And vField < Me.FieldCount Then bToggle = True
    Else
        ' check if the field name exists
        For lCount = 0 To Me.FieldCount - 1
            If UCase(Trim(Me.Fields(lCount).Name)) = UCase(Trim(vField)) Then
                bToggle = True
                Exit For
            End If
        Next lCount
    End If

    If Not bToggle Then
        ' get out of the function since the field doesnt exist.
        FindLast = -1
        Exit Function
    End If

    If FindMethod = MYSQL_FIND.MY_FIND_GOTO_RECORD_AND_RETURN_POSITION Then
        'auto udpate as needed ...
        CheckAutoUpdate
    End If

    ' get the current record in case we are going to return to it
    lOldPosition = mlCurrentRecord

    ' set this to RecordCount + 1 so our do/loop functions correctly
    mlCurrentRecord = RecordCount + 1

    ' reset this variable
    bToggle = False

    Do Until mlCurrentRecord = 1
        ' go to next row.
        mlCurrentRecord = mlCurrentRecord - 1
        'reposition the record pointer
        GetRow True
        If Me.Fields(vField).Value = vData Then
            ' set the flag to true, we found it
            bToggle = True
            Exit Do
        End If
    Loop

    If Not bToggle Then
        ' couldnt find the record. exit with a -1, and return to
        ' the record we were at
        mlCurrentRecord = lOldPosition
        'reposition the record pointer
        GetRow True
        FindLast = -1
        Exit Function
    End If

    ' if we are here, then the record was found.
    FindLast = mlCurrentRecord

    Select Case FindMethod
        Case MYSQL_FIND.MY_FIND_AND_RETURN_POSITION_ONLY
            mlCurrentRecord = lOldPosition
            'reposition the record pointer
            GetRow True
        Case MYSQL_FIND.MY_FIND_GOTO_RECORD_AND_RETURN_POSITION
            ' we are already on the record we found ...
    End Select
End Function

Public Function FindNext(ByVal vField As Variant, ByVal vData As Variant, ByVal FindMethod As MYSQL_FIND) As Long
Attribute FindNext.VB_Description = "Find the next row where the given field matches the given value."
    Dim lOldPosition As Long
    Dim lCount As Long
    Dim bToggle As Boolean

    If Me.RecordCount = 0 Or myState = MY_RS_CLOSED Then
        ' get out of the function, since we have no records
        FindNext = -1
        Exit Function
    End If

    If IsNumeric(vField) Then
        ' check if it's out of bounds
        If vField >= 0 _
                And vField < Me.FieldCount Then bToggle = True
    Else
        ' check if the field name exists
        For lCount = 0 To Me.FieldCount - 1
            If UCase(Trim(Me.Fields(lCount).Name)) = UCase(Trim(vField)) Then
                bToggle = True
                Exit For
            End If
        Next lCount
    End If

    If Not bToggle Then
        ' get out of the function since the field doesnt exist.
        FindNext = -1
        Exit Function
    End If

    If FindMethod = MYSQL_FIND.MY_FIND_GOTO_RECORD_AND_RETURN_POSITION Then
        'auto udpate as needed ...
        CheckAutoUpdate
    End If

    ' get the current record in case we are going to return to it
    lOldPosition = mlCurrentRecord

    ' reset this variable
    bToggle = False

    Do Until mlCurrentRecord = RecordCount
        ' go to next row.
        mlCurrentRecord = mlCurrentRecord + 1
        'reposition the record pointer
        GetRow True
        If Me.Fields(vField).Value = vData Then
            ' set the flag to true, we found it
            bToggle = True
            Exit Do
        End If
    Loop

    If Not bToggle Then
        ' couldnt find the record. exit with a -1, and return to
        ' the record we were at
        mlCurrentRecord = lOldPosition
        'reposition the record pointer
        GetRow True
        FindNext = -1
        Exit Function
    End If

    ' if we are here, then the record was found.
    FindNext = mlCurrentRecord

    Select Case FindMethod
        Case MYSQL_FIND.MY_FIND_AND_RETURN_POSITION_ONLY
            mlCurrentRecord = lOldPosition
            'reposition the record pointer
            GetRow True
        Case MYSQL_FIND.MY_FIND_GOTO_RECORD_AND_RETURN_POSITION
            ' we are already on the record we found ...
    End Select
End Function

Public Function FindPrevious(ByVal vField As Variant, ByVal vData As Variant, ByVal FindMethod As MYSQL_FIND) As Long
Attribute FindPrevious.VB_Description = "Find the previous row where the given field matches the given value."
    Dim lOldPosition As Long
    Dim lCount As Long
    Dim bToggle As Boolean

    If Me.RecordCount = 0 Or myState = MY_RS_CLOSED Then
        ' get out of the function, since we have no records
        FindPrevious = -1
        Exit Function
    End If

    If IsNumeric(vField) Then
        ' check if it's out of bounds
        If vField >= 0 _
                And vField < Me.FieldCount Then bToggle = True
    Else
        ' check if the field name exists
        For lCount = 0 To Me.FieldCount - 1
            If UCase(Trim(Me.Fields(lCount).Name)) = UCase(Trim(vField)) Then
                bToggle = True
                Exit For
            End If
        Next lCount
    End If

    If Not bToggle Then
        ' get out of the function since the field doesnt exist.
        FindPrevious = -1
        Exit Function
    End If

    If FindMethod = MYSQL_FIND.MY_FIND_GOTO_RECORD_AND_RETURN_POSITION Then
        'auto udpate as needed ...
        CheckAutoUpdate
    End If

    ' get the current record in case we are going to return to it
    lOldPosition = mlCurrentRecord

    ' reset this variable
    bToggle = False

    Do Until mlCurrentRecord = 1
        ' go to next row.
        mlCurrentRecord = mlCurrentRecord - 1
        'reposition the record pointer
        GetRow True
        If Me.Fields(vField).Value = vData Then
            ' set the flag to true, we found it
            bToggle = True
            Exit Do
        End If
    Loop

    If Not bToggle Then
        ' couldnt find the record. exit with a -1, and return to
        ' the record we were at
        mlCurrentRecord = lOldPosition
        'reposition the record pointer
        GetRow True
        FindPrevious = -1
        Exit Function
    End If

    ' if we are here, then the record was found.
    FindPrevious = mlCurrentRecord

    Select Case FindMethod
        Case MYSQL_FIND.MY_FIND_AND_RETURN_POSITION_ONLY
            mlCurrentRecord = lOldPosition
            'reposition the record pointer
            GetRow True
        Case MYSQL_FIND.MY_FIND_GOTO_RECORD_AND_RETURN_POSITION
            ' we are already on the record we found ...
    End Select
End Function

Public Function RsToHtmlTable() As String
Attribute RsToHtmlTable.VB_Description = "Creates an HTML table based on the data in the record set."
    Dim lOldPosition As Long
    Dim lCount As Long
    Dim v As Variant

    If Me.RecordCount = 0 Or myState = MY_RS_CLOSED Then
        'no data
        Exit Function
    End If

    ' get the current record ... we are going to return to it
    lOldPosition = mlCurrentRecord
    ' set this to 0 so our do/loop functions correctly
    mlCurrentRecord = 0

    'loop thru records to build html table
    Do Until mlCurrentRecord >= mlRecordCount
        ' go to next row.
        mlCurrentRecord = mlCurrentRecord + 1
        'reposition the record pointer
        GetRow True

        RsToHtmlTable = RsToHtmlTable & "<tr>"
        For lCount = 0 To (mlFieldCount - 1)
            v = Me.Fields(lCount).Value
            If IsNull(v) Then
                v = "&nbsp;"
            ElseIf Len(Trim(v)) = 0 Then
                v = "&nbsp;"
            End If
            RsToHtmlTable = RsToHtmlTable & "<td>" & v & "</td>"
        Next lCount
        RsToHtmlTable = RsToHtmlTable & "</tr>" & vbCrLf
    Loop
    RsToHtmlTable = "<table border=1>" & RsToHtmlTable & "</table>" & vbCrLf

    'return to the record we were at
    mlCurrentRecord = lOldPosition
    'reposition the record pointer
    GetRow True
End Function

Public Function AddNew() As Long
Attribute AddNew.VB_Description = "Enables add mode."
    AddNew = -1
    If State = MY_RS_CLOSED Then Exit Function             'can't add if rs is closed
    If Me.EditMode <> MY_EDIT_NONE Then Exit Function      'shouldn't be here if any edit is in progress
    If Me.TableCount <> 1 Then Exit Function               'can't add if there's more than 1 table in rs

    'set edit mode to "add"
    MyEditMode = MY_EDIT_ADD
    'if we haven't done any add's yet, make a new update object for adding
    If MyAdd Is Nothing Then Set MyAdd = New MYSQL_UPDATE
    'create new add record so we can start "letting" fields
    MyAdd.Add
    AddNew = 0
End Function

Public Function Update() As Long
Attribute Update.VB_Description = "Updates any changes to the current row or adds a new record."
    Update = -1
    If State = MY_RS_CLOSED Or BOF Or EOF Then Exit Function    'can't update if we aren't pointing at a record
    If Me.TableCount <> 1 Then Exit Function               'can't update if there's more than 1 table in rs

    Select Case Me.EditMode
        Case MY_EDIT_IN_PROGRESS
            'update the record
            Update = DoUpdate
            If Update > 0 Then
                'success ...
                SetUpdatedFlags
            End If
        Case MY_EDIT_ADD
            'add the record
            Update = DoAdd
            If Update < 1 Then
                'failure ...
                Me.MyAddObject.Remove Me.MyAddObject.Count
            End If
        Case Else
            Exit Function                                  'shouldn't be here
    End Select

    'done with edit ...
    Me.EditMode = MY_EDIT_NONE
End Function

Public Function Delete() As Long
Attribute Delete.VB_Description = "Delete the current row."
    Delete = -1
    If State = MY_RS_CLOSED Or BOF Or EOF Then Exit Function    'can't delete if we aren't pointing at a record
    If Me.EditMode <> MY_EDIT_NONE Then Exit Function      'shouldn't be here if any edit is in progress
    If Me.TableCount <> 1 Then Exit Function               'can't delete if there's more than 1 table in rs

    'delete the record
    Delete = DoDelete

    If Delete > 0 Then
        'success ...
        If MyDelete Is Nothing Then Set MyDelete = New MYSQL_UPDATE
        'track the deleted record
        MyDelete.Add Me.AbsolutePosition
    End If
End Function

Public Sub CancelUpdate()
Attribute CancelUpdate.VB_Description = "Cancels any edits in progress."
    If State = MY_RS_CLOSED Or BOF Or EOF Then Exit Sub    'cant' cancel if we aren't pointing at a record

    Select Case Me.EditMode
        Case MY_EDIT_IN_PROGRESS
            'currently editting ...
            Me.MyUpdateObject.Remove mlUpdateRecord, True
        Case MY_EDIT_ADD
            'currently adding ...
            Me.MyAddObject.Remove Me.MyAddObject.Count
        Case Else
            Exit Sub                                       'shouldn't be here
    End Select

    'done with edit ...
    Me.EditMode = MY_EDIT_NONE
End Sub

Public Sub Requery()
Attribute Requery.VB_Description = "Requeries the record set."
    'execute the current query again ...
    MyCONN.InternalExecute Me, Me.QueryString
End Sub

'===============================================================================================
'========= private methods =====================================================================
'===============================================================================================

Private Sub Class_Initialize()
    'create a new field object
    Set MyFIELDS = GetNewField
    'set the rs state to closed
    myState = MY_RS_CLOSED
End Sub

Private Sub Class_Terminate()
    'insures that the rs is closed
    RealClose
    'release my objects
    Set MyCONN = Nothing
    Set MyFIELDS = Nothing
End Sub

Private Sub RealClose()
    'release result set memory
    mysql_free_result mlMYSQL_RES

    'result my module wide variables
    mlMYSQL_RES = 0
    mlMYSQL_ROW = 0
    mlMYSQL_FIELD_LENGTHS = 0
    mlCurrentRecord = 0
    mlRecordCount = 0
    mlFieldCount = 0
    mlAffectedRecords = 0
    msQueryString = ""

    'create a new field object
    Set MyFIELDS = GetNewField

    'release my edit objects
    Set MyAdd = Nothing
    Set MyReplace = Nothing
    Set MyUpdate = Nothing
    Set MyDelete = Nothing

    'rs is now closed
    myState = MY_RS_CLOSED
    'Debug.Print "MySQL RS closed."
End Sub

Private Sub CheckAutoUpdate()
    If MyCONN.AutoUpdate Then
        'auto update is enabled ...
        'update the record as needed
        Update
    Else
        'cancel any updated in progress
        CancelUpdate
    End If

    'done with edit ...
    MyEditMode = MY_EDIT_NONE
End Sub

Private Function GetNewField() As MYSQL_FIELD
    'create and return a new field object
    Dim MyField As New MYSQL_FIELD
    Set MyField.rs = Me
    Set GetNewField = MyField
End Function

Private Sub GetRow(Optional bRetainEditObjects As Boolean)
    'auto update as needed ...
    If Not bRetainEditObjects Then CheckAutoUpdate

    'are we on a valid record ???
    If mlCurrentRecord > 0 _
            And mlCurrentRecord <= mlRecordCount Then
        Dim cSeekRow As Currency

        'adjust because currency fields have 4 fixed decimals
        ' 1.000 gets adjusted to 0.001
        cSeekRow = (mlCurrentRecord - 1) / (10 ^ 4)

        'move to the specified row ...
        mysql_data_seek mlMYSQL_RES, cSeekRow

        'get pointer to row ...
        mlMYSQL_ROW = mysql_fetch_row(mlMYSQL_RES)
        'Debug.Print "Pointer to mysql row = " & mlMYSQL_ROW

        'get pointer to array of field lengths ...
        mlMYSQL_FIELD_LENGTHS = mysql_fetch_lengths(mlMYSQL_RES)
        'Debug.Print "Pointer to array of column lengths = " & mlMYSQL_FIELD_LENGTHS
    Else
        'not on a valid record ...
        mlMYSQL_ROW = 0
        mlMYSQL_FIELD_LENGTHS = 0
    End If
End Sub

Private Function BuildTableArray() As Variant
    Dim lCnt As Long
    Dim lCnt2 As Long
    Dim sTbl As String
    Dim vTbl As Variant
    Dim bInArr As Boolean

    'check the table name associated with each field ...
    'track unique table names in array ...
    For lCnt = 0 To (FieldCount - 1)
        sTbl = Trim(Fields(lCnt).Table)
        If Len(sTbl) > 0 Then
            If Not IsArray(vTbl) Then
                ReDim vTbl(0)
                vTbl(0) = sTbl
            Else
                bInArr = False
                lCnt2 = 0
                Do
                    If vTbl(lCnt2) = sTbl Then
                        bInArr = True
                        Exit Do
                    End If
                    lCnt2 = lCnt2 + 1
                Loop While lCnt2 <= UBound(vTbl)

                If Not bInArr Then
                    ReDim Preserve vTbl(0 To (UBound(vTbl) + 1))
                    vTbl(UBound(vTbl)) = sTbl
                End If
            End If
        End If
    Next lCnt

    'return the array
    BuildTableArray = vTbl
End Function

Private Function DoAdd() As Long
    Dim sql As String
    Dim sTable As String
    Dim sFld As String
    Dim sVal As String

    'get the name of the table
    sTable = Me.Fields(0).Table
    'gotta have a table to do the add
    If Len(sTable) = 0 Then Exit Function
    'get the field and value lists for the insert statement
    GetAddString sFld, sVal
    'gotta have a field and value list
    If Len(sFld) = 0 Or Len(sVal) = 0 Then Exit Function

    sql = "INSERT INTO " & sTable & " (" & sFld & ") VALUES(" & sVal & ")"
    'Debug.Print sql
    MyCONN.Execute sql, DoAdd
End Function

Private Function DoUpdate() As Long
    Dim sql As String
    Dim sTable As String
    Dim sUpd As String
    Dim sWhere As String

    'get the name of the table
    sTable = Me.Fields(0).Table
    'gotta have a table to do the add
    If Len(sTable) = 0 Then Exit Function
    'get the field list for the update statement
    GetUpdateString sUpd
    'gotta have an update list
    If Len(sUpd) = 0 Then Exit Function
    'make the where clause for this row
    BuildWhere sWhere
    'gotta have a where clause
    If Len(sWhere) = 0 Then Exit Function

    sql = "UPDATE " & sTable & " SET " & sUpd & " WHERE " & sWhere
    '    Debug.Print sql
    MyCONN.Execute sql, DoUpdate
End Function

Private Function DoDelete() As Long
    Dim sql As String
    Dim sTable As String
    Dim sWhere As String

    'get the name of the table
    sTable = Me.Fields(0).Table
    'gotta have a table to do the add
    If Len(sTable) = 0 Then Exit Function
    'make the where clause for this row
    BuildWhere sWhere
    'gotta have a where clause
    If Len(sWhere) = 0 Then Exit Function

    sql = "DELETE FROM " & sTable & " WHERE " & sWhere
    'Debug.Print sql
    MyCONN.Execute sql, DoDelete
End Function

Private Sub GetAddString(ByRef sFields As String, ByRef sValues As String)
    Dim lCnt As Long
    Dim uFld As New MYSQL_UPDATE_FIELD

    If Me.MyAddObject.Count < 1 Then Exit Sub              'no fields were "let" for this add ...

    'build the field and value lists based on the fields that were "let"
    For lCnt = 1 To Me.MyAddObject.Item(Me.MyAddObject.Count).Count
        Set uFld = Me.MyAddObject.Item(Me.MyAddObject.Count).Item(lCnt)
        sFields = sFields & IIf(Len(sFields) > 0, ",", "") & uFld.Name
        sValues = sValues & IIf(Len(sValues) > 0, ",", "") & DbTick(uFld.Value, Fields(uFld.Name).MySqlType)
    Next lCnt
End Sub

Private Sub GetUpdateString(ByRef sUpdate As String)
    Dim lCnt As Long
    Dim uFld As New MYSQL_UPDATE_FIELD

    If Me.MyUpdateObject.Count < 1 Then Exit Sub           'no fields were "let" for this update ...

    'build the field list based on the fields that were "let" and haven't been updated yet ...
    For lCnt = 1 To Me.MyUpdateObject.Item(mlUpdateRecord, True).Count
        Set uFld = Me.MyUpdateObject.Item(mlUpdateRecord, True).Item(lCnt)
        If Not uFld.Updated Then
            sUpdate = sUpdate _
                    & IIf(Len(sUpdate) > 0, ",", "") _
                    & uFld.Name & "=" & DbTick(uFld.Value, MyFIELDS.MySqlType)
        End If
    Next lCnt
End Sub

Private Sub BuildWhere(ByRef sWhere As String)
    Dim lCnt As Long

    'build the where clause for each primary key field in the current record
    For lCnt = 0 To (FieldCount - 1)
        If (Fields(lCnt).Flags And MY_PRI_KEY_FLAG) Then
            sWhere = sWhere & IIf(Len(sWhere) > 0, " AND ", "") _
                    & Fields(lCnt).Name & IIf(IsNull(Fields(lCnt).Value), " IS ", "=") & DbTick(Fields(lCnt).Value, Fields(lCnt).MySqlType)
        End If
    Next lCnt
End Sub

Private Sub SetUpdatedFlags()
    Dim lCnt As Long
    Dim uFld As New MYSQL_UPDATE_FIELD

    If Me.MyUpdateObject.Count < 1 Then Exit Sub           'no fields were "let" for this update ...

    'set updated = true for each update field
    For lCnt = 1 To Me.MyUpdateObject.Item(mlUpdateRecord, True).Count
        Me.MyUpdateObject.Item(mlUpdateRecord, True).Item(lCnt).Updated = True
    Next lCnt
End Sub

Private Function DbTick(ByVal v As Variant, fType As MYSQL_FIELD_TYPES) As String
    Dim sFormat As String

    If IsNull(v) Then
        DbTick = "NULL"
        Exit Function
    End If

    Select Case fType
        Case MYSQL_DATE:
            sFormat = "YYYY-MM-DD"
            DbTick = "0000-00-00"
        Case MYSQL_TIME:
            sFormat = "HH:MM:SS"
            DbTick = "00:00:00"
        Case MYSQL_DATETIME:
            sFormat = "YYYY-MM-DD HH:MM:SS"
            DbTick = "0000-00-00 00:00:00"
    End Select

    If Len(sFormat) > 0 Then
        If IsDate(v) Then DbTick = Format(v, sFormat)
    Else
        DbTick = Replace(v, "\", "\\", , , vbTextCompare)  'this must be the first Replace
        DbTick = Replace(DbTick, Chr(0), "\0", , , vbTextCompare)
        DbTick = Replace(DbTick, Chr(8), "\b", , , vbTextCompare)
        DbTick = Replace(DbTick, Chr(9), "\t", , , vbTextCompare)
        DbTick = Replace(DbTick, "'", "\'", , , vbTextCompare)
        DbTick = Replace(DbTick, """", "\""", , , vbTextCompare)
        DbTick = Replace(DbTick, Chr(10), "\n", , , vbTextCompare)
        DbTick = Replace(DbTick, Chr(13), "\r", , , vbTextCompare)
    End If

    DbTick = "'" & DbTick & "'"
End Function

'===============================================================================================
'===============================================================================================
'===============================================================================================

