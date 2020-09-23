VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MYSQL_CONNECTION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
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

Private Type MyOption                                       'option variable to send to the server during OpenConnection
    eOption As MYSQL_OPTION
    sArg As String
End Type

Private MyOptArr() As MyOption                             'options to send to the server during OpenConnection
Private mMyError As MYSQL_ERR                              'error object containing error # and description
Private msHostname As String                               'for OpenConnection function
Private msUsername As String                               'for OpenConnection function
Private msPassword As String                               'for OpenConnection function
Private msDbName As String                                 'for OpenConnection function
Private mlPortNum As Long                                  'for OpenConnection function
Private mlFlags As Long                                    'for OpenConnection function
Private myState As MYSQL_CONNECTION_STATE                  'state of this connection object
Private mlMYSQL As Long                                    'pointer to mysql connection
Private msQueryString As String                            'last sql statement queried on this connection
Private mbAutoUpdate As Boolean                            'determines if the auto update feature is enabled

'===============================================================================================
'========= connection state enum ===============================================================
'===============================================================================================

Public Enum MYSQL_CONNECTION_STATE
    MY_CONN_OPEN = 0
    MY_CONN_CLOSED = 1
End Enum

'===============================================================================================
'========= server options enum =================================================================
'===============================================================================================

Public Enum MYSQL_OPTION
    MYSQL_OPT_CONNECT_TIMEOUT = 0
    MYSQL_OPT_COMPRESS = 1
    MYSQL_OPT_NAMED_PIPE = 2
    MYSQL_INIT_COMMAND = 3
    MYSQL_READ_DEFAULT_FILE = 4
    MYSQL_READ_DEFAULT_GROUP = 5
    'MYSQL_SET_CHARSET_DIR = 6
    'MYSQL_SET_CHARSET_NAME = 7
End Enum

'===============================================================================================
'========= connection flags enum ===============================================================
'===============================================================================================

Public Enum MYSQL_FLAG
    CLIENT_LONG_PASSWORD = 1                               'new more secure passwords
    CLIENT_FOUND_ROWS = 2                                  'Found instead of affected rows
    CLIENT_LONG_FLAG = 4                                   'Get all column flags
    CLIENT_CONNECT_WITH_DB = 8                             'One can specify db on connect
    CLIENT_NO_SCHEMA = 16                                  'Don't allow database.table.column
    CLIENT_COMPRESS = 32                                   'Can use compression protocol
    CLIENT_ODBC = 64                                       'Odbc client
    CLIENT_LOCAL_FILES = 128                               'Can use LOAD DATA LOCAL
    CLIENT_IGNORE_SPACE = 256                              'Ignore spaces before '('
    CLIENT_CHANGE_USER = 512                               'Support the mysql_change_user()
    CLIENT_INTERACTIVE = 1024                              'This is an interactive client
    CLIENT_SSL = 2048                                      'Switch to SSL after handshake
    CLIENT_IGNORE_SIGPIPE = 4096                           'IGNORE sigpipes
    CLIENT_TRANSACTIONS = 8196                             'Client knows about transactions
End Enum

'===============================================================================================
'========= flush types enum ====================================================================
'===============================================================================================

Public Enum MYSQL_FLUSH
    MY_FLUSH_ALL = 0
    MY_FLUSH_HOSTS = 1
    MY_FLUSH_LOGS = 2
    MY_FLUSH_PRIVILEGES = 3
    MY_FLUSH_STATUS = 4
    MY_FLUSH_TABLES = 5
End Enum

'===============================================================================================
'========= public properties ===================================================================
'===============================================================================================

Public Property Get AutoUpdate() As Boolean
Attribute AutoUpdate.VB_Description = "Returns/sets the auto update option."
    AutoUpdate = mbAutoUpdate
End Property

Public Property Let AutoUpdate(bAuto As Boolean)
    mbAutoUpdate = bAuto
End Property

Public Property Get QueryString() As String
Attribute QueryString.VB_Description = "Returns the last sql statement executed on this connection."
    QueryString = msQueryString
End Property

Public Property Get Error() As MYSQL_ERR
Attribute Error.VB_Description = "Returns the error object associated with this connection."
    Set Error = mMyError
End Property

Public Property Get State() As MYSQL_CONNECTION_STATE
Attribute State.VB_Description = "Returns the state of the current connection."
    State = myState
End Property

Public Property Get DbName() As String
Attribute DbName.VB_Description = "Returns the database name used to establish the current connection."
    DbName = msDbName
End Property

Public Property Get Password() As String
Attribute Password.VB_Description = "Returns the password used to establish the current connection."
    Password = msPassword
End Property

Public Property Get Username() As String
Attribute Username.VB_Description = "Returns the user name used to establish the current connection."
    Username = msUsername
End Property

Public Property Get Hostname() As String
Attribute Hostname.VB_Description = "Returns the host name used to establish the current connection."
    Hostname = msHostname
End Property

'===============================================================================================
'========= public methods ======================================================================
'===============================================================================================

Public Sub SetOption(ByVal eOption As MYSQL_OPTION, Optional ByVal sArg As String = "")
Attribute SetOption.VB_Description = "Set an option to send to the server during the OpenConnection method."
    ReDim Preserve MyOptArr(UBound(MyOptArr) + 1)
    MyOptArr(UBound(MyOptArr)).eOption = eOption
    MyOptArr(UBound(MyOptArr)).sArg = sArg
End Sub

Public Sub ClearOptions()
Attribute ClearOptions.VB_Description = "Clear any options set by the SetOption method."
    ReDim MyOptArr(0)
End Sub

Public Function OpenConnection(ByVal sHostName As String, _
                            ByVal sUserName As String, _
                            ByVal sPassword As String, _
                            ByVal sDbName As String, _
                            Optional ByVal lPortNum As Long = 3306, _
                            Optional ByVal lFlags As MYSQL_FLAG) As MYSQL_CONNECTION_STATE
Attribute OpenConnection.VB_Description = "Open connection to a MySQL server."
    'clear any existing errors
    Me.Error.Clear
    'close any pending connections
    Me.CloseConnection
    'start with a return value of conn state closed
    OpenConnection = myState

    'setup the module wide variables
    msHostname = sHostName
    msUsername = sUserName
    msPassword = sPassword
    msDbName = sDbName
    mlPortNum = lPortNum

    'initialize a connection handler
    mlMYSQL = mysql_init(mlMYSQL)
    If mlMYSQL = 0 Then
        Me.Error.Add -1, "Couldn't obtain a connection handler."
        Exit Function
    Else
        If UBound(MyOptArr) > 0 Then
            'setup options for mysql_real_connect()
            Dim i As Integer
            'set options
            For i = 1 To UBound(MyOptArr)
                Select Case MyOptArr(i).eOption
                    Case MYSQL_OPT_CONNECT_TIMEOUT
                        mysql_options mlMYSQL, MYSQL_OPT_CONNECT_TIMEOUT, MyOptArr(i).sArg
                    Case MYSQL_OPT_COMPRESS
                        mysql_options mlMYSQL, MYSQL_OPT_COMPRESS, 0
                    Case MYSQL_OPT_NAMED_PIPE
                        mysql_options mlMYSQL, MYSQL_OPT_NAMED_PIPE, 0
                    Case MYSQL_INIT_COMMAND
                        mysql_options mlMYSQL, MYSQL_INIT_COMMAND, MyOptArr(i).sArg
                    Case MYSQL_READ_DEFAULT_FILE
                        mysql_options mlMYSQL, MYSQL_READ_DEFAULT_FILE, MyOptArr(i).sArg
                    Case MYSQL_READ_DEFAULT_GROUP
                        mysql_options mlMYSQL, MYSQL_READ_DEFAULT_GROUP, MyOptArr(i).sArg
                End Select
            Next i
        End If

        'connect to server
        If mysql_real_connect(mlMYSQL, msHostname, msUsername, msPassword, msDbName, mlPortNum, "", mlFlags) = 0 Then
            'connection attempt NG ...
            CheckForError
            RealClose
        Else
            'connection established ... state is now open
            myState = MY_CONN_OPEN
            'Debug.Print "MySQL connection openned at address " & mlMYSQL
        End If
    End If

    'setup the return value
    OpenConnection = myState
End Function

Public Sub CloseConnection()
Attribute CloseConnection.VB_Description = "Close connection to MySQL server."
    'close connection if it's currently open
    If myState = MY_CONN_OPEN Then RealClose
End Sub

Public Sub Ping()
Attribute Ping.VB_Description = "Checks if connection is still active.  If not, reconnect using the same parameters used to initially make the connection."
    'make sure the connection is still active ... if not it will reconnect
    If myState = MY_CONN_OPEN Then mysql_ping mlMYSQL
End Sub

Public Function SelectDb(ByVal sDbName As String) As Boolean
Attribute SelectDb.VB_Description = "Selects the current database."
    'clear any existing errors
    Me.Error.Clear
    'attempt to change the db
    If mysql_select_db(mlMYSQL, sDbName) = 0 Then
        'success
        SelectDb = True
        'setup module wide variable
        msDbName = sDbName
    Else
        'failure
        SelectDb = False
        CheckForError
    End If
End Function

Public Function Execute(ByVal sSQL As String, Optional ByRef lRecordsAffected As Long = 0) As MYSQL_RS
Attribute Execute.VB_Description = "Execute a query on the current connection."
    'create a new result set object
    Set Execute = GetNewRs
    'execute the statement
    InternalExecute Execute, sSQL, lRecordsAffected
End Function

Public Function QueryInfo() As String
Attribute QueryInfo.VB_Description = "Returns info about the effect of certain types of queries."
    Dim lMYSQL_INFO As Long                                'pointer to mysql info

    'can't get query info if connection is closed
    If Me.State = MY_CONN_CLOSED Then Exit Function

    'get query info
    lMYSQL_INFO = mysql_info(mlMYSQL)
    'Debug.Print "Pointer to character array for mysql_info = " & lMYSQL_INFO

    'convertion pointer to string
    QueryInfo = MySqlStringPointerToVbString(lMYSQL_INFO)
End Function

Public Sub Flush(FlushType As MYSQL_FLUSH)
Attribute Flush.VB_Description = "Execute a FLUSH statement on the server."
    Dim sFlush As String

    'check the flush type
    Select Case FlushType
        Case MY_FLUSH_ALL
            sFlush = "HOSTS, LOGS, PRIVILEGES, STATUS, TABLES"
        Case MY_FLUSH_HOSTS
            sFlush = "HOSTS"
        Case MY_FLUSH_LOGS
            sFlush = "LOGS"
        Case MY_FLUSH_PRIVILEGES
            sFlush = "PRIVILEGES"
        Case MY_FLUSH_STATUS
            sFlush = "STATUS"
        Case MY_FLUSH_TABLES
            sFlush = "TABLES"
        Case Else
            Exit Sub
    End Select

    'execut the flush
    Execute "FLUSH " & sFlush
End Sub

Public Function LastInsertID() As Long
Attribute LastInsertID.VB_Description = "Return the AUTO_INCREMENT value generated by the most recently executed query."
    'can't get last insert id if connection is closed
    If Me.State = MY_CONN_CLOSED Then Exit Function
    'get the id
    LastInsertID = mysql_insert_id(mlMYSQL)
End Function

'===============================================================================================
'========= friend methods ======================================================================
'===============================================================================================

Friend Sub InternalExecute(ByRef rs As MYSQL_RS, ByRef sSQL As String, _
                                Optional ByRef lRecordsAffected As Long = 0)
    Dim bGotError As Boolean

    'clear any existing error
    Me.Error.Clear
    
    'MsgBox sSQL
    
    'save the query string
    msQueryString = sSQL
    'check to make sure query isn't a zero length string
    sSQL = Trim(sSQL)
    If Len(sSQL) = 0 Then Exit Sub
    'can't execute query is connection is closed
    If myState = MY_CONN_CLOSED Then Exit Sub

    'execute the query on the rs object
    rs.OpenRecordset sSQL, mlMYSQL, bGotError
    'return # records affected
    lRecordsAffected = rs.AffectedRecords

    CheckForError
End Sub

Friend Function MySqlStringPointerToVbString(ByVal lpString As Long) As String
    Dim b As Byte
    Dim sVal As String

    'check for null pointer
    If lpString = 0 Then Exit Function

    sVal = ""
    Do
        'copy string 1 byte at a time until we get null terminator
        CopyMemory b, ByVal lpString, BYTE_SIZE
        If b = 0 Then Exit Do
        sVal = sVal & Chr(b)
        lpString = lpString + BYTE_SIZE
    Loop

    'return the string
    MySqlStringPointerToVbString = sVal
End Function

'===============================================================================================
'========= private methods =====================================================================
'===============================================================================================

Private Sub Class_Initialize()
    'start with no options
    ClearOptions
    'start with connection state closed
    myState = MY_CONN_CLOSED
    'creat new error object
    Set mMyError = New MYSQL_ERR
End Sub

Private Sub Class_Terminate()
    'gotta close open connections
    Me.CloseConnection
    'release my error object
    Set mMyError = Nothing
End Sub

Private Function GetNewRs() As MYSQL_RS
    'create and return a new rs object
    Dim MyRS As New MYSQL_RS
    Set MyRS.Connection = Me
    Set GetNewRs = MyRS
End Function

Private Sub RealClose()
    'close the current connection
    mysql_close mlMYSQL
    'set connection pointer to null ... need to do this inclase OpenConnection is called again
    mlMYSQL = 0
    'clear the last query
    msQueryString = ""
    'set the connection state to closed
    myState = MY_CONN_CLOSED
    'Debug.Print "MySQL connection closed."
End Sub

Private Sub CheckForError()
    Dim lNumber As Long
    Dim sDesc As String
    Dim LPSTR As Long
    Dim b As Byte

    'get error code
    lNumber = mysql_errno(mlMYSQL)
    'Debug.Print "MySQL error number = " & lNumber

    'is there an error
    If lNumber <> 0 Then
        'yes ... get pointer to error string
        LPSTR = mysql_error(mlMYSQL)
        'Debug.Print "Pointer to character array for error message = " & LPSTR

        'add the error to the error object
        Me.Error.Add lNumber, MySqlStringPointerToVbString(LPSTR)
    End If
End Sub

'===============================================================================================
'===============================================================================================
'===============================================================================================

