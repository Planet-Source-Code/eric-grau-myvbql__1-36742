Attribute VB_Name = "modDeclare"
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

Public Const LONG_SIZE = 4
Public Const INT_SIZE = 2
Public Const BYTE_SIZE = 1

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (lpDestination As Any, _
        lpSource As Any, _
        ByVal lLength As Long)

'connection management routines
Public Declare Sub mysql_close Lib "libmySQL" _
        (ByVal lMYSQL As Long)
Public Declare Function mysql_init Lib "libmySQL" _
        (ByVal lMYSQL As Long) As Long
Public Declare Function mysql_options Lib "libmySQL" _
        (ByVal lMYSQL As Long, _
        ByVal lOption As Long, _
        ByVal sArg As String) As Long
Public Declare Function mysql_ping Lib "libmySQL" _
        (ByVal lMYSQL As Long) As Long
Public Declare Function mysql_real_connect Lib "libmySQL" _
        (ByVal lMYSQL As Long, _
        ByVal sHostName As String, _
        ByVal sUserName As String, _
        ByVal sPassword As String, _
        ByVal sDbName As String, _
        ByVal lPortNum As Long, _
        ByVal sSocketName As String, _
        ByVal lFlags As Long) As Long

'status and error-reporting routines
Public Declare Function mysql_errno Lib "libmySQL" _
        (ByVal lMYSQL As Long) As Long
Public Declare Function mysql_error Lib "libmySQL" _
        (ByVal lMYSQL As Long) As Long

'query contruction and execution routines
Public Declare Function mysql_query Lib "libmySQL" _
        (ByVal lMYSQL As Long, _
        ByVal sQueryString As String) As Long
Public Declare Function mysql_select_db Lib "libmySQL" _
        (ByVal lMYSQL As Long, _
        ByVal sDbName As String) As Long

'result set processing routines
Public Declare Function mysql_affected_rows Lib "libmySQL" _
        (ByVal lMYSQL_RES As Long) As Long
Public Declare Sub mysql_data_seek Lib "libmySQL" _
        (ByVal lMYSQL_RES As Long, ByVal lOffset As Currency)
Public Declare Function mysql_fetch_field_direct Lib "libmySQL" _
        (ByVal lMYSQL_RES As Long, ByVal lFieldNum As Long) As Long
Public Declare Function mysql_fetch_lengths Lib "libmySQL" _
        (ByVal lMYSQL_RES As Long) As Long
Public Declare Function mysql_fetch_row Lib "libmySQL" _
        (ByVal lMYSQL_RES As Long) As Long
Public Declare Function mysql_field_count Lib "libmySQL" _
        (ByVal lMYSQL As Long) As Long
Public Declare Sub mysql_free_result Lib "libmySQL" _
        (ByVal lMYSQL As Long)
Public Declare Function mysql_info Lib "libmySQL" _
        (ByVal lMYSQL As Long) As Long
Public Declare Function mysql_insert_id Lib "libmySQL" _
        (ByVal lMYSQL As Long) As Long
Public Declare Function mysql_num_fields Lib "libmySQL" _
        (ByVal lMYSQL_RES As Long) As Long
Public Declare Function mysql_num_rows Lib "libmySQL" _
        (ByVal lMYSQL_RES As Long) As Long
Public Declare Function mysql_store_result Lib "libmySQL" _
        (ByVal lMYSQL As Long) As Long
Public Declare Function mysql_use_result Lib "libmySQL" _
        (ByVal lMYSQL As Long) As Long


