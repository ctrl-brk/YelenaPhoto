<%Option Explicit%>
<!--#include virtual="/Include/Functions.asp"-->
<!--#include virtual="/Include/OpenDBConnect.asp"-->
<%
dim sSQL

sSQL = "exec PutError @Type='Client'" &_
                    ",@ClientIP=" & ToSQLString( Request.ServerVariables( "REMOTE_ADDR" ) ) &_
                    ",@Source=" & ToSQLString( RequestString( "URL" ) ) &_
                    ",@Line=" & ToSQLString( RequestString( "Line" ) ) &_
                    ",@Description=" & ToSQLString( RequestString( "Msg" ) ) &_
                    ",@Comments=" & ToSQLString( RequestString( "Cmt" ) )

obConn.Execute( DebugSQL( sSQL ) )
ClearDebug()

%>
<!--#include virtual="/Include/CloseDBConnect.asp"-->
