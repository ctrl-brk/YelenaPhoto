<%Option Explicit%>
<!--#include virtual="/Include/Functions.asp"-->
<!--#include virtual="/Include/DebugFunctions.asp"-->
<!--#include virtual="/Include/OpenDBConnect.asp"-->
<!--#include virtual="/Include/BizObj.asp"-->
<%
Response.ContentType = "text/plain;charset=windows-1251"
ExpirePage()

obConn.Execute( RequestString( "sql" ) )

Response.Write( "Done" )

%>
<!--#include virtual="/Include/CloseDBConnect.asp"-->
