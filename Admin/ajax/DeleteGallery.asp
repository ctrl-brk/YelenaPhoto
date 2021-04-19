<%Option Explicit%>
<!--#include virtual="/Include/Functions.asp"-->
<!--#include virtual="/Include/DebugFunctions.asp"-->
<!--#include virtual="/Include/OpenDBConnect.asp"-->
<!--#include virtual="/Include/BizObj.asp"-->
<%
Response.ContentType = "text/plain;charset=windows-1251"
ExpirePage()

dim Gallery

set Gallery = new CGallery
Gallery.Init( RequestString( "id" ) )
Gallery.Delete()
set Gallery = nothing
Response.Write( "Gallery deleted" ) %>

<!--#include virtual="/Include/CloseDBConnect.asp"-->
