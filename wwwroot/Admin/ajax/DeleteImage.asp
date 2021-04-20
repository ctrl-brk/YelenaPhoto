<%Option Explicit%>
<!--#include virtual="/Include/Functions.asp"-->
<!--#include virtual="/Include/DebugFunctions.asp"-->
<!--#include virtual="/Include/OpenDBConnect.asp"-->
<!--#include virtual="/Include/BizObj.asp"-->
<%
Response.ContentType = "text/plain;charset=windows-1251"
ExpirePage()

dim Image

set Image = new CImage
Image.Init( RequestString( "id" ) )
Image.Delete()
set Image = nothing
Response.Write( "Image deleted" ) %>

<!--#include virtual="/Include/CloseDBConnect.asp"-->
