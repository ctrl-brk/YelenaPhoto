<%Option Explicit%>
<!--#include virtual="/Include/Functions.asp"-->
<!--#include virtual="/Include/DebugFunctions.asp"-->
<!--#include virtual="/Include/OpenDBConnect.asp"-->
<!--#include virtual="/Include/BizObj.asp"-->
<%
Response.ContentType = "text/plain;charset=windows-1251"
ExpirePage()
%>
<select name='Galleries' size='13' onchange='OnGalleryChange()' style='width:260px'>
<% dim obRec
   set obRec = ExecuteBlob( "exec AdminGetGalleries" )
   while ( not obRec.EOF )
      Response.Write( "<option value='" & obRec( "GalleryId" ) & "'>" & obRec( "GalleryName" ) & "</option>" & vbCrLf )
      obRec.MoveNext()
   wend
   obRec.Close() : set obRec = nothing %>
</select>

<!--#include virtual="/Include/CloseDBConnect.asp"-->
