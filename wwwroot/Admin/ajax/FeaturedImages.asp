<%Option Explicit%>
<!--#include virtual="/Include/Functions.asp"-->
<!--#include virtual="/Include/DebugFunctions.asp"-->
<!--#include virtual="/Include/OpenDBConnect.asp"-->
<!--#include virtual="/Include/BizObj.asp"-->
<%
Response.ContentType = "text/plain;charset=windows-1251"
ExpirePage()

dim obRec

set obRec = ExecuteBlob( "exec GetFeaturedImages" )

if ( not obRec.EOF ) then %>
<table border='0'>
   <% while( not obRec.EOF ) %>
      <tr>
         <td><a href='javascript:OnImageSelect(<%=obRec( "ImageId" )%>,<%=obRec( "GalleryId" )%>)'><img src='<%=obRec( "RelativePath" ) & "/thmb/" & obRec( "ImageName" )%>' border='0'></a></td>
         <td><%=obRec( "Description" ) %></td>
      </tr>
      <% obRec.MoveNext()
   wend %>
</table>
<% end if

obRec.Close() : set obRec = nothing
%>
<!--#include virtual="/Include/CloseDBConnect.asp"-->
