<%Option Explicit%>
<!--#include virtual="/Include/Functions.asp"-->
<!--#include virtual="/Include/DebugFunctions.asp"-->
<!--#include virtual="/Include/OpenDBConnect.asp"-->
<!--#include virtual="/Include/BizObj.asp"-->
<%
Response.ContentType = "text/plain;charset=windows-1251"
ExpirePage()

dim obRec, nId, sSelect, sColor, i

sColor = "#352D2B"
nId = RequestString( "iid" )
sSelect = RequestString( "sel" )
set obRec = ExecuteBlob( "exec AdminGetGalleryImages @GalleryId=" & ToSQL( RequestString( "id" ) ) )

i = 0

if ( not obRec.EOF ) then %>
   <hr>
   <table cellpadding='0' cellspacing='0' style='margin-top:5px; padding-top:5px;'>
      <tr>
      <% while( not obRec.EOF )
         if ( sSelect = "Y" ) then
            if ( CInt( obRec( "ImageId" ) ) = CInt( nId ) ) then sColor = "#EBD3AA" else sColor = "#352D2B" end if
         end if %>
         <td id='tdimg_<%=obRec( "ImageId" )%>' style='background-color:black; width:100px; height:100px; text-align:center; border:2px solid <%=sColor%>'>
            <a href='javascript:OnImageSelect(<%=obRec( "ImageId" ) %>)'><img id='img_<%=obRec( "ImageId" )%>' src='<%=ThumbnailImage( obRec( "RelativePath" ), obRec( "ImageName" ) )%>' border='0'></a>
         </td>
         <% if ( i > 9 and i Mod 10 = 0 ) then %></tr><tr><% end if %>
         <% i = i + 1
         obRec.MoveNext()
      wend %>
      </tr>
   </table>
<% end if %>

<% obRec.Close() : set obRec = nothing %>

<!--#include virtual="/Include/CloseDBConnect.asp"-->
