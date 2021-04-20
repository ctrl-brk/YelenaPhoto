<%Option Explicit%>
<%
OpenDBConnect()

dim obRec, sImg

set obRec = ExecuteBlob( DebugSQL( "exec GetGalleries" ) )
ClearDebug()

gsSubmenu = ""
while ( not obRec.EOF )
   gsSubmenu = gsSubmenu & "<a href='/Gallery.asp?id=" & obRec( "GalleryId" ) & "'>&raquo; " & obRec( "GalleryName" ) & "</a>"
   obRec.MoveNext()
wend
obRec.MoveFirst()
%>
<!--#include virtual="/Framework/Open.asp"-->

<script language='JavaScript'>

function OnImgOver( img, bOver )
{
   if ( IsIE() ) img.filters(0).enabled = !bOver;
}
</script>

<style>
.gtbl1 td {padding-bottom:5px;}
.gtbl1 td img {border:0; filter:progid:DXImageTransform.Microsoft.Alpha(opacity=100,finishopacity=0,style=2)}
.gtbl1 td a {text-decoration:none;}
.gtbl1 td.st01 {width:110px;text-align:left; border-right:1px solid #62584e;}
.gtbl1 td.st02 {vertical-align:top;padding-left:10px;}
</style>

<h1>Photo <strong>galleries</strong></h1>
<hr>
<table border='0' cellpadding='0' cellspacing='3' class='gtbl1'>
   <% while ( not obRec.EOF )
      if ( obRec( "Thumbnail" ) <> "" ) then sImg = obRec( "RelativePath" ) & "/" & obRec( "Thumbnail" ) else sImg = "/Galleries/Thumb.gif" end if %>
      <tr>
         <td class='st01'><a href='/Gallery.asp?id=<%=obRec( "GalleryId" )%>'><img src='<%=sImg%>' onmouseover='OnImgOver(this,true)' onmouseout='OnImgOver(this,false)'></a></td>
         <td class='st02'>
            <a href='/Gallery.asp?id=<%=obRec( "GalleryId" )%>'>&raquo; <%=obRec( "GalleryName" )%></a><br><br>
            <%=obRec( "Description" ) %>
         </td>
      </tr>
      <% obRec.MoveNext()
   wend
   
   obRec.Close() : set obRec = nothing %>
<!--   <tr><td colspan='2'>
      <hr><span style='color:red;'>*</span> <i>Please note that this site is best viewed in Internet Explorer 7 or above.</i>
   </td></tr>-->
</table>

<!--#include virtual="/Framework/Close.asp"-->
