<%Option Explicit%>
<!--#include virtual="/Include/Functions.asp"-->
<!--#include virtual="/Include/DebugFunctions.asp"-->
<!--#include virtual="/Include/OpenDBConnect.asp"-->
<!--#include virtual="/Include/BizObj.asp"-->
<%
Response.ContentType = "text/plain;charset=windows-1251"
ExpirePage()

dim nParentId

nParentId = 0

if ( RequestString( "id" ) <> "" ) then nParentId = CLng( RequestString( "id" ) )

%>
<table border='0' style='margin-top:10px'>
   <tr>
      <td>Image</td>
      <td><input type='file' name='GalleryImage' maxlength='50'>&nbsp;&nbsp;<input type='checkbox' name='GalleryImageResize'> resize
   </tr>
   <tr>
      <td>Name</td>
      <td><input type='text' name='GalleryName' maxlength='50'>
   </tr>
   <tr>
      <td valign='top'>Description</td>
      <td><textarea name='GalleryDesc' cols='80' rows='5'></textarea>
   </tr>
   <tr>
      <td>Path</td>
      <td><input type='text' name='GalleryPath' value='/Galleries/' maxlength='100'>
   </tr>
   <tr>
      <td>Order</td>
      <td><input type='text' name='GallerySortValue' maxlength='50'>
   </tr>
   <tr>
      <td>Enabled</td>
      <td><input type='checkbox' name='GalleryEnabled' checked></td>
   </tr>
   <tr>
      <td>&nbsp;</td>
      <td><input type='button' value='Save' onClick='<% if ( nParentId = 0 ) then %>OnGallerySave()<% else %>OnSubGallerySave(<%=nParentId%>)<% end if %>'></td>
   </tr>

</table>

<!--#include virtual="/Include/CloseDBConnect.asp"-->
