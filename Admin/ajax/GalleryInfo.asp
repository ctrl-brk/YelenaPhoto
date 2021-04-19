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
if ( Gallery.Exists ) then %>
   <table border='0' style='margin-top:10px'>
      <tr>
         <td>Image</td>
         <td><input type='file' name='GalleryImage' maxlength='50'>&nbsp;&nbsp;<input type='checkbox' name='GalleryImageResize'> resize<br><i>(only select a file if you want to replace current one)</i>
      </tr>
      <tr>
         <td>Name</td>
         <td><input type='text' name='GalleryName' maxlength='50' value='<%=Gallery.Name %>'>
      </tr>
      <tr>
         <td valign='top'>Description</td>
         <td><textarea name='GalleryDesc' cols='80' rows='5'><%=Gallery( "Description" ) %></textarea>
      </tr>
      <tr>
         <td>Path</td>
         <td><input type='text' name='GalleryPath' maxlength='100' value='<%=Gallery( "RelativePath" ) %>'>
      </tr>
      <tr>
         <td>Order</td>
         <td><input type='text' name='GallerySortValue' maxlength='50' value='<%=Gallery( "SortValue" ) %>'>
      </tr>
      <tr>
         <td>Enabled</td>
         <td><input type='checkbox' name='GalleryEnabled'<% if ( Gallery.Active ) then %> checked<% end if %>></td>
      </tr>
      <tr>
         <td>&nbsp;</td>
         <td><input type='button' value='Save' onClick='OnGallerySave(<%=Gallery.Id%>)'>&nbsp;&nbsp;&nbsp;<input type='button' value='Delete' onClick='OnGalleryDelete(<%=Gallery.Id%>)'></td>
      </tr>
   </table>
<% end if

set Gallery = nothing
%>
<!--#include virtual="/Include/CloseDBConnect.asp"-->
