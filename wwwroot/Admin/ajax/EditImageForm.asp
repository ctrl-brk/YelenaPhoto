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
%>

<table border='0'>
   <tr>
      <td valign='top'>File</td>
      <td><input type='file' name='ImageFile' maxlength='50'><br><i>(only select a file if you want to replace<br>current one)</i>
   </tr>
   <tr>
      <td valign='top'>Description</td>
      <td><textarea name="ImageDesc" cols='30' rows='5'><%=Image( "Description" )%></textarea>
   </tr>
   <tr>
      <td>Order</td>
      <td><input type='text' name='ImageSortValue' maxlength='50' value='<%=Image( "SortValue" )%>'>
   </tr>
   <tr>
      <td>Featured</td>
      <td><input type='checkbox' name='ImageFeatured'<% if ( Image.Featured ) then Response.Write( " checked" ) end if %>></td>
   </tr>
   <tr>
      <td>Enabled</td>
      <td><input type='checkbox' name='ImageEnabled'<% if ( Image.ActiveFlag = true ) then Response.Write( " checked" ) end if %>></td>
   </tr>
   <tr>
      <td>&nbsp;</td>
      <td><input type='button' name='btnImgSave' value='Save' onClick='OnImageSave(<%=Image.Id%>)'>&nbsp;&nbsp;&nbsp;<input type='button' name='btnImgDelete' value='Delete' onClick='OnImageDelete(<%=Image.Id%>)'></td>
   </tr>
</table>

<% set Image = nothing %>

<!--#include virtual="/Include/CloseDBConnect.asp"-->
