<%Option Explicit%>
<!--#include virtual="/Include/Functions.asp"-->
<!--#include virtual="/Include/DebugFunctions.asp"-->
<!--#include virtual="/Include/OpenDBConnect.asp"-->
<!--#include virtual="/Include/BizObj.asp"-->
<%
Response.ContentType = "text/plain;charset=windows-1251"
ExpirePage()
%>

<table border='0'>
   <tr>
      <td>File</td>
      <td><input type='file' name='ImageFile' maxlength='50'>
   </tr>
   <tr>
      <td valign='top'>Description</td>
      <td><textarea name="ImageDesc" cols='30' rows='5'></textarea>
   </tr>
   <tr>
      <td>Order</td>
      <td><input type='text' name='ImageSortValue' maxlength='50'>
   </tr>
   <tr>
      <td>Featured</td>
      <td><input type='checkbox' name='ImageFeatured' checked></td>
   </tr>
   <tr>
      <td>Enabled</td>
      <td><input type='checkbox' name='ImageEnabled' checked></td>
   </tr>
   <tr>
      <td>&nbsp;</td>
      <td><input type='button' name='btnImgAddSave' value=' Add ' onClick='OnAddImageSave()'></td>
   </tr>
</table>

<!--#include virtual="/Include/CloseDBConnect.asp"-->
