<%Option Explicit%>
<!--#include virtual="/Include/Functions.asp"-->
<%
Response.Clear()

dim sHTML

if ( not Application( "bDebug" ) = true ) then

   dim obError, sDesc, sSrc, sLine
   sDesc = "" : sSrc = "" : sLine = ""

   set obError = Server.GetLastError()
   if ( obError.ASPDescription <> "" ) then sDesc = obError.ASPDescription else sDesc = obError.Description end if
   if ( obError.ASPCode > "" ) then sDesc = sDesc & obError.ASPCode
   sDesc = sDesc & " (0x" & Hex( obError.Number ) & ")"

   sSrc = obError.Source()
  
   if ( obError.File <> "?" ) then sLine = obError.File
   if ( obError.Line > 0 ) then sLine = sLine & ", line "   & obError.Line
   if ( obError.Column > 0 ) then sLine = sLine & ", column " & obError.Column

else

  sHTML = Send500Mail( false )
  Response.Write( sHTML )
  Response.End()

end if
%>

<!--#include virtual="/Include/DocType.inc"-->
<html>
<head>
   <!--#include virtual="/Include/Meta.inc"-->
   <title>YelenaPhoto.com - Server error</title>
   <link rel="stylesheet" href="/CSS/main.css">
</head>

<body>

<div align='center'>
<table border='0' class='header' cellspacing='0' align='center'>
   <tr>
      <td class='logo'><a href='/'><img src='/Images/logo.gif' border='0'></a></td>
      <td class='nav'>&nbsp;</td>
   </tr>
   <tr>
      <td>&nbsp;</td>
      <td class='subnav'>&nbsp;</td>
</table>
</div>

<div align='center'>
<div align='center' class='content'>
   <table border='0' class='tblcnt' cellspacing='0' align='center'>
      <tr>
         <td>
            <h1>Page is <strong>temporarily unavailable</strong></h1>
            <p>
               Sorry for the inconvenience, but we are currently working on this function of the site at this moment.
               Please check back later to complete your current step.<br>
               Thanks for your understanding and patience.
            </p>
            <br>
            <p style='text-align:right'><i>YelenaPhoto.com</i></p>
         </td>
      </tr>
   </table>
</div>
</div>

</body>
</html>

<%
on error resume next
Response.Flush()
Send500Mail( true )
%>