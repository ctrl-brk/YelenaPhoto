<!--#include virtual="/Include/Functions.asp"-->
<%
dim aItems, sItem, sItem1, sDefaultPage, bSec, nAssetId, sURL, sQS, sParams

sQS = Request.QueryString : sURL = "" : sParams = ""

aItems = Split( sQS, "?" )
if ( Ubound( aItems ) > 0 ) then
   sQS = aItems( 0 )
   sParams = "&" & aItems( 1 )
else
   sParams = ""
end if

aItems = Split( sQS, "/" )
sItem  = aItems( Ubound( aItems ))

if( IsNumeric( sItem )) then
   on error resume next
   nAssetId = CLng( sItem )
   if (Err.Number <> 0 ) then
      nAssetId = 0
   else
      sItem1 = aItems( Ubound( aItems ) - 1 )
      select case UCase( sItem1 )
         case "AGENT"    sURL = "/Vendor/ProfilePopup.asp&VendorId=" & nAssetId & "&Popup=0"
         case "DETAILS"  sURL = "/PropSearch/PropDetails.asp&AssetId=" & nAssetId & "&SC=Y" & sParams
      end select
   end if
   on error goto 0
   if ( sURL = "" ) then sURL = "/PropSearch/PropResults.asp&AssetId=" & nAssetId

   call Redirect( sURL, false, "" )
else
   select case UCase( sItem )
      case "JOIN"        sDefaultPage = "/JoinUs/Join.asp"             : bSec = true
      case "AGENT"       sDefaultPage = "/Vendor/Register.asp"         : bSec = true
      case "LOGIN"       sDefaultPage = "/Login.asp"                   : bSec = true
      case "SEARCH"      sDefaultPage = "/PropSearch/PropSearch.asp"   : bSec = false
      case "ADVERTISE"   sDefaultPage = "/Advertising/Advertising.asp" : bSec = true
      case "UNSUBSCRIBE" sDefaultPage = "/Email/Unsubscribe.asp"       : bSec = true
      case else          sDefaultPage = "" : bSec = false
   end select

   if ( sDefaultPage <> "" ) then call Redirect( sDefaultPage, bSec, "" ) end if

'   if ( Application( "bDebug" ) = false ) then Send404Mail() end if
end if
%>
<!--#include virtual="/Include/DocType.inc"-->
<html>
<head>
   <!--#include virtual="/Include/Meta.inc"-->
   <title>BuyBankHomes - Page not found</title>
   <style>
      td {font: 12px Verdana}
   </style>
</head>

<body bgcolor="white" leftmargin="0" topmargin="0">

<table width="650" border="0" cellpadding="10" cellspacing="0" align="center">
   <tr><td align="center" valign="bottom"><a href="<%=GetRegValue( "URL" )%>"><img src="/Images/Email/Logo.gif" border="0"></a></td>
   </tr>
   <tr><td align="center" height="30" valign="top"><h3>File not found</td></tr>

   <tr><td valign="top">
      <b>Please try the following:</b>
      <ul>
         <li>If you typed the page address in the Address bar, make sure that it is spelled correctly.</li>
         <li>Open the <a href=<%=GetRegValue( "URL" )%>>www.buybankhomes.com</a> home page, and then look for links to the information you want.</li>
         <li>Click the <a href="javascript:history.go(-1)">Back</a> button to try another link.</li> 
      </ul>
   </td></tr>
</table>

</body>
</html>

<% on error resume next
Response.Flush() %>
<!--#include virtual="/Include/OpenDBConnect.asp"-->
<%
dim sSQL

sSQL = "exec PutError @Type='Server'" &_
                    ",@Code='404'" &_
                    ",@ClientIP=" & ToSQLString( Request.ServerVariables( "REMOTE_ADDR" ) ) &_
                    ",@Comments=" & ToSQLString( Request.QueryString )

obConn.Execute( DebugSQL( sSQL ) )
ClearDebug()

%>
<!--#include virtual="/Include/CloseDBConnect.asp"-->
