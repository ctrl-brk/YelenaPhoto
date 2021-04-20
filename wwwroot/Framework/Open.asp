<%
dim gsSubmenu, gbGallery, gbAdmin
%>
<!--#include virtual="/Include/Functions.asp"-->
<!--#include virtual="/Include/DebugFunctions.asp"-->
<!--#include virtual="/Include/DocType.inc"-->
<!--#include virtual="/Include/OpenDBConnect.asp"-->
<!--#include virtual="/Include/BizObj.asp"-->
<!--#include virtual="/Include/UserFunctions.asp"-->
<!--#include virtual="/Include/CreateUser.asp"-->
<% if ( not gcUser.bLoggedIn ) then gcUser.LoginFromCookie() %>
<html>
<head>
   <!--#include virtual="/Include/Meta.inc"-->
   <!--#include virtual="/Include/MetaKeywords.inc"-->
   <link rel="stylesheet" href="/CSS/main.css">
   <title>Yelena Gracheva - Photographer</title>
</head>

<script language="JavaScript" src="/Include/Lib.js" type="text/javascript"></script>
<script language="JavaScript" src="/Include/Form.js" type="text/javascript"></script>
<script language="JavaScript" src="/Include/Ajax.js" type="text/javascript"></script>

<body onload='_OnLoad_()'>

<script language="JavaScript">

var _gbInSubmit = false;

function OnDefaultSubmit()
{
   return( false );
}

function SubmitByEnter( e )
{
   var key;

   if( IsIE() ) key = window.event.keyCode;
   else key = e.which;
   if( key == 13 && typeof( OnSubmit ) != "undefined" ) OnSubmit();
}

function EnableSubmitControl( sCtrl )
{
   GetControl( sCtrl ).onkeypress = SubmitByEnter;
}

function EnableSubmitControls()
{
   var obCtrl = null, nCnt = 0;

   for( var i=0; i < document.forms[0].elements.length; i++ )
      {
      obCtrl = document.forms[0].elements[i];
      if ( obCtrl )
         {
         if ( obCtrl.type == "text" || obCtrl.type == "password" || obCtrl.type == "file" || obCtrl.type == "checkbox" || obCtrl.type == "radio" )
            {
            obCtrl.onkeypress = SubmitByEnter;
            nCnt++;
            }
         }
      }
   return( nCnt );
}

function _OnLoad_()
{
  if ( typeof( OnLoad ) != "undefined" ) OnLoad();
}

function SubmitForm( sURL, sDisableCtrls )
{
   if ( !_gbInSubmit )
      {
      var aCtrls;

      _gbInSubmit = true;
      if ( sDisableCtrls )
         {
         aCtrls = sDisableCtrls.split( "," );
         for ( var i = 0; i<aCtrls.length; i++ ) GetControl( aCtrls[i] ).disabled = true;
         }
      document.F0.action = sURL;
      document.F0.submit();
      }
}

</script>

<form name="F0" method="post" onSubmit="return(OnDefaultSubmit())">

<div align='center'>
<table border='0' class='header' cellspacing='0' align='center'>
   <tr><td class='logo' colspan='4'><a href='/'></a></td></tr>
   <tr>
      <td class='nav' style='width:412px; text-align:right'><a href='/Galleries.asp' class='gal'></a></td>
      <td class='nav' style='width:158px; text-align:center'><a href='/Testimonials.asp' class='tst'></a></td>
      <td class='nav' style='width:68px; text-align:left'><a href='/Store.asp' class='str'></a></td>
      <td class='nav' style='text-align:left;'><a href='/About.asp' class='abt'></a></td>
   </tr>
   <tr>
      <td class='subnav' colspan='4'>
         <% if ( gsSubmenu <> "" ) then %>
            <div id='subnav'><%=gsSubmenu%></div>
         <% else %>
            &nbsp;
         <% end if %>
      </td>
</table>
</div>

<div align='center'>
<div align='center' class='<%
   if ( gbGallery = true ) then
      Response.Write( "g" )
   elseif ( gbAdmin = true ) then
      Response.Write( "a" )
   end if %>content'>
<% if ( gbGallery <> true and gbAdmin <> true ) then %>
   <table border='0' class='tblcnt' cellspacing='0' align='center'>
      <tr>
         <td>
<% end if %>
