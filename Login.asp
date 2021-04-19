<%Option Explicit%>
<!--#include virtual="/Framework/Open.asp"-->
<%
dim sRedir, sEmail, sPass

if ( gcUser.bLoggedIn ) then sRedir = gcUser.HomePage

if ( sRedir <> "" ) then
   if ( RequestString( "goto" ) <> "" ) then sRedir = RequestString( "goto" ) & "?" & RestoreRequestString() end if
   call Redirect( sRedir )
end if
%>
<script language='JavaScript'>
function OnLoad()
{
   <% if ( RequestString( "LoginEmail" ) <> "" ) then %>SetControlValue( "LoginEmail", "<%=RequestString( "LoginEmail" )%>" );<% end if %>
   ValidateRequiredControl( "LoginEmail" );
   ValidateRequiredControl( "LoginPassword" );
   ValidateEmailControl( "LoginEmail", true );
   EnableSubmitControls();
   SetControlFocus( "LoginEmail" );
}

function OnSubmit()
{
   isValidEmailCtrl( "LoginEmail", true, "Email" );
   GetControlValue( "LoginPassword", true, "Password" );

   if ( !IsValidationOK() )
      {
      SetValidationFocus();
      FinalValidationWarning();
      }
   else SubmitForm( "Login.asp?<%=RestoreRequestString()%>" );
}
</script>
<%
if ( RequestString( "LoginEmail" ) <> "" and RequestString( "LoginPassword" ) <> "" ) then
   sEmail = RequestString( "LoginEmail" )
   sPass = RequestString( "LoginPassword" )
elseif( RequestString( "Email" ) <> "" and RequestString( "Password" ) <> "" ) then
   sEmail = RequestString( "Email" )
   sPass = RequestString( "Password" )
else
   sEmail = "" : sPass = ""
end if

if ( sEmail <> "" and sPass <> "" ) then
   call gcUser.Init( sEmail, sPass )
   if ( gcUser.Login() ) then
      gcUser.bPersist = true
      sRedir = gcUser.HomePage
      DestroyUser()
      if ( RequestString( "goto" ) <> "" ) then sRedir = RequestString( "goto" ) & "?" & RestoreRequestString() end if
      Redirect( hRef( sRedir ) )
   end if
end if
%>

<input type="hidden" name="goto" value="<%=RequestString( "Goto" )%>">

<h1>Please <strong>login</strong></h1>
<table border='0'>
   <% if ( sEmail <> "" or sPass <> "" ) then %>
      <tr><td>&nbsp;</td><td><b>Invalid login.</b></td></tr>
   <% end if %>
   <tr>
      <td align='right'>Email:</td>
      <td><input type='text' name='LoginEmail' size='30' maxlength='100'></td>
   </tr>
   <tr>
      <td align='right'>Password:</td>
      <td><input type='password' name='LoginPassword' size='30' maxlength='100'></td>
   </tr>
   <tr><td>&nbsp;</td><td><input type='button' value='Login' onClick='OnSubmit()' class='btn'></td></tr>
</table>

<!--#include virtual="/Framework/Close.asp"-->
