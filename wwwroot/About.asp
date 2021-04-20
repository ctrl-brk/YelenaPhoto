<%Option Explicit%>
<!--#include virtual="/Framework/Open.asp"-->

<%
if ( RequestString( "Action" ) = "S" ) then

   dim sBody

   sBody = "Name: " & RequestString( "Name" ) & vbCrLf &_
           "Email: " & RequestString( "Email" )  & vbCrLf & vbCrLf &_
           RequestString( "Message" )

   call SendMail( false, "YelenaPhoto.com - Visitor's feedback", sBody, GetEmailAddress( "From", "System" ), GetEmailAddress( "To", "Webmaster" ), "", "" )

end if
%>
<script language='JavaScript'>

function OnLoad()
{
<% if ( RequestString( "Action" ) = "S" ) then %>alert( "Thank you for your message. I will get back to you ASAP." ); <% end if %>
}

function OnSubmit()
{
   SubmitForm( "About.asp" );
}

</script>

<input type='hidden' name='Action' value='S'>

<table border='0' cellspacing='0' class='tblAbout'>
   <tr>
      <td class='abst01'>
         <h1>About <strong>me</strong></h1>
         <p>
            <img src='/Images/Yelena.jpg' border='0' align='left'>
            Photography for me is one of the ways to communicate with the world around. Never ending source of bright and powerful emotions. I like to work in different areas of photography but taking pictures of people is my favorite!<br><br>
            When I look at someone through my viewfinder I just see the picture of his/her spirit and beauty.<br><br>
            I like to experiment with lighting, use different techniques and approaches... but the most exciting thing is to find unique and special aesthetics in every subject.<br><br>
            <b>Photography is unlimited source of possibilities and excitement!</b>
         </p>
      </td>
      <td class='abst02'>
         <h1>Contact <strong>me</strong></h1>
         <p>You are very welcome to contact me about any questions or suggestions you have. I would like to hear your opinion about my work and hopefuly make some new friends.</p>
         <br>
         Please feel free to<br>- email me at<br>&nbsp;&nbsp;&nbsp;<%=EmailLink( "Emails\To\Owner", "" ) %><br>- call me at<br>&nbsp;&nbsp;&nbsp;<span class='highlite'>(303) 882-4321</span><br> or simply use the form below.
      </td>
   </tr>
</table>
<hr>
<h1>Contact <strong>form</strong></h1>
<table border='0' cellspacing='0' cellpadding='0'>
   <tr>
      <td style='vertical-align:top;width:230px;'>
         Your name:<br>
         <input type='text' name='Name' maxlength='100' style='width:200px'><br><br>
         Email:<br>
         <input type='text' name='Email' maxlength='100' style='width:200px'>
      </td>
      <td valign='top'>
         Message:<br>
         <textarea name='Message' rows='4' style='width:200px'></textarea>
      </td>
      <td><input type='button' value='Submit' onClick='OnSubmit()' style='margin-left:55px;'></td>
   </tr>
</table>
<!--#include virtual="/Framework/Close.asp"-->
