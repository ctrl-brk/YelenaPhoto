<%Option Explicit%>
<!--#include virtual="/Framework/Open.asp"-->

<h1>Welcome to <strong>the photo site!</strong></h1>
<p style='text-align:justify'>
   <b>Hi! My name is Yelena. I am a Wedding and Lifestyle Photographer in Denver,CO metro area</b>, providing creative, artistic and unique photography inspired by your beauty, emotions, style and fashion.
   I am able to capture most important moments of your wedding day and present them in beautiful professionally designed storytelling wedding album.
   I also offer commercial and private portrait sessions, including high school seniors, family, maternity and kids. The studio's product line includes albums, coffee table books,
   wedding photo invitations, post cards, picture frames, large mounted and gallery canvas wrapped prints.
</p>
<br>
<p>
   <span class='highlite'>Please note that we offer FREE SERVICES for non-profit organizations, community events, charities, pet rescue, etc.</span>
</p>
<hr>
<table border='0' width='100%'>
   <tr>
      <td style='border-right:1px solid #62584e; width:60%; vertical-align:top'>
         <h1 style='margin:0 0 2px 0;padding:0'>Latest <strong>works</strong></h1>
         <% dim obRec
         set obRec = obConn.Execute( DebugSQL( "exec GetFeaturedImages" ) )
         ClearDebug()

         if ( not obRec.EOF ) then %>
            <table border='0' class='tblLatest' cellspacing='0'>
               <tr>
                  <% while( not obRec.EOF ) %>
                     <td><a href='/Gallery.asp?id=<%=obRec( "GalleryId" )%>&iid=<%=obRec( "ImageId" )%>'><img src='<%=obRec( "RelativePath" ) & "/thmb/" & obRec( "ImageName" )%>'></a></td>
                     <% obRec.MoveNext()
                  wend %>
            </tr>
         </table>
         <% else %>
         Coming soon...
         <% end if
         obRec.Close() : set obRec = nothing %>
      </td>
      <td style='padding-left:10px; vertical-align:top'>
         <h1>My <strong>news</strong></h1>
         <b>Winter season is coming!</b><br>
         Summer season is over and there were some beautiful weddings I was pleasured to service.
         I'm still looking forward for more because I love taking pictures of people.
         But don't get cold at winter. Please contact me with any photography needs you have.
      </td>
   </tr>
</table>


<!--#include virtual="/Framework/Close.asp"-->
