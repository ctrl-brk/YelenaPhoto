<% if ( gbGallery <> true and gbAdmin <> true ) then %>
         </td>
      </tr>
   </table>
<% end if %>
</div>
</div>

<% if ( gbAdmin <> true ) then %>
   <div align='center'>
      <div class='<% if ( gbGallery = true ) then %>g<% end if %>footer'>
         <div class='cpy'>Copyright &copy; <%=DatePart( "yyyy", Now() ) %></div>
      </div>
   </div>
<% end if %>

</form>

<script src="http://www.google-analytics.com/urchin.js" type="text/javascript"></script>
<script type="text/javascript">
   _uacct = "UA-3862081-1";
   urchinTracker();
</script>

<img src='/Images/btn/galleries_over.gif' style='display:none'>
<img src='/Images/btn/testimonials_over.gif' style='display:none'>
<img src='/Images/btn/store_over.gif' style='display:none'>
<img src='/Images/btn/about_over.gif' style='display:none'>

</body>
</html>
<!--#include virtual="/Include/CloseDBConnect.asp"-->
<!--#include virtual="/Include/DestroyUser.asp"-->
