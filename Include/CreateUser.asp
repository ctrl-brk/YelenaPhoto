<% dim gcUser

sub CreateUser()

   set gcUser = new CurrentUser
   gcUser.bPersist = false
end sub

CreateUser()
%>
