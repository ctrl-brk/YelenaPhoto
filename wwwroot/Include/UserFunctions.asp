<script language="VBScript" runat="Server">

function usr_CreateSQLObject( sSQL )

   dim obRec, obField, obDict

   set obRec = obConn.Execute( DebugSQL( sSQL ) )
   ClearDebug()

   if ( not obRec.EOF ) then
      set obDict = Server.CreateObject( "Scripting.Dictionary" )

      for each obField in obRec.Fields
         obDict.Add obField.Name, CStr( NullToStr( obField.Value ) )
      next

      set usr_CreateSQLObject = obDict
   else
      set usr_CreateSQLObject = nothing
   end if

   obRec.Close()
   set obRec  = nothing : set obDict = nothing

end function

'-----------------------------------------------------------------------------

sub LoginRedirect( sURL )

   call Redirect( "/Login.asp?goto=" & sURL & "&" & RestoreRequestString() )

end sub

'-----------------------------------------------------------------------------

class CurrentUser

   private sEmail, sPassword, nId, sGUID, bPopulated, obDB
   public  bPersist
   
   '---------------------------
   private sub ClassInit()
      sEmail = "" : sPassword = "" : nId = 0 : bPopulated = false
      bPersist = true
   end sub

   '---------------------------
   private sub Class_Initialize
      ClassInit()
      Deserialize()
   end sub

   '---------------------------
   private sub Class_Terminate
      Serialize()
   end sub

   '---------------------------
   private function Serialize()
      if ( bPersist ) then
         Session( "cu__Serialized" ) = true
         Session( "cu__Email"      ) = sEmail
         Session( "cu__Password"   ) = sPassword
         Session( "cu__GUID"       ) = sGUID
         if ( nId <> 0 ) then Session( "cu_UserId" ) = nId
      end if
   end function

   '---------------------------
   private function Deserialize()
      if ( Session( "cu__Serialized" ) = true ) then 
         if ( Session( "cu__Email"    ) <> "" ) then sEmail    = Session( "cu__Email"    )
         if ( Session( "cu__Password" ) <> "" ) then sPassword = Session( "cu__Password" )
         if ( Session( "cu__GUID"     ) <> "" ) then sGUID     = Session( "cu__GUID"     )
         if ( Session( "cu_UserId"    ) <> "" ) then nId = CLng( Session( "cu_UserId"  ) )
      end if
   end function

   '---------------------------
   public function Init( eml, pass )

      sEmail = eml : sPassword = pass

   end function

   '---------------------------
   public sub Logout()

      dim sVar, bFound

      Response.Cookies( "yelenaphoto.com" )( "UserGUID" ) = ""
      ClassInit()
      bPersist = false
      do
         bFound = false
         for each sVar in Session.Contents
            if ( Left( sVar, 3 ) = "cu_" ) then
               Session.Contents.Remove( sVar )
               bFound = true
               exit for
            end if
         next
      loop while bFound = true

   end sub

   '---------------------------
   private function Populate()
      if ( not bPopulated ) then
         set obDB = usr_CreateSQLObject( "exec GetWebUser @UserId=" & nId )
         bPopulated = true
      end if
   end function

   '---------------------------
   public default function DB( sParam )
      Populate()
      DB = obDB( sParam )
   end function

   '---------------------------
   public function Login()

      dim obRec

      if ( bLoggedIn ) then
         Login = true
      else
         set obRec = obConn.Execute( DebugSQL( "exec UserLogin @Email=" & ToSQLString( sEmail ) & ",@Password=" & ToSQLString( sPassword ) & ",@IP='" & Request.ServerVariables( "REMOTE_ADDR" ) & "'" ) )
         ClearDebug()
         if ( not IsNull( obRec( "UserId" ) ) ) then
            nId = obRec( "UserId" ).Value
            sGUID = obRec( "GUID" )
            
            if ( Request.Cookies( "yelenaphoto.com" )( "UserGUID" ) = "" and Response.Buffer = true ) then
               Response.Cookies( "yelenaphoto.com" )( "UserGUID" ) = obRec( "GUID" )
               Response.Cookies( "yelenaphoto.com" ).Expires = Date() + 365
            end if
            Login = true

         else
            Login = false
         end if

         obRec.Close() : set obRec = nothing
      end if

   end function

   '---------------------------
   public function ConfirmLogin( sGUID )

      dim obRec

      ConfirmLogin = false
      on error resume next
      set obRec = obConn.Execute( "exec ConfirmWebUser @GUID=" & ToSQLString( sGUID ) & ",@IP='" & Request.ServerVariables( "REMOTE_ADDR" ) & "'" )
      if ( Err.Number <> 0 ) then exit function
      on error goto 0

      if ( not obRec.EOF ) then
         Logout()
         call Init( obRec( "Email" ), obRec( "Password" ) )
         Login()
         call SendMail( false, "YelenaPhoto.com user registration", "Email: " & sEmail & vbCrLf & "Comments: " & obRec( "Comments" ), GetEmailAddress( "From", "System" ), GetEmailAddress( "To", "Webmaster" ), "", "" )
         ConfirmLogin = true
      end if

      obRec.Close() : set obRec = nothing

   end function

   '---------------------------

   public sub LoginFromCookie()

      dim sGUID

      sGUID = Request.Cookies( "yelenaphoto.com" )( "UserGUID" )
      if ( sGUID <> "" ) then

         dim obRec

         set obRec = obConn.Execute( DebugSQL( "exec UserLogin @UserGUID=" & ToSQLString( sGUID ) & ",@IP='" & Request.ServerVariables( "REMOTE_ADDR" ) & "'" ) )
         ClearDebug()

         if ( not IsNull( obRec( "UserId" ).Value ) ) then
            call Init( obRec( "Email" ).Value, obRec( "Password" ).Value )
            nId = obRec( "UserId" ).Value
            sGUID = obRec( "GUID" )

            bPersist = true
         end if
            
      end if

   end sub

   '---------------------------

   public property get bLoggedIn
      bLoggedIn = false
      if ( nId > 0 ) then
         if ( bActive ) then bLoggedIn = true
      end if
   end property

   public property get bRegistered
      if ( nId > 0 ) then bRegistered = true else bRegistered = false end if
   end property

   public property get Id
     Id = nId
   end property

   public property get Status
     Status = DB( "Status" )
   end property

   public property get StatusDescription
   
   if ( nId = 0 ) then
      StatusDescription = "Not registered"
   else
      select case( Status )
         case "A"  : StatusDescription = "Acitve"
         case "W"  : StatusDescription = "Confirmation pending"
         case "N"  : StatusDescription = "Inactive"
         case "B"  : StatusDescription = "Disabled"
         case else : StatusDescription = "Unknown"
      end select
   end if

   end property

   public property get bActive
     bActive = false
     if ( nId > 0 ) then
        if ( Status = "A" ) then bActive = true
     end if
   end property

   public property get bAdmin
      bAdmin = false
      if ( bLoggedIn ) then
         if ( DB( "UserType" ) = "A" ) then bAdmin = true
      end if
   end property

   public property get FirstName
      FirstName = DB( "FirstName" )
   end property

   public property get LastName
      LastName = DB( "LastName" )
   end property

   public property get Name
      Name = FirstName & " " & LastName
   end property

   public property get Email
      Email = sEmail
   end property

   public property let Email( Value )
      sEmail = Value
      bPersist = true
   end property

   public property get Password
      Password = sPassword
   end property

   public property get HomePage
      
      HomePage = "/"

   end property

end class

</script>
