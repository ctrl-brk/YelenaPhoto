<script language="VBScript" runat="Server">

class Browser

   private obBT, sAgent, bLoaded

   private sub Class_Initialize
      bLoaded = false
      set obBT = Server.CreateObject( "MSWC.BrowserType" )
      sAgent = Request.ServerVariables( "HTTP_USER_AGENT" )
   end sub

   private sub Class_Terminate
      set obBT = nothing
   end sub

   private sub Load()
      if ( bLoaded ) then exit sub
      set obBT = Server.CreateObject( "MSWC.BrowserType" )
      bLoaded = true
   end sub

   public property get Browser
      Browser = "Unknown"
      if ( InStr( sAgent, "MSIE" ) > 0 ) then
         Browser = "IE"
      elseif ( InStr( sAgent, "Opera" ) > 0 ) then
         Browser = "Opera"
      elseif ( InStr( sAgent, "Firefox" ) > 0 ) then
         Browser = "Firefox"
      elseif ( InStr( sAgent, "Navigator" ) > 0 ) then
         Browser = "Netscape"
      end if
   end property

   public property get bIE
      if ( Browser = "IE" ) then bIE = true else bIE = false end if
   end property

   public property get bOpera
      if ( Browser = "Opera" ) then bOpera = true else bOpera = false end if
   end property

   public property get bFirefox
      if ( Browser = "Firefox" ) then bFirefox = true else bFirefox = false end if
   end property

   public property get bNetscape
      if ( Browser = "Netscape" ) then bNetscape = true else bNetscape = false end if
   end property

   public property get BrowserCap
      Load()
      BrowserCap = obBT.Browser
   end property

   public property get Version
      Load()
      Version = obBT.Version
   end property

   public property get MajorVer
      Load()
      MajorVer = obBT.Majorver
   end property

   public property get Minorver
      Load()
      Minorver = obBT.Minorver
   end property

   public property get Platform
      Load()
      Platform = obBT.Platform
   end property

   public property get Frames
      Load()
      Frames = obBT.Frames
   end property

   public property get Tables
      Load()
      Tables = obBT.Tables
   end property

   public property get Cookies
      Load()
      Cookies = obBT.Cookies
   end property

   public property get JavaScript
      Load()
      JavaScript = obBT.JavaScript
   end property

end class

class CGallery

   private nId, obDB, bPopulated, bDeleted

   '---------------------------

   private sub Class_Initialize
      InitVars()
   end sub

   '---------------------------

   private sub Class_Terminate
      set obDB = nothing
   end sub

   '---------------------------

   private sub InitVars()
      nId = 0 : bPopulated = false : bDeleted = false
   end sub

   '---------------------------

   public sub Init( GalleryId )

      if ( not IsNumeric( GalleryId ) ) then exit sub

      nId = CLng( GalleryId )

   end sub

   '---------------------------

   private sub Populate()

      if ( not bPopulated and nId > 0 ) then
         
         dim sSQL
         sSQL = "exec GetGalleries @GalleryId=" & Id
         set obDB = bo_CreateLargeSQLObject( sSQL )
         bPopulated = true
         if ( obDB is nothing ) then
            nId = 0
         end if

      end if

   end sub

   '---------------------------

   public default function DB( sParam )
      Populate()
      DB = obDB( LCase( sParam ) )
   end function

   '---------------------------

   public sub Delete

      obConn.Execute( "exec DeleteGallery @GalleryId=" & Id )
      bDeleted = true

   end sub

   '---------------------------

   public property get Id
      if ( bDeleted ) then Id = 0 else Id = nId end if
   end property

   public property get Exists
      Populate()
      if ( Id = 0 ) then Exists = false else Exists = true end if
   end property

   public property get Name
      Populate()
      Name = Me( "GalleryName" )
   end property

   public property get Path
      Populate()
      Path = Server.MapPath( Me( "RelativePath" ) )
   end property

   public property get Thumbnail
      Populate()
      if ( Me( "Thumbnail" ) <> "" ) then Thumbnail = Me( "RelativePath" ) & "/Thumb.jpg" else Thumbnail = "/Galleries/Thumb.gif" end if
   end property

   public property get Active
      Populate()
      if ( Me( "ActiveFlag" ) = "Y" ) then Active = true else Active = false end if
   end property

end class

'-----------------------------------------------------------------------------

class CImage

   private nId, obDB, bPopulated, bDeleted

   '---------------------------

   private sub Class_Initialize
      InitVars()
   end sub

   '---------------------------

   private sub Class_Terminate
      set obDB = nothing
   end sub

   '---------------------------

   private sub InitVars()
      nId = 0 : bPopulated = false : bDeleted = false
   end sub

   '---------------------------

   public sub Init( ImageId )

      if ( not IsNumeric( ImageId ) ) then exit sub

      nId = CLng( ImageId )

   end sub

   '---------------------------

   private sub Populate()

      if ( not bPopulated and nId > 0 ) then
         
         dim sSQL
         sSQL = "exec GetImages @ImageId=" & Id
         set obDB = bo_CreateLargeSQLObject( sSQL )
         bPopulated = true
         if ( obDB is nothing ) then nId = 0

      end if

   end sub

   '---------------------------

   public default function DB( sParam )
      Populate()
      DB = obDB( LCase( sParam ) )
   end function

   '---------------------------

   public sub Delete

      obConn.Execute( "exec DeleteGalleryImage @ImageId=" & Id )
      bDeleted = true

   end sub

   '---------------------------

   public property get Id
      if ( bDeleted ) then Id = 0 else Id = nId end if
   end property

   public property get Exists
      Populate()
      if ( Id = 0 ) then Exists = false else Exists = true end if
   end property

   public property get Featured
      Populate()
      if ( Me( "FeaturedFlag" ) = "Y" ) then Featured = true else Featured = false end if
   end property

   public property get Name
      Populate()
      Name = Me( "ImageName" )
   end property

   public property get RelativePath
      Populate()
      Name = Me( "RelativePath" ) & "/" & Name
   end property

   public property get ActiveFlag
      Populate()
      if ( Me( "ActiveFlag" ) = "Y" ) then ActiveFlag = true else ActiveFlag = false end if
   end property

   public property get Width
      Populate()
      Width = Me( "Width" )
   end property

   public property get Height
      Populate()
      Height = DB( "Height" )
   end property

   public property get ThumbWidth
      Populate()
      ThumbWidth = Me( "ThumbWidth" )
   end property

   public property get ThumbHeight
      Populate()
      ThumbHeight = DB( "ThumbHeight" )
   end property

end class

'-----------------------------------------------------------------------------

class CUser

   private nId, obDB, bPopulated

   '---------------------------

   private sub Class_Initialize
      InitVars()
   end sub

   '---------------------------

   private sub Class_Terminate
      set obDB = nothing
   end sub

   '---------------------------

   private sub InitVars()
      nId = 0 : bPopulated = false
   end sub

   '---------------------------

   public sub Init( UserId )

      if ( not IsNumeric( UserId ) ) then exit sub

      nId = CLng( UserId )

   end sub

   '---------------------------

   private sub Populate()

      if ( not bPopulated and nId > 0 ) then
         
         dim sSQL
         sSQL = "exec GetWebUser @UserId=" & Id
         set obDB = bo_CreateSQLObject( sSQL )
         bPopulated = true
         if ( obDB is nothing ) then
            nId = 0
         end if

      end if

   end sub

   '---------------------------

   public default function DB( sParam )
      Populate()
      DB = obDB( LCase( sParam ) )
   end function

   '---------------------------

   public property get Id
      Id = nId
   end property

   public property get Exists
      Populate()
      if ( nId = 0 ) then Exists = false else Exists = true end if
   end property

   public property get FirstName
      Populate()
      FirstName = Me( "FirstName" )
   end property

   public property get LastName
      Populate()
      LastName = Me( "FirstName" )
   end property

   public property get Name
      Name = FirstName & " " & LastName
   end property

   public property get ActiveFlag
      Populate()
      Status = Me( "ActiveFlag" )
   end property

   public property get Email
      Populate()
      Email = Me( "Email" )
   end property

   public property get UserType
      Populate()
      UserType = DB( "UserType" )
   end property

end class

'-----------------------------------------------------------------------------

function bo_CreateSQLObject( sSQL )

   dim obRec, obField, obDict

   set obRec = obConn.Execute( DebugSQL( sSQL ) )
   ClearDebug()

   if ( not obRec.EOF ) then
      set obDict = Server.CreateObject( "Scripting.Dictionary" )

      for each obField in obRec.Fields
         call obDict.Add( LCase( obField.Name ), CStr( NullToStr( obField.Value ) ) )
      next

      set bo_CreateSQLObject = obDict
   else
      set bo_CreateSQLObject = nothing
   end if

   obRec.Close()
   set obRec = nothing : set obDict = nothing

end function

'-----------------------------------------------------------------------------

function bo_CreateLargeSQLObject( sSQL )

   dim obRec, obField, obDict

   set obRec = Server.CreateObject( "ADODB.Recordset" )
   obRec.CursorLocation = 3
   call obRec.Open( DebugSQL( sSQL ), obConn, 3 )
   ClearDebug()

   if ( not obRec.EOF ) then
      set obDict = Server.CreateObject( "Scripting.Dictionary" )

      for each obField in obRec.Fields
         call obDict.Add( LCase( obField.Name ), CStr( NullToStr( obField.Value ) ) )
      next

      set bo_CreateLargeSQLObject = obDict
   else
      set bo_CreateLargeSQLObject = nothing
   end if

   obRec.Close()
   set obRec = nothing : set obDict = nothing

end function

</script>
