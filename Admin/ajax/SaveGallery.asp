<%Option Explicit%>
<!--#include virtual="/Include/Functions.asp"-->
<!--#include virtual="/Include/DebugFunctions.asp"-->
<!--#include virtual="/Include/OpenDBConnect.asp"-->
<!--#include virtual="/Include/BizObj.asp"-->
<%
Response.ContentType = "text/plain;charset=windows-1251"
ExpirePage()

dim obRec, Stream, DOM, File, Resize, GalleryId, Name, Description, Path, SortValue, ActiveFlag, bImageExists, bResize, bRet, sFileName

set Stream = Server.CreateObject( "ADODB.Stream" )
set DOM = Server.CreateObject( "MSXML2.DOMDocument" )
DOM.load( Request )

call Dom.setProperty( "SelectionLanguage", "XPath" )

dim ImageExists
set ImageExists = DOM.selectSingleNode( "//root/ImageExists" )

set GalleryId = DOM.selectSingleNode( "//root/GalleryId" )
set Name = DOM.selectSingleNode( "//root/Name" )
set Description = DOM.selectSingleNode( "//root/Description" )
set Path = DOM.selectSingleNode( "//root/Path" )
set SortValue = DOM.selectSingleNode( "//root/SortValue" )
set ActiveFlag = DOM.selectSingleNode( "//root/ActiveFlag" )

if ( GalleryId.nodeTypedValue = "" ) then

   dim sFolder, obFS

   sFolder = Path.nodeTypedValue
   if ( Left( sFolder, 1 ) <> "/" ) then sFolder = "/" & sFolder
   set obFS = CreateObject( "Scripting.FileSystemObject" )
   if ( not obFS.FolderExists( Server.MapPath( sFolder ) ) ) then
      obFS.CreateFolder( Server.MapPath( sFolder ) )
      obFS.CreateFolder( Server.MapPath( sFolder & "/thmb" ) )
   end if
   set obFS = nothing

end if

bResize = false
if ( ImageExists.nodeTypedValue = "Y" ) then
   bImageExists = true
   set Resize = DOM.selectSingleNode( "//root/ResizeImage" )
   if ( Resize.nodeTypedValue = "Y" ) then bResize = true
else
   bImageExists = false
end if

if ( bImageExists ) then
   ' retrieve XML node with binary content
   set File = DOM.selectSingleNode( "//root/Image" )

   ' open stream object and store XML node content into it   
   Stream.Type = 1  ' 1=adTypeBinary 
   Stream.open() 
   Stream.Write( File.nodeTypedValue )
   ' save uploaded file
   if ( bResize ) then sFileName = "ThumbSource.jpg" else sFileName = "Thumb.jpg" end if
   call Stream.SaveToFile( Server.MapPath( Path.nodeTypedValue & "/" & sFileName ), 2 ) ' 2=adSaveCreateOverWrite 
   Stream.close()
end if

bRet = true
if ( bImageExists and bResize ) then bRet = CreateGalleryThumbnail( Server.MapPath( Path.nodeTypedValue & "/ThumbSource.jpg" ), Server.MapPath( Path.nodeTypedValue & "/Thumb.jpg" ) )

if ( GalleryId.nodeTypedValue <> "" ) then

   dim sSQL

   sSQL = "exec UpdateGallery @GalleryId=" & GalleryId.nodeTypedValue &_
                            ",@GalleryName=" & ToSQLString( Name.nodeTypedValue ) &_
                            ",@Description=" & ToSQLString( Description.nodeTypedValue ) &_
                            ",@Path=" & ToSQLString( Path.nodeTypedValue ) &_
                            ",@SortValue=" & ToSQLString( SortValue.nodeTypedValue ) &_
                            ",@ActiveFlag=" & ToSQLString( ActiveFlag.nodeTypedValue )

   if ( bImageExists ) then sSQL = sSQL & ",@Thumbnail='Thumb.jpg'" 

else

   sSQL = "exec AddGallery @Description=" & ToSQLString( Description.nodeTypedValue ) &_
                         ",@GalleryName=" & ToSQLString( Name.nodeTypedValue ) &_
                         ",@Path=" & ToSQLString( Path.nodeTypedValue ) &_
                         ",@SortValue=" & ToSQLString( SortValue.nodeTypedValue ) &_
                         ",@ActiveFlag=" & ToSQLString( ActiveFlag.nodeTypedValue )

   if ( bImageExists ) then sSQL = sSQL & ",@Thumbnail='Thumb.jpg'" 

end if

set obRec = obConn.Execute( sSQL )

Response.Write( obRec( "GalleryId" ) )
obRec.Close() : set obRec = nothing

set DOM = nothing : set GalleryId = nothing : set Name = nothing : set Description = nothing : set Path = nothing : set SortValue = nothing : set ActiveFlag = nothing

%>
<!--#include virtual="/Include/CloseDBConnect.asp"-->
