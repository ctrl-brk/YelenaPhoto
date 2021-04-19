<script language="VBScript" runat="Server">

sub dbg( str )

   Response.Clear()
   Response.Write( str )
   Response.End()

end sub

'-----------------------------------------------------------------------------

function DebugSQL( sSQL )

   Session( "ssn_sSQL" ) = sSQL
   DebugSQL = sSQL

end function

'-----------------------------------------------------------------------------

function ClearDebug()

   ClearDebug = Session( "ssn_sSQL" )
   Session( "ssn_sSQL" ) = ""

end function

'-----------------------------------------------------------------------------

</script>
