<%
dim obConn, gb_Connected

sub OpenDBConnect()

   if ( gb_Connected <> true ) then

      set obConn = Server.CreateObject( "ADODB.Connection" )
      obConn.Open( GetRegValue( "DSN") )
      gb_Connected = true

   end if

end sub

OpenDBConnect()
%>
