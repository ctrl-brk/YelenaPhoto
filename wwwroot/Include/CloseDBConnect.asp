<%
sub CloseDBConnect()

    if ( gb_Connected ) then  obConn.Close()
    set obConn = nothing

end sub

CloseDBConnect()
%>
