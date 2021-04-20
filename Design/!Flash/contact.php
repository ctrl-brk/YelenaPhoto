<?
Error_Reporting(E_ALL & ~E_NOTICE);

while ( $request = current( $_REQUEST ) )
   {
   if ( key ( $_REQUEST ) != 'recipient' )
      {
      $pre_array = split( "&777&",  $request );
      $post_vars[ key( $_REQUEST ) ][ 0 ] = $pre_array[0];
      $post_vars[ key( $_REQUEST ) ][ 1 ] = $pre_array[1];
      }
   next($_REQUEST);
   }

reset($post_vars);

$subject = "From ".$post_vars['your_name' ][0];
$headers = "From: ".$post_vars['your_email'][0] ."\n";
$headers.='Content-type: text/html; charset=iso-8859-1';

$message='';

while ( $mess = current( $post_vars ) )
   {
   if ( ( key( $post_vars ) != "i" ) && ( key( $post_vars ) != "your_email" ) && ( key( $post_vars ) != "your_name" ) )
      {
      $message.="<strong>".$mess[1]."</strong>&nbsp;&nbsp;&nbsp;".$mess[0]."<br>";
      }
   next( $post_vars );
   }

mail( 'yelena@yelenaphoto.com', $subject,  "
<html>
<head>
 <title>Contact letter from yelenaphoto.com</title>
</head>
<body>
<br>
  ".$message."
</body>
</html>" , $headers );

echo ( "Your message was successfully sent!" );

?>
<script>
   resizeTo(300, 300);
</script>
