<?php

if(isset($_REQUEST['targetPath']) && opendir($_REQUEST['targetPath'])) {
   $target_path = $_REQUEST['targetPath'];
} else {
   $target_path = "../uploads/";
}

$filenameToMove = basename($_FILES['Filedata']['name']);

// Check for old files
function getFiles($directory) {
    // Try to open the directory
    if($dir = opendir($directory)) {
        // Create an array for all files found
        $tmp = Array();
        // Add the files
        while($file = readdir($dir)) {
            // Make sure the file exists
            if($file != "." && $file != ".." && $file[0] != '.') {
                // If it's a directory, list all files within it
                if(is_dir($directory . "/" . $file)) {
                    /*
                    $tmp2 = getFiles($directory . "/" . $file);
                    if(is_array($tmp2)) {
                        $tmp = array_merge($tmp, $tmp2);
                    }
                    */
                } else {
                    // array_push($tmp, $directory . $file);
                    array_push($tmp, $file);
                }
            }
        }
        // Finish off the function
        closedir($dir);
        return $tmp;
    }
}

/*
$filesArray = Array();
$filesArray = getFiles($target_path);

$newNameFound = false;
while ($newNameFound == false) {
   $nameNotFound = true;
   foreach ($filesArray as $value) {
      if($filenameToMove == $value) {
         $nameNotFound = false;
         $randomNumber = '';
         for($i=0; $i < 3; $i++) {
            $randomNumber .= rand(0, 9);
         }
         $filenameToMove = $randomNumber . '_' . $filenameToMove;
         break;
      }
   }
   
   if($nameNotFound) {
      $newNameFound = true;
      break;
   }
}
*/

$target_path = $target_path . $filenameToMove;

if(move_uploaded_file($_FILES['Filedata']['tmp_name'], $target_path))
{
    chmod($target_path, 0644);
    echo "The file ". basename( $_FILES['Filedata']['name']). " has been uploaded";
}
else
{
     echo "There was an error uploading the file, please try again!";
}

?>
