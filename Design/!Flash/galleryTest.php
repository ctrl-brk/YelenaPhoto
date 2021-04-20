<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>PHP Checker</title>
</head>
<body>
    <div style="padding:20px;">
        If PHP is successfully installed, you should see messages printed below the line.
        <br><br>
        If you do not have PHP installed, you may receive a server error message, the code you copied displayed, a prompt to download the PHP file, or any other number of reasons.
    </div>
    <hr>
    <div style="padding:20px;">
        <?php     
            // Echo
            echo 'PHP is installed';
            echo '<br><br>';

            // Version
            echo '- Current PHP version: ' . phpversion();
            echo '<br><br>';

            // GD
            if(function_exists('gd_info')) {
                echo '- GD library is installed.';
            } else {
                echo '- GD library is not installed on the server.';
            }
            echo '<br><br>';

            // Safe mode
            if(ini_get('safe_mode')) {
                echo '- Safe mode is ON. Gallery will not work.';
            }else{
                echo '- Safe mode is OFF.';
            }
            echo '<br><br>';

            // Checking php
            // phpinfo();
            
            if(!ini_get('safe_mode')) {
            
               echo 'Checking gallery folders and files...';
               echo '<br><br>';
               
               $dirs = array();
               $dirsLower = array();
               $galleryDataFile = '';
                          
               // Get _gallery.php (html) file contents
               if ($handle = opendir('.')) {
                while (false !== ($file = readdir($handle))) {
                    if ($file != "." && $file != "..") {
                     if(is_dir($file)) {
                        $dirs[] = $file;
                        $dirsLower[] = strtolower($file);
                     } else {             
                        if(preg_match("/\b((\S{5})(_gallery.)(html|php))\b/", $file)) {
                           echo '- Gallery data file found - '.$file;
                           $galleryDataFile = $file;
                           break;
                        }
                     }
                    }
                }
                closedir($handle);
            }
               
               // Get existing gallery categories from file if it was found
               if($galleryDataFile != '') {
                  
                  // Get file contents
                  $fileContents = file_get_contents($galleryDataFile);
                 $fileContents = str_replace("\n", "", $fileContents);
                 $fileContents = str_replace("\r", "", $fileContents);
         
                 $tempArray = explode("&", $fileContents);
         
                 for ($i=0; $i < count($tempArray); $i++) {
                     if(!is_null($tempArray[$i]) && strlen($tempArray[$i]) > 0 && substr_count($tempArray[$i], "=") == 1) {
                         $values = explode("=", $tempArray[$i]);
                         $dataArray[$values[0]] = $values[1];
                     }
                 }
                 
                 $categories = array();
                 
                 $i = 0;
                 // Loop through categories
               while(isset($dataArray['cat_'.$i])) {
                  $categories[] = $dataArray['cat_'.$i];
                  $i++;
               }
               
               // Check categories
               foreach($categories as $cat) {
                  if(in_array(strtolower($cat), $dirsLower)) {
                     $key = array_search(strtolower($cat), $dirsLower);
                     echo '<br><br>Category found - '.$cat.'.<br>';
                     // Permissions
                     echo 'Try to set permissions...<br>';              
                     if(chmod($dirs[$key], 0777)) {
                        echo 'OK. Permissions set.<br>';
                     } else {
                        echo 'ERROR. Couldn\'t set permissions on the category folder. Gallery may not work.<br>';
                     }
                     // Folders check
                     $cdirs = array();
                     $cdirsLower = array();
                     if ($handle = opendir($dirs[$key])) {
                         while (false !== ($file = readdir($handle))) {
                             if ($file != "." && $file != ".." && is_dir($dirs[$key].'/'.$file)) {
                              $cdirs[] = $file;
                              $cdirsLower[] = strtolower($file);
                             }
                         }
                         closedir($handle);
                     }
                     
                     // big                  
                     if(in_array('big', $cdirsLower)) {
                        $ckey = array_search('big', $cdirsLower);
                        // Permissions
                        if(chmod($dirs[$key].'/'.$cdirs[$ckey], 0777)) {
                           echo 'OK. Permissions set.<br>';
                        } else {
                           echo 'ERROR. Couldn\'t set permissions on the category folder. Gallery may not work.<br>';
                        }
                        // Case
                        if($cdirs[$ckey] != 'big') {
                           if(!rename($dirs[$key].'/'.$cdirs[$ckey], $dirs[$key].'/big')) {
                              echo 'ERROR. Couldn\'t rename folder inside category folder. Gallery may not work.<br>';
                           }
                        }
                     } else {
                        echo '`big` not found. Try to create...<br>';
                        if(mkdir($dirs[$key].'/big', 0777)) {
                           echo '`big` folder created.<br>';
                        } else {
                           echo 'ERROR. Couldn\'t create folder inside category folder. Gallery may not work.<br>';
                        }
                     }
                     
                     // small                
                     if(in_array('small', $cdirsLower)) {
                        $ckey = array_search('small', $cdirsLower);
                        // Permissions
                        if(chmod($dirs[$key].'/'.$cdirs[$ckey], 0777)) {
                           echo 'OK. Permissions set.<br>';
                        } else {
                           echo 'ERROR. Couldn\'t set permissions on the category folder. Gallery may not work.<br>';
                        }
                        // Case
                        if($cdirs[$ckey] != 'small') {
                           if(!rename($dirs[$key].'/'.$cdirs[$ckey], $dirs[$key].'/small')) {
                              echo 'ERROR. Couldn\'t rename folder inside category folder. Gallery may not work.<br>';
                           }
                        }
                     } else {
                        echo '`big` not found. Try to create...<br>';
                        if(mkdir($dirs[$key].'/small', 0777)) {
                           echo '`small` folder created.<br>';
                        } else {
                           echo 'ERROR. Couldn\'t create folder inside category folder. Gallery may not work.<br>';
                        }
                     }
                     
                     // Case
                     if($cat != $dirs[$key]) {
                        // Try to rename
                        echo 'Directory name has different letter case. Renaiming...<br>';
                        if(@rename($dirs[$key], $cat)) {
                           echo 'OK. '.$dirs[$key].' was successfully renamed to '.$cat.'.<br>';
                        } else {
                           echo 'ERROR. Couldn\'t rename category folder. Gallery may not work.<br>';
                        }
                     }
                     
                  } else {
                     echo 'Gallery category folder not found on the server - '.$cat.'.<br>';
                     echo 'Try to create...<br>';
                     if(!mkdir($cat, 0777)) {
                              echo 'ERROR. Couldn\'t create new category folder. Gallery may not work.<br>';
                          } else {
                           echo 'New category folder created.<br>';
                              if(mkdir($cat.'/big', 0777)) {
                           echo '`big` folder created.<br>';
                        } else {
                           echo 'ERROR. Couldn\'t create folder inside category folder. Gallery may not work.<br>';
                        }
                              if(mkdir($cat.'/small', 0777)) {
                           echo '`small` folder created.<br>';
                        } else {
                           echo 'ERROR. Couldn\'t create folder inside category folder. Gallery may not work.<br>';
                        }
                          }
                  }
               }
               
               } else {
                  echo '- Gallery data file not found. Gallery will not work.<br>';
               }
            }

        ?>
    </div>
</body>
</html>