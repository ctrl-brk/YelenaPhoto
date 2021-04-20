<?php
// Auto creating MethodTable
// include(AMFPHP_BASE . "util/MethodTable.php");

require_once("thumbnail.inc.php");

class Gallery
{

    var $galleryDataFilePath = '../../../../tfile_gallery.php';
    var $galleryDefaultsFilePath = '../../../../tfile_gallery_defaults.html';

    function Gallery()
   {
       // $this->methodTable = MethodTable::create(__FILE__);

        // Define the methodTable for this class in the constructor
        $this->methodTable = array(
            "getGalleryOptions" => array(
                "description" => "Return gallery preferences",
                "access" => "remote"
            ),
            "getPreferences" => array(
                "description" => "Return gallery preferences",
                "access" => "remote"
            ),
            "getData" => array(
                "description" => "Return all the images from the gallery",
                "access" => "remote"
            ),
            "putData" => array(
                "description" => "Put categoories and images to the gallery file",
                "access" => "remote"
            )
        );

    }

    /**
        @desc Return gallery options
        @access remote
        @return array
    */
    function getGalleryOptions()
    {
        $tempArray = array();
        $dataArray = array();

        $fileContents = file_get_contents($this->galleryDefaultsFilePath);
        $fileContents = str_replace("\n", "", $fileContents);
        $fileContents = str_replace("\r", "", $fileContents);

        $tempArray = explode("&", $fileContents);

        for ($i=0; $i < count($tempArray); $i++) {
            if(!is_null($tempArray[$i]) && strlen($tempArray[$i]) > 0 && substr_count($tempArray[$i], "=") == 1) {
                $values = explode("=", $tempArray[$i]);
                $dataArray[$values[0]] = $values[1];
            }
        }

        // Options
        $options = array();

        if(function_exists("gd_info")) {
         $options['gd'] = 'true';
      } else {
            $options['gd'] = 'false';
      }

        $options['maxGalleriesAllowed'] = $dataArray['maxGalleriesAllowed'];

        $options['minImageWidth'] = $dataArray['minImageWidth'];
        $options['minImageHeight'] = $dataArray['minImageHeight'];
        $options['maxImageWidth'] = $dataArray['maxImageWidth'];
        $options['maxImageHeight'] = $dataArray['maxImageHeight'];

        $options['minPreviewOffset'] = $dataArray['minPreviewOffset'];
        $options['maxPreviewOffset'] = $dataArray['maxPreviewOffset'];

        $options['minHSpace'] = $dataArray['minHSpace'];
        $options['maxHSpace'] = $dataArray['maxHSpace'];

        $options['minVSpace'] = $dataArray['minVSpace'];
        $options['maxVSpace'] = $dataArray['maxVSpace'];

        $options['minListingOffset'] = $dataArray['minListingOffset'];
        $options['maxListingOffset'] = $dataArray['maxListingOffset'];

        $options['minColsNumber'] = $dataArray['minColsNumber'];
        $options['maxColsNumber'] = $dataArray['maxColsNumber'];

        $options['minRowsNumber'] = $dataArray['minRowsNumber'];
        $options['maxRowsNumber'] = $dataArray['maxRowsNumber'];

        $options['maxCommentChars'] = $dataArray['maxCommentChars'];

        $options['preview_position'] = $dataArray['preview_position'];
        $options['preview_offset'] = $dataArray['preview_offset'];
        $options['image_width'] = $dataArray['image_width'];
        $options['image_height'] = $dataArray['image_height'];
        $options['h_space'] = $dataArray['h_space'];
        $options['v_space'] = $dataArray['v_space'];
        $options['open_in_browser'] = $dataArray['open_in_browser'];
        $options['comment_position'] = $dataArray['comment_position'];
        $options['listing_offset'] = $dataArray['listing_offset'];
        $options['colls'] = $dataArray['colls'];
        $options['rows'] = $dataArray['rows'];

        return $options;
    }

    /**
        @desc Return gallery preferences
        @access remote
        @return array
    */
    function getPreferences()
    {
        $tempArray = array();
        $dataArray = array();

        $fileContents = file_get_contents($this->galleryDataFilePath);
        $fileContents = str_replace("\n", "", $fileContents);
        $fileContents = str_replace("\r", "", $fileContents);

        $tempArray = explode("&", $fileContents);

        for ($i=0; $i < count($tempArray); $i++) {
            if(!is_null($tempArray[$i]) && strlen($tempArray[$i]) > 0 && substr_count($tempArray[$i], "=") == 1) {
                $values = explode("=", $tempArray[$i]);
                $dataArray[$values[0]] = $values[1];
            }
        }

        // Prefrences
        $preferences = array();

        $preferences['preview_position'] = $dataArray['preview_position'];
        $preferences['preview_offset'] = $dataArray['preview_offset'];
        $preferences['image_width'] = $dataArray['image_width'];
        $preferences['image_height'] = $dataArray['image_height'];
        $preferences['h_space'] = $dataArray['h_space'];
        $preferences['v_space'] = $dataArray['v_space'];
        $preferences['colls'] = $dataArray['colls'];
        $preferences['rows'] = $dataArray['rows'];
        $preferences['listing_offset'] = $dataArray['listing_offset'];
        $preferences['open_in_browser'] = $dataArray['open_in_browser'];
        $preferences['comment_position'] = $dataArray['comment_position'];

        return $preferences;
    }

    /**
        @desc Return all the images from the gallery
        @access remote
        @return array
    */
    function getData()
    {
        $galleryPath = '../../../../';

        $tempArray = array();
        $dataArray = array();

        $fileContents = file_get_contents($this->galleryDataFilePath);
        $fileContents = str_replace("\n", "", $fileContents);
        $fileContents = str_replace("\r", "", $fileContents);

        $tempArray = explode("&", $fileContents);

        for ($i=0; $i < count($tempArray); $i++) {
            if(!is_null($tempArray[$i]) && strlen($tempArray[$i]) > 0 && substr_count($tempArray[$i], "=") == 1) {
                $values = explode("=", $tempArray[$i]);
                $dataArray[$values[0]] = $values[1];
            }
        }

        // Categories
        $gallery_images = array();
        $i = 0;
        while(isset($dataArray['cat_'.$i])) {
            $gallery_images[$i]['id'] = $i;
            $gallery_images[$i]['category_name'] = $dataArray['cat_'.$i];

            $j = 0;
            while(isset($dataArray['cat_'.$i.'_pic_'.$j])) {
                $gallery_images[$i]['images'][$j]['id'] = $j;
                $gallery_images[$i]['images'][$j]['pic'] = $dataArray['cat_'.$i.'_pic_'.$j];
                $gallery_images[$i]['images'][$j]['comment'] = $dataArray['cat_'.$i.'_comment_'.$j];

                // Get image info (filename, filesize, image width, image height)
                $imagePath = $galleryPath.$gallery_images[$i]['category_name'].'/big/'.$gallery_images[$i]['images'][$j]['pic'];
                if (file_exists($imagePath)) {
                    if(function_exists("gd_info")) {
                        $image = new Thumbnail($imagePath);
                        $gallery_images[$i]['images'][$j]['filesize'] = filesize($imagePath)/1024;
                        $gallery_images[$i]['images'][$j]['width'] = $image->getCurrentWidth();
                        $gallery_images[$i]['images'][$j]['height'] = $image->getCurrentHeight();
                        $image->destruct();
                    } else {
                        $gallery_images[$i]['images'][$j]['filesize'] = filesize($imagePath)/1024;
                        $isize = getimagesize($imagePath);
                        $gallery_images[$i]['images'][$j]['width'] = ($isize[0]) ? $isize[0] : '';
                        $gallery_images[$i]['images'][$j]['height'] = ($isize[1]) ? $isize[1] : '';
                    }
                } else {
                  $gallery_images[$i]['images'][$j]['filesize'] = ' ';
                    $gallery_images[$i]['images'][$j]['width'] = '';
                    $gallery_images[$i]['images'][$j]['height'] = '';
                }
                $j++;
                if($j>9999) break;
            }
            $i++;
            if($i>9999) break;
        }
        return $gallery_images;
    }

    /**
        @desc Put categoories and images to the gallery file
        @access remote
        @return boolean
    */
    function putData($dataObject, $preferences, $options, $imagesToRemove = null, $moveTo = false)
    {
        $galleryPath = '../../../../';
        $categoryDeleted = false;

        // GD check
        /*
        if(!function_exists("gd_info")) {
         return array('error' => true, 'message' => 'You do not have the GD Library installed.  Gallery admin tool requires the GD library to function properly.');
      }
      */

        $dataString = "// GALLERY \n\n";

        for($i = 0; $i < count($dataObject); $i++) {
         $categoryNumberToWrite = $dataObject[$i]['id'];
            // Check if category was deleted
            if(isset($dataObject[$i]['deleteCategory']) && $dataObject[$i]['deleteCategory'] == 'delete') {
                //@chmod($galleryPath.$dataObject[$i]['name'], 0777);
                if($this->deleteDirectory($galleryPath.$dataObject[$i]['name'])) {
                  $categoryDeleted = true;
                } else {
                  return array('error' => true, 'message' => 'Couldn\'t delete category folder.');
                }
            } else {
               if($categoryDeleted) $categoryNumberToWrite--;

                // Create new gallery folder
                if(isset($dataObject[$i]['newCategory']) && $dataObject[$i]['newCategory'] == 'create') {
                    if(!mkdir($galleryPath.$dataObject[$i]['name'])) {
                        return array('error' => true, 'message' => 'Couldn\'t create new category folder.');
                    } else {
                        //chmod($galleryPath.$dataObject[$i]['name'], 0777);
                        if(!mkdir($galleryPath.$dataObject[$i]['name'].'/big')) return array('error' => true, 'message' => 'Couldn\'t create folder inside category folder.');
                        if(!mkdir($galleryPath.$dataObject[$i]['name'].'/small')) return array('error' => true, 'message' => 'Couldn\'t create folder inside category folder.');
                    }
                }

                $dataString .= "&cat_".$categoryNumberToWrite."=".$dataObject[$i]['name']."& \n\n";

                // Rename folder if category name was changed
                if(isset($dataObject[$i]['old_name']) && $dataObject[$i]['name'] != $dataObject[$i]['old_name']) {
                    if(!rename($galleryPath.$dataObject[$i]['old_name'], $galleryPath.$dataObject[$i]['name'])) {
                        return array('error' => true, 'message' => 'Couldn\'t rename category folder.');
                    }
                }

                // Images loop
                for($j = 0; $j < count($dataObject[$i]['images']); $j++) {
                    if(isset($dataObject[$i]['images'][$j]['renamePic'])) {
                     $bigImagePath = $galleryPath.$dataObject[$i]['name']."/big/".$dataObject[$i]['images'][$j]['pic'];
                     $smallImagePath = $galleryPath.$dataObject[$i]['name']."/small/".$dataObject[$i]['images'][$j]['pic'];
                        // Thumbnail
                        if($dataObject[$i]['images'][$j]['renamePic'] == 'createThumbnail') {
                           if (file_exists($bigImagePath) && function_exists("gd_info")) {
                               $thumb = new Thumbnail($bigImagePath);
                               $thumb->resize($preferences['image_width'], $preferences['image_height']);
                               $thumb->save($smallImagePath);
                               $thumb->destruct();
                           }
                        } else {
                            //chmod($galleryPath.$dataObject[$i]['name']."/small/".$dataObject[$i]['images'][$j]['renamePic'], 0755);
                            rename($galleryPath.$dataObject[$i]['name']."/small/".$dataObject[$i]['images'][$j]['renamePic'], $smallImagePath);
                            // CHeck thumbnail size - resample
                            if (file_exists($smallImagePath) && function_exists("gd_info")) {
                               $thumb = new Thumbnail($smallImagePath);
                               if($thumb->getCurrentWidth() > $preferences['image_width'] || $thumb->getCurrentHeight() > $preferences['image_height']) {
                                   $thumb->resize($preferences['image_width'], $preferences['image_height']);
                                   $thumb->save($smallImagePath);
                               }
                               $thumb->destruct();
                            }
                        }
                        // Check size of the new uploaded full image
                        if (file_exists($bigImagePath) && function_exists("gd_info")) {
                           $thumb = new Thumbnail($bigImagePath);
                           if($thumb->getCurrentWidth() > $options['maxPreviewWidth'] || $thumb->getCurrentHeight() > $options['maxPreviewHeight']) {
                               $thumb->resize($options['maxPreviewWidth'], $options['maxPreviewHeight']);
                               $thumb->save($bigImagePath);
                           }
                           $thumb->destruct();
                        }
                    }

                    $dataString .= "&cat_".$categoryNumberToWrite."_pic_".$j."=".$dataObject[$i]['images'][$j]['pic']."& \n";

                    if(strlen($dataObject[$i]['images'][$j]['comment']) > 0) {
                        $dataString .= "&cat_".$categoryNumberToWrite."_comment_".$j."=".$dataObject[$i]['images'][$j]['comment']."& \n";
                    }
                }
                $dataString .= "\n\n";

            }
        }
        
        $dataString .= "// PREFERENCES \n\n";

        foreach ($preferences as $key => $value) {
            if($key == 'length') continue;
            $dataString .= "&".$key."=".$value."& \n";
        }
        $dataString .= "\n\n";

        // Remove images
        if($imagesToRemove && count($imagesToRemove) > 0) {
            if($moveTo) {
                for($i = 0; $i < count($imagesToRemove); $i++) {
                    copy($galleryPath.$imagesToRemove[$i]['category']."/big/".$imagesToRemove[$i]['pic'], $galleryPath.$imagesToRemove[$i]['target']."/big/".$imagesToRemove[$i]['pic']);
                    copy($galleryPath.$imagesToRemove[$i]['category']."/small/".$imagesToRemove[$i]['pic'], $galleryPath.$imagesToRemove[$i]['target']."/small/".$imagesToRemove[$i]['pic']);

                    $imageToRemove = $galleryPath.$imagesToRemove[$i]['category']."/big/".$imagesToRemove[$i]['pic'];
                  if (file_exists($imageToRemove)) {
                     @unlink($imageToRemove);
                  }
                  $imageToRemove = $galleryPath.$imagesToRemove[$i]['category']."/small/".$imagesToRemove[$i]['pic'];
                  if (file_exists($imageToRemove)) {
                     @unlink($imageToRemove);
                  }
                }
            } else {
                for($i = 0; $i < count($imagesToRemove); $i++) {
                  $imageToRemove = $galleryPath.$imagesToRemove[$i]['category']."/big/".$imagesToRemove[$i]['pic'];
                  if (file_exists($imageToRemove)) {
                     @unlink($imageToRemove);
                  }
                  $imageToRemove = $galleryPath.$imagesToRemove[$i]['category']."/small/".$imagesToRemove[$i]['pic'];
                  if (file_exists($imageToRemove)) {
                     @unlink($imageToRemove);
                  }
                }
            }
        }
        
        // Put contents
        $file = @fopen($this->galleryDataFilePath, "w");

        if(!$file) {
            return array('error' => true, 'message' => 'Couldn\'t open gallery data file.');
        } else {
            fwrite($file, $dataString);
            fclose($file);
            return array();
        }
    }

    function deleteDirectory($dirname, $only_empty=false) {
       if (!is_dir($dirname))
           return false;
       $dscan = array(realpath($dirname));
       $darr = array();
       while (!empty($dscan)) {
           $dcur = array_pop($dscan);
           $darr[] = $dcur;
           if ($d=opendir($dcur)) {
               while ($f=readdir($d)) {
                   if ($f=='.' || $f=='..')
                       continue;
                   $f=$dcur.'/'.$f;
                   if (is_dir($f))
                       $dscan[] = $f;
                   else
                      unlink($f);
               }
               closedir($d);
           }
       }
       $i_until = ($only_empty)? 1 : 0;
       for ($i=count($darr)-1; $i>=$i_until; $i--) {
           rmdir($darr[$i]);
       }
       return (($only_empty)? (count(scandir)<=2) : (!is_dir($dirname)));
    }

}

?>