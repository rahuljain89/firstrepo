<link rel="stylesheet" href="../templates/ppezone/css/ppezone.css" type="text/css" media="screen" />
<?php
define('_JEXEC', 1);
define('JPATH_BASE', dirname(dirname(__FILE__)));
define('DS', DIRECTORY_SEPARATOR);
require_once ( JPATH_BASE . DS . 'includes' . DS . 'defines.php' );
require_once ( JPATH_BASE . DS . 'includes' . DS . 'framework.php' );

$execute = 0;
$mainframe = & JFactory::getApplication('site');
 
$filename = JPATH_BASE . "/scripts/equipmentproducts.xls";
@chmod($filename, 0777);
                        
$ext = pathinfo($filename, PATHINFO_EXTENSION);

if($ext == 'xlsx' || $ext == 'xls') {                                
    require_once(JPATH_BASE.'/components/com_jshopping/lib/excel_reader2.php');        
    require_once(JPATH_BASE.'/components/com_jshopping/lib/SpreadsheetReader.php');        
    $Spreadsheet = new SpreadsheetReader($filename);                
    $Spreadsheet->ChangeSheet(0);
    $data = array();
    foreach ($Spreadsheet as $Key => $Row)
    {
        $data[] = $Row;
    } 

    $total_rows = 0;
    
    $db = JFactory::getDBO();

    //Get Jobzones
    $query = $db->getQuery(true);
    $query = "SELECT GROUP_CONCAT(id) as jobzones from `#__jobrole`";      
    $db->setQuery($query);
    $jobzones = $db->loadObject();
    $jobzones = $jobzones->jobzones;

    //Get Categories
    $query = $db->getQuery(true);
    $query = "SELECT category_id, `name_en-GB` as category_title from `#__jshopping_categories` where category_parent_id = 100";      
    $db->setQuery($query);
    $equipmentCategories = $db->loadObjectList();
    
    $validCategories = array();
    foreach($equipmentCategories as $category) {
        $validCategories[trim(strtolower($category->category_title))] = $category->category_id;
    }

    echo '<table width="100%" class="main_table">';
    echo '<tr>';
        echo '<th>Model No</th>';        
        echo '<th>Product Title</th>';                
        echo '<th>Bin-location</th>';                        
        echo '<th>Category</th>';                        
        echo '<th>Price ($)</th>';                        
        echo '<th>Jobzones</th>';                  
        echo '<th>Result</th>';                  
    echo '</tr>';    
    
    $mapping_array = array();
    $success = array();
    $failed = array();    
    foreach($data as $key => $value) 
    {                    
        if($key == 0) continue;

        //Category & Bin-location adjustment
        $bin_location = trim($value[2]);
        $value[2] = 'Location';
        
        echo '<tr>';
        echo '<td>';            
            echo $value[0];
            $value[0] = trim($value[0]);
        echo '</td>';

        echo '<td>';
            echo $value[1];
            $value[1] = trim($value[1]);
        echo '</td>';

        echo '<td>';
            echo $bin_location;            
        echo '</td>';

        echo '<td>';
            echo $value[2];
            $value[2] = trim(strtolower($value[2]));
            echo isset($validCategories[$value[2]]) ?  ' ('.$validCategories[$value[2]].')' : ' (Invalid)';
        echo '</td>';

        echo '<td>';
            echo $value[3];
        echo '</td>';

        echo '<td>';
            echo $jobzones;
        echo '</td>';
                
        echo '<td>';
                        
        if($execute == 1) {
            $db = JFactory::getDBO();

            $query = $db->getQuery(true);
            $query = "SELECT count(*) as total from `#__jshopping_categories` where `name_en-GB` = '".trim($value[2])."'";      
            $db->setQuery($query);
            $category = $db->loadObject();

            if(isset($validCategories[$value[2]])) 
            { 
                $category_id = $validCategories[$value[2]];
                
                $query = $db->getQuery(true);
                $query = "SELECT MAX(product_ordering) as ordering from `#__jshopping_products_to_categories` where category_id = '".$category_id."' ";      
                $db->setQuery($query);
                $data = $db->loadObject();
                $ordering = $data->ordering + 1;
                
                $query = $db->getQuery(true);
                echo $query = "INSERT into `#__jshopping_products` 
                         (`product_id`, `is_app`, `model_no`, `product_quantity`,
                         `unlimited`, `product_date_added`, `date_modify`,
                         `product_publish`, `product_tax_id`, `currency_id`,
                         `product_jobrole`, `product_template`, `product_price`, `min_price`, 
                         `product_weight`, `product_height`, `product_width`, `product_length`, 
                         `access`, `name_en-GB`, `bin_location`) 
                         VALUES 
                         ('', '1', '".trim($value[0])."', '1',
                          '1', '".date('Y-m-d H:i:s')."', '".date('Y-m-d H:i:s')."',
                          '1', '1', '2',
                          '".$jobzones."', 'default', '".$value[3]."', '".$value[3]."', 
                          '1', '1', '1', '1',
                          '1', '".$db->escape($value[1])."', '".$db->escape($bin_location)."')";      
                $db->setQuery($query);
                $db->query();
                $product_id = $db->insertid();

                if($product_id > 0) {
                    $query = $db->getQuery(true);
                    $query = "INSERT into `#__jshopping_products_to_categories` 
                             (`product_id`, `category_id`, `product_ordering`) 
                             VALUES 
                             ('".$product_id."', '".$category_id."', '".$ordering."')";      
                    $db->setQuery($query);
                    $db->query();

                    $query = $db->getQuery(true);
                    $query = "INSERT into `#__company_product_map` 
                             (`company_id`, `product_id`) 
                             VALUES 
                             ('7', '".$product_id."')";      
                    $db->setQuery($query);
                    $db->query();

                    $success[] = $value[0];
                } else {
                    $failed[] = $value[0];    
                }

            } else {
                $failed[] = $value[0];
            }
            
        }

        echo '</td>';
        echo '</tr>';
        $total_rows++;        
    }

    echo '</table>';

    echo '<br/><br/><br/>';
    echo 'Total Rows :'.$total_rows.'<br/>';	
    echo 'Total Success :'.count($success).'<br/>';    
    echo 'Total Failed :'.count($failed).'<br/>';    
}
?>
