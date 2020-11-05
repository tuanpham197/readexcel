<?php
    //Nhúng file PHPExcel
require_once 'PHPExcel/Classes/PHPExcel.php';

//Đường dẫn file
$file = 'HotSALE_0511/mainpage.xlsx';
//Tiến hành xác thực file
$objFile = PHPExcel_IOFactory::identify($file);
$objData = PHPExcel_IOFactory::createReader($objFile);

//Chỉ đọc dữ liệu
$objData->setReadDataOnly(true);

// Load dữ liệu sang dạng đối tượng
$objPHPExcel = $objData->load($file);



//Chọn trang cần truy xuất
$sheet = $objPHPExcel->setActiveSheetIndex(0);

//Lấy ra số dòng cuối cùng
$Totalrow = $sheet->getHighestRow();
//Lấy ra tên cột cuối cùng
$LastColumn = $sheet->getHighestColumn();

//Chuyển đổi tên cột đó về vị trí thứ, VD: C là 3,D là 4
$TotalCol = PHPExcel_Cell::columnIndexFromString($LastColumn);

//Tạo mảng chứa dữ liệu
$data = [];

//Tiến hành lặp qua từng ô dữ liệu

$tmp = 0;
$str = "";
for ($i = 1; $i <= $Totalrow; $i++) {
    //----Lặp cột
    $tmp++;
    // $star_date = substr($sheet->getCellByColumnAndRow(2, $i)->getValue(),0,4).'-'.substr($sheet->getCellByColumnAndRow(2, $i)->getValue(),4,2).'-'.substr($sheet->getCellByColumnAndRow(2, $i)->getValue(),6,2);
    // $end_date = substr($sheet->getCellByColumnAndRow(3, $i)->getValue(),0,4).'-'.substr($sheet->getCellByColumnAndRow(3, $i)->getValue(),4,2).'-'.substr($sheet->getCellByColumnAndRow(3, $i)->getValue(),6,2);
    // $str .= "('_HOTSALE_".$sheet->getCellByColumnAndRow(2, $i)->getValue().$sheet->getCellByColumnAndRow(1, $i)->getValue()."','".$sheet->getCellByColumnAndRow(0, $i)->getValue()."'),"."</br>";
    $str .= "('".$sheet->getCellByColumnAndRow(0, $i)->getValue()."_HOME_HOTSALE".$sheet->getCellByColumnAndRow(2, $i)->getValue().$sheet->getCellByColumnAndRow(3, $i)->getValue()."','".$sheet->getCellByColumnAndRow(1, $i)->getValue()."'),"."</br>";
}

echo $str;
echo "Tổng số sản phẩm: ".$tmp;
//Hiển thị mảng dữ liệu
echo '<pre>';
var_dump($data);