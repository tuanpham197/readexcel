<?php
    //Nhúng file PHPExcel
require_once 'PHPExcel/Classes/PHPExcel.php';

//Đường dẫn file
$file = 'HOTSALE_2210/hotsale.xlsx';
//Tiến hành xác thực file
$objFile = PHPExcel_IOFactory::identify($file);
$objData = PHPExcel_IOFactory::createReader($objFile);

//Chỉ đọc dữ liệu
$objData->setReadDataOnly(true);

// Load dữ liệu sang dạng đối tượng
$objPHPExcel = $objData->load($file);

//Lấy ra số trang sử dụng phương thức getSheetCount();
// Lấy Ra tên trang sử dụng getSheetNames();

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
//----Lặp dòng, Vì dòng đầu là tiêu đề cột nên chúng ta sẽ lặp giá trị từ dòng 2

$tmp = 0;
$str = "";
for ($i = 1; $i <= $Totalrow; $i++) {
    //----Lặp cột
    $tmp++;
    // $star_date = substr($sheet->getCellByColumnAndRow(2, $i)->getValue(),0,4).'-'.substr($sheet->getCellByColumnAndRow(2, $i)->getValue(),4,2).'-'.substr($sheet->getCellByColumnAndRow(2, $i)->getValue(),6,2);
    // $end_date = substr($sheet->getCellByColumnAndRow(3, $i)->getValue(),0,4).'-'.substr($sheet->getCellByColumnAndRow(3, $i)->getValue(),4,2).'-'.substr($sheet->getCellByColumnAndRow(3, $i)->getValue(),6,2);
    $str .= "('_HOTSALE_".$sheet->getCellByColumnAndRow(2, $i)->getValue().$sheet->getCellByColumnAndRow(1, $i)->getValue()."','".$sheet->getCellByColumnAndRow(0, $i)->getValue()."'),"."</br>";
    //$str .= "('HOME_HOTSALE_2','".$sheet->getCellByColumnAndRow(0, $i)->getValue()."'),"."</br>";
}
echo $str;
echo $tmp;
//Hiển thị mảng dữ liệu
echo '<pre>';
var_dump($data);

echo "123";