<?php
ini_set('display_errors', 1);
ini_set('display_startup_errors', 1);
error_reporting(E_ALL);
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
// phpinfo();
// $link = mysqli_connect('mysql', 'root', 'root');
// if (!$link) {
//     die('Ошибка соединения: ' . mysqli_error());
// }
// echo 'Успешно соединились';
// mysqli_close($link);
function ReadData()
{
    try {
        $conn = new PDO("sqlsrv:Server=tcp:nameserver,port;Database=base", "username", "password");
        $conn->setAttribute( PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION );  
        $conn->setAttribute( PDO::SQLSRV_ATTR_QUERY_TIMEOUT, 1 );
        } catch (PDOException $e) {
            print "Error!: " . $e->getMessage();
            die();
        }      
    $query = "SELECT * FROM dbo.rpLit";
    $spreadsheet = new Spreadsheet();  
    $spreadsheet->getDefaultStyle()->getFont()->setName('Times New Roman');
    $spreadsheet->getDefaultStyle()->getFont()->setSize(12);
    $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(50);
    $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(30);   
    $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(30); 
    $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(50); 
    $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(15);  
    $spreadsheet->getActiveSheet()->getCell('A1')->getStyle()->getFont()->setBold(true);
    $spreadsheet->getActiveSheet()->getCell('B1')->getStyle()->getFont()->setBold(true);
    $spreadsheet->getActiveSheet()->getCell('C1')->getStyle()->getFont()->setBold(true);
    $spreadsheet->getActiveSheet()->getCell('D1')->getStyle()->getFont()->setBold(true);
    $spreadsheet->getActiveSheet()->getCell('E1')->getStyle()->getFont()->setBold(true);
    $spreadsheet->getActiveSheet()->getCell('A1')->getStyle()->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
    $spreadsheet->getActiveSheet()->getCell('B1')->getStyle()->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
    $spreadsheet->getActiveSheet()->getCell('C1')->getStyle()->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
    $spreadsheet->getActiveSheet()->getCell('D1')->getStyle()->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
    $spreadsheet->getActiveSheet()->getCell('E1')->getStyle()->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
    $sheet = $spreadsheet->getActiveSheet();;   
    $sheet->setCellValue('A1', 'Заглавие');
    $sheet->setCellValue('B1', 'Назначение');
    $sheet->setCellValue('C1', 'Авторы');
    $sheet->setCellValue('D1', 'Издательство');
    $sheet->setCellValue('E1', 'Год издания');
    $stmt = $conn->query( $query );  
    $count = 2;
    while ( $row = $stmt->fetch( PDO::FETCH_ASSOC ) ){  
        $sheet->setCellValue('A'.$count.'', ''.$row['litName'].'');
        $spreadsheet->getActiveSheet()->getCell('A'.$count.'')->getStyle()->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->getCell('A'.$count.'')->getStyle()->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP);
        $spreadsheet->getActiveSheet()->getCell('A'.$count.'')->getStyle()->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);
        $sheet->setCellValue('B'.$count.'', ''.$row['nameProlong'].'');
        $spreadsheet->getActiveSheet()->getCell('B'.$count.'')->getStyle()->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP);
        $spreadsheet->getActiveSheet()->getCell('B'.$count.'')->getStyle()->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);
        $sheet->setCellValue('C'.$count.'', ''.$row['authors'].'');
        $spreadsheet->getActiveSheet()->getCell('C'.$count.'')->getStyle()->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->getCell('C'.$count.'')->getStyle()->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP);
        $spreadsheet->getActiveSheet()->getCell('C'.$count.'')->getStyle()->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);
        $sheet->setCellValue('D'.$count.'', ''.$row['publishing'].'');
        $spreadsheet->getActiveSheet()->getCell('D'.$count.'')->getStyle()->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->getCell('D'.$count.'')->getStyle()->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP);
        $spreadsheet->getActiveSheet()->getCell('D'.$count.'')->getStyle()->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);
        $sheet->setCellValue('E'.$count.'', ''.$row['imprintDate'].'');
        $spreadsheet->getActiveSheet()->getCell('E'.$count.'')->getStyle()->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP);
        $spreadsheet->getActiveSheet()->getCell('E'.$count.'')->getStyle()->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);
        $count++;
    }
    $writer = new Xlsx($spreadsheet); //создаем
    $writer->save('Список библиотекаря.xlsx'); //сохраняем
}
ReadData();