<?php
/**
 * Created by PhpStorm.
 * User: Andrey
 * Date: 18.12.2015
 * Time: 1:51
 */

require_once 'Classes/PHPExcel.php';

$objPHPExcel = new PHPExcel();

/* Настройки */
$objPHPExcel->setActiveSheetIndex(0);
$active_sheet = $objPHPExcel->getActiveSheet();
$active_sheet->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
$active_sheet->getPageSetup()->SetPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);

$active_sheet->setTitle("Футбольная АПЛ");

$active_sheet->getColumnDimension('A')->setWidth(12);
$active_sheet->getColumnDimension('B')->setWidth(12);
$active_sheet->getColumnDimension('C')->setWidth(15);
$active_sheet->getColumnDimension('D')->setWidth(15);
$active_sheet->getColumnDimension('E')->setWidth(15);

$active_sheet->mergeCells('A1:C1');
$active_sheet->setCellValue('A1', 'API Football');
$active_sheet->setCellValue('A2','Match date');
$active_sheet->setCellValue('B2','Match time');
$active_sheet->setCellValue('C2','Home team');
$active_sheet->setCellValue('D2','Guest team');
$active_sheet->setCellValue('E2','Account');
$active_sheet->setCellValue('F2','Author goal');

//массив стилей
$style_header = array(
    //Шрифт
    'font'=>array(
        'bold' => true,
        'name' => 'Times New Roman',
        'size' => 12
    ),
    //Заполнение цветом
    'fill' => array(
        'type' => PHPExcel_STYLE_FILL::FILL_SOLID,
        'color'=>array(
            'rgb' => 'CCC'
        )
    )
);

$active_sheet->getStyle('A2:F2')->applyFromArray($style_header);


$row_start = 3;
$i = 0;

// Заполняем ячейки данными
foreach($player as $item) {

    $row_next = $row_start + $i;

    $active_sheet->setCellValue('A'.$row_next, $item['match_formatted_date']);
    $active_sheet->setCellValue('B'.$row_next, $item['match_time']);
    $active_sheet->setCellValue('C'.$row_next, $item['match_localteam_name']);
    $active_sheet->setCellValue('D'.$row_next, $item['match_visitorteam_name']);
    $active_sheet->setCellValue('E'.$row_next, $item['match_ft_score']);

    $str = '';

    foreach($item['match_events'] as $events) {
        if($events['event_type'] == 'goal') {
            $str.= 'GOAL!!! '.$events['event_minute']. 'min, ' . $events['event_player'] . '|';
        }
    }


    $active_sheet->setCellValue('F'.$row_next, $str);

    $i++;
}

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');