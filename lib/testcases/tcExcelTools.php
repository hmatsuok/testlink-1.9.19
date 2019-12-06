<?php
require_once "../../vendor/vendor/autoload.php";


function downloadExcelToFile($content,$export_filename){
    $spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
    $sheet = $spreadsheet->getSheet(0);

    exportheaderExcel($sheet);

    $row = 3;
    $xml = @simplexml_load_string($content);
    if($xml !== FALSE) {
        if ($xml->getName() == 'testcases'){
            $xmlTCs = $xml->xpath('//testcase');
            exportTestCaseExcel($xmlTCs,$row,$sheet);
        }
        if ($xml->getName() == 'testsuite'){
            $xmlTSs = $xml->xpath('//testsuite');
            foreach ($xmlTSs as $el_TS){
                $TS_name = $el_TS['name'];
                $TS_detail = stripHTML($el_TS->details[0]);
                $sheet->setCellValue('A'.$row, $TS_name);
                $sheet->setCellValue('B'.$row, $TS_detail);
                $xmlTCs = $el_TS->xpath('testcase');
                exportTestCaseExcel($xmlTCs,$row,$sheet);
            }
        }
    }

    $writer = new PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
    $filename = '/var/www/html/testlink/upload_area/'.$export_filename;
    $writer->save($filename);

    ob_get_clean();
    header('Content-Description: File Transfer');
    header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; name=".basename($filename));
    header('Content-Transfer-Encoding: BASE64;' );
    header('Content-Disposition: attachment; filename="'.basename($filename).'"');
    readfile($filename); // send file
    unlink($filename);

}
function stripHTML($in_html){
    $out_html = $in_html;
    $out_html = str_replace('&nbsp;',' ',$out_html);
    $out_html = str_replace('&quot;','"',$out_html);
    $out_html = str_replace('<p>','',$out_html);
    $out_html = str_replace('</p>','',$out_html);
    return $out_html;
}
function exportheaderExcel($sheet){
    $sheet->getColumnDimension('A')->setWidth(20);
    $sheet->getColumnDimension('B')->setWidth(20);
    $sheet->getColumnDimension('C')->setWidth(20);
    $sheet->getColumnDimension('D')->setWidth(20);
    $sheet->getColumnDimension('E')->setWidth(20);
    $sheet->getColumnDimension('F')->setWidth(10);
    $sheet->getColumnDimension('G')->setWidth(8);
    $sheet->getColumnDimension('H')->setWidth(8);
    $sheet->getColumnDimension('I')->setWidth(10);
    $sheet->getColumnDimension('J')->setWidth(20);
    $sheet->getColumnDimension('K')->setWidth(40);
    $sheet->getColumnDimension('L')->setWidth(40);

    $style_head = array(
        'borders' => array(
            'outline' => array(
                'borderStyle' => PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN
            )
        ),
        'alignment' => array(
            'horizontal' => PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
        ),
        'fill' => array(
            'fillType' => PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
            'color' => array(
                'argb' => '92D050'
            ),
        ),
    );
    $sheet->setCellValue('A1', "テストスイート");
    $sheet->mergeCells('A1:B1');
    $sheet->getStyle('A1:B1')->applyFromArray($style_head);
    $sheet->setCellValue('A2', "名前");
    $sheet->getStyle('A2')->applyFromArray($style_head);
    $sheet->setCellValue('B2', "詳細");
    $sheet->getStyle('B2')->applyFromArray($style_head);
    $sheet->setCellValue('C1', "テストケース");
    $sheet->mergeCells('C1:J1');
    $sheet->getStyle('C1:I1')->applyFromArray($style_head);
    $sheet->setCellValue('C2', "名前");
    $sheet->getStyle('C2')->applyFromArray($style_head);
    $sheet->setCellValue('D2', "詳細");
    $sheet->getStyle('D2')->applyFromArray($style_head);
    $sheet->setCellValue('E2', "前提条件");
    $sheet->getStyle('E2')->applyFromArray($style_head);
    $sheet->setCellValue('F2', "ステータス");
    $sheet->getStyle('F2')->applyFromArray($style_head);
    $sheet->setCellValue('G2', "重要度");
    $sheet->getStyle('G2')->applyFromArray($style_head);
    $sheet->setCellValue('H2', "実行タイプ");
    $sheet->getStyle('H2')->applyFromArray($style_head);
    $sheet->setCellValue('I2', "予定実行時間");
    $sheet->getStyle('I2')->applyFromArray($style_head);
    $sheet->setCellValue('J2', "キーワード");
    $sheet->getStyle('J2')->applyFromArray($style_head);
    $sheet->setCellValue('K1', "ステップ");
    $sheet->mergeCells('K1:L1');
    $sheet->getStyle('K1:L1')->applyFromArray($style_head);
    $sheet->setCellValue('K2', "ステップ");
    $sheet->getStyle('K2')->applyFromArray($style_head);
    $sheet->setCellValue('L2', "期待効果");
    $sheet->getStyle('L2')->applyFromArray($style_head);

}
function exportTestCaseExcel($xmlTCs,&$row,$sheet){
    $list_status = [
        '1'=>'ドラフト',
        '2'=>'レビュー待ち',
        '3'=>'レビュー中',
        '4'=>'やり直し',
        '5'=>'廃止',
        '6'=>'先送り',
        '7'=>'完了',
    ];
    $list_importance = [
        '1'=>'高',
        '2'=>'中',
        '3'=>'低',
    ];
    $list_execution_type = [
        '1'=>'手動',
        '2'=>'自動',
    ];
    foreach ($xmlTCs as $el_TC){
        $TC_name = $el_TC['name'];
        $TC_summary = stripHTML($el_TC->summary[0]);
        $TC_preconditions = stripHTML($el_TC->preconditions[0]);
        $TC_status = (string)$el_TC->status[0];
        $TC_status_str = $list_status[$TC_status];
        $TC_importance = (string)$el_TC->importance[0];
        $TC_importance_str = $list_importance[$TC_importance];
        $TC_execution_type = (string)$el_TC->execution_type[0];
        $TC_execution_type_str = $list_execution_type[$TC_execution_type];
        $TC_estimated_exec_duration = $el_TC->estimated_exec_duration[0];
        $sheet->setCellValue('C'.$row, $TC_name);
        $sheet->setCellValue('D'.$row, $TC_summary);
        $sheet->setCellValue('E'.$row, $TC_preconditions);
        $sheet->setCellValue('F'.$row, $TC_status_str);
        $sheet->setCellValue('G'.$row, $TC_importance_str);
        $sheet->setCellValue('H'.$row, $TC_execution_type_str);
        $sheet->setCellValue('I'.$row, $TC_estimated_exec_duration);
        cellStyle($sheet,$row);
        $TC_Keywords = '';
        $xmlKeywords = $el_TC->xpath('keywords');
        foreach ($xmlKeywords as $el_keywords){
            foreach ($el_keywords as $el_keyword){
                if ($TC_Keywords != ''){
                    $TC_Keywords .= "\n";
                }
                $TC_Keywords .= $el_keyword['name'];
            }
        }
        $sheet->setCellValue('J'.$row, $TC_Keywords);
        $xmlSteps = $el_TC->xpath('steps');
        foreach ($xmlSteps as $el_Steps){
            $xmlStep = $el_Steps->xpath('step');
            foreach ($xmlStep as $el_Step){
                $Step_action = stripHTML($el_Step->actions[0]);
                $Step_result = stripHTML($el_Step->expectedresults[0]);
                $sheet->setCellValue('K'.$row, $Step_action);
                $sheet->setCellValue('L'.$row, $Step_result);
                cellStyle($sheet,$row);
                $row++;
            }
        }
    }

}
function cellStyle($sheet,$row){
    $style_cell = array(
        'borders' => array(
            'outline' => array(
                'borderStyle' => PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN
            )
        ),
        'alignment' => array(
            'horizontal' => PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
            'vertical' => PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP,
            'wrapText' => true
        ),
    );
    $sheet->getStyle('A'.$row)->applyFromArray($style_cell);
    $sheet->getStyle('B'.$row)->applyFromArray($style_cell);
    $sheet->getStyle('C'.$row)->applyFromArray($style_cell);
    $sheet->getStyle('D'.$row)->applyFromArray($style_cell);
    $sheet->getStyle('E'.$row)->applyFromArray($style_cell);
    $sheet->getStyle('F'.$row)->applyFromArray($style_cell);
    $sheet->getStyle('G'.$row)->applyFromArray($style_cell);
    $sheet->getStyle('H'.$row)->applyFromArray($style_cell);
    $sheet->getStyle('I'.$row)->applyFromArray($style_cell);
    $sheet->getStyle('J'.$row)->applyFromArray($style_cell);
    $sheet->getStyle('K'.$row)->applyFromArray($style_cell);
    $sheet->getStyle('L'.$row)->applyFromArray($style_cell);
}
function importTestCaseDataFromEXCEL(&$db,$fileName,$parentID,$tproject_id,$userID,$options=null)
{
    tLog('importTestCaseDataFromEXCEL called for file: '. $fileName);
    $file_check = array('status_ok' => 0, 'msg' => 'xml_load_ko');

    $list_status = [
        'ドラフト' => '1',
        'レビュー待ち' => '2',
        'レビュー中' => '3',
        'やり直し' => '4',
        '廃止' => '5',
        '先送り' => '6',
        '完了' => '7',
    ];
    $list_importance = [
        '高' => '1',
        '中' => '2',
        '低' => '3',
    ];
    $list_execution_type = [
        '手動' => '1',
        '自動' => '2',
    ];

    if (file_exists($fileName)) {
        $reader = new PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $spreadsheet = $reader->load($fileName);
        $sheet = $spreadsheet->getActiveSheet();
        $highestRow = $sheet->getHighestRow();

        $xml = '<?xml version="1.0" encoding="UTF-8"?>'.PHP_EOL;

        $sw_testsuite = false;
        $head1 = $sheet->getCell("A1")->getValue();
        $head1a = $sheet->getCell("A2")->getValue();
        $head1b = $sheet->getCell("A6")->getValue();
        if (($head1 == 'テストスイート名') && ($head1a == 'テストストート要約') && ($head1b == '入力項目/設定値')){
            for ($s = 0; $s < $spreadsheet->getSheetCount(); $s++){
                $sheet = $spreadsheet->getSheet($s);
                $highestRow = $sheet->getHighestRow();

                $TS_name = $sheet->getCell("B1")->getValue();
                $TS_detail = $sheet->getCell("B2")->getValue();
                $TS_detail = convText($TS_detail);
                $TC_name = $sheet->getCell("B3")->getValue();
                $TC_summary = $sheet->getCell("B4")->getValue();
                $TC_summary = convText($TC_summary);
                $TC_preconditions = $sheet->getCell("B5")->getValue();
                $TC_preconditions = convText($TC_preconditions);

                if (!$sw_testsuite){
                    $xml .= '<testsuite name="'.$TS_name.'">'.PHP_EOL;
                    $xml .= '<details><![CDATA['.$TS_detail.']]></details>'.PHP_EOL;
                    $sw_testsuite = true;
                }

                $TC_execution_type_num = 1;
                $TC_importance_num = 2;
                $TC_status_num = 1;
                $TC_CellName = [];
                $TC_Action = [];
                $TC_Result = [];
                $max_x = PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($sheet->getHighestColumn());
                for ($x = 0; $x < ($max_x - 5); $x++){
                    $cell_name = PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($x+5);
                    $testNum = $sheet->getCell($cell_name."6")->getValue();
                    if ($testNum == ''){
                        break;
                    }
                    $TC_CellName[] = $cell_name;
                    $TC_Action[$cell_name] = '';
                    $TC_Result[$cell_name] = '';
                }

                $label_a = '';
                $label_b = '';
                for ($row = 7; $row <= $highestRow; $row++){
                    $value_a = $sheet->getCell("A".$row)->getValue();
                    $value_b = $sheet->getCell("B".$row)->getValue();
                    $value_c = $sheet->getCell("C".$row)->getValue();
                    $value_d = $sheet->getCell("D".$row)->getValue();
                    if ($value_a == 'アクション'){
                        $label_a = $value_a;
                    }
                    else if ($value_a == '期待する結果'){
                        $label_a = $value_a;
                    }
                    if ($value_b != ''){
                        $label_b = $value_b;
                    }

                    if ($label_a == 'アクション'){
                        foreach ($TC_CellName as $cell_name){
                            $mark = $sheet->getCell($cell_name.$row)->getValue();
                            if ($mark == '○'){
                                if ($TC_Action[$cell_name] != ''){
                                    $TC_Action[$cell_name] .= PHP_EOL;
                                }
                                $TC_Action[$cell_name] .= '<p>'.$label_b.'　'.$value_d.'</p>';
                            }
                        }
                    }
                    else if ($label_a == '期待する結果'){
                        foreach ($TC_CellName as $cell_name){
                            $mark = $sheet->getCell($cell_name.$row)->getValue();
                            if ($mark == '○'){
                                if ($TC_Result[$cell_name] != ''){
                                    $TC_Result[$cell_name] .= PHP_EOL;
                                }
                                $TC_Result[$cell_name] .= '<p>'.$label_b.'　'.$value_d.'</p>';
                            }
                        }
                    }
                }
                $xml .= '<testcase name="'.$TC_name.'">'.PHP_EOL;
                $xml .= '<summary><![CDATA['.$TC_summary.']]></summary>'.PHP_EOL;
                $xml .= '<preconditions><![CDATA['.$TC_preconditions.']]></preconditions>'.PHP_EOL;
                $xml .= '<execution_type><![CDATA['.$TC_execution_type_num.']]></execution_type>'.PHP_EOL;
                $xml .= '<importance><![CDATA['.$TC_importance_num.']]></importance>'."\n";
                $xml .= '<estimated_exec_duration>'.'</estimated_exec_duration>'.PHP_EOL;
                $xml .= '<status>'.$TC_status_num.'</status>'.PHP_EOL;
                $xml .= '<is_open>1</is_open>'.PHP_EOL;
                $xml .= '<active>1</active>'.PHP_EOL;
                $xml .= '<steps>'.PHP_EOL;
                $step_num = 1;
                foreach ($TC_CellName as $cell_name){
                    $xml .= '<step>'.PHP_EOL;
                    $xml .= '<step_number><![CDATA['.$step_num.']]></step_number>'.PHP_EOL;
                    $xml .= '<actions><![CDATA['.$TC_Action[$cell_name].']]></actions>'.PHP_EOL;
                    $xml .= '<expectedresults><![CDATA['.$TC_Result[$cell_name].']]></expectedresults>'.PHP_EOL;
                    $xml .= '<execution_type><![CDATA['.$TC_execution_type_num.']]></execution_type>'.PHP_EOL;
                    $xml .= '</step>'.PHP_EOL;
                    $step_num++;
                }
                $xml .= '</steps>'.PHP_EOL;
                $xml .= '</testcase>'.PHP_EOL;

            }
            if ($sw_testsuite){
                $xml .= '</testsuite>'.PHP_EOL;
            }


        }else{
            $sw_TS = false;
            $sw_TC = false;
            $step_num = 0;
            for ($row = 3; $row <= $highestRow; $row++){
                $TS_name = $sheet->getCell("A".$row)->getValue();
                $TS_detail = $sheet->getCell("B".$row)->getValue();
                $TS_detail = convText($TS_detail);
                $TC_name = $sheet->getCell("C".$row)->getValue();
                $TC_summary = $sheet->getCell("D".$row)->getValue();
                $TC_summary = convText($TC_summary);
                $TC_preconditions = $sheet->getCell("E".$row)->getValue();
                $TC_preconditions = convText($TC_preconditions);
                $TC_status = $sheet->getCell("F".$row)->getValue();
                $TC_status_num = $list_status[$TC_status];
                $TC_importance = $sheet->getCell("G".$row)->getValue();
                $TC_importance_num = $list_importance[$TC_importance];
                $TC_execution_type = $sheet->getCell("H".$row)->getValue();
                $TC_execution_type_num = $list_execution_type[$TC_execution_type];
                $TC_estimated = $sheet->getCell("I".$row)->getValue();
                $TC_keyword = $sheet->getCell("J".$row)->getValue();
                $Step_action = $sheet->getCell("K".$row)->getValue();
                $Step_action = convText($Step_action);
                $Step_result = $sheet->getCell("L".$row)->getValue();
                $Step_result = convText($Step_result);
                if ($TS_name != ''){
                    if ($sw_TS){
                        $xml .= '</testsuite>'.PHP_EOL;
                    }
                    $sw_TS = true;
                    $sw_TC = false;
                    $xml .= '<testsuite name="'.$TS_name.'">'.PHP_EOL;
                    $xml .= '<details><![CDATA['.$TS_detail.']]></details>'.PHP_EOL;
                }
                if ($TC_name != ''){
                    if ($sw_TC){
                        $xml .= '</steps>'.PHP_EOL;
                        $xml .= '</testcase>'.PHP_EOL;
                    }
                    $sw_TC = true;
                    $xml .= '<testcase name="'.$TC_name.'">'.PHP_EOL;
                    $xml .= '<summary><![CDATA['.$TC_summary.']]></summary>'.PHP_EOL;
                    $xml .= '<preconditions><![CDATA['.$TC_preconditions.']]></preconditions>'.PHP_EOL;
                    $xml .= '<execution_type><![CDATA['.$TC_execution_type_num.']]></execution_type>'.PHP_EOL;
                    $xml .= '<importance><![CDATA['.$TC_importance_num.']]></importance>'."\n";
                    $xml .= '<estimated_exec_duration>'.$TC_estimated.'</estimated_exec_duration>'.PHP_EOL;
                    $xml .= '<status>'.$TC_status_num.'</status>'.PHP_EOL;
                    $xml .= '<is_open>1</is_open>'.PHP_EOL;
                    $xml .= '<active>1</active>'.PHP_EOL;
                    if ($TC_keyword != ''){
                        $TC_keyword_arr = explode("\n",$TC_keyword);
                        $xml .= '<keywords>'.PHP_EOL;
                        foreach ($TC_keyword_arr as $line_k){
                            $xml .= '<keyword name="'.$line_k.'"><notes><![CDATA[]]></notes></keyword>'.PHP_EOL;
                        }
                        $xml .= '</keywords>'.PHP_EOL;
                    }
                    $xml .= '<steps>'.PHP_EOL;
                    $step_num = 1;
                }
                if ($Step_action != ''){
                    $xml .= '<step>'.PHP_EOL;
                    $xml .= '<step_number><![CDATA['.$step_num.']]></step_number>'.PHP_EOL;
                    $xml .= '<actions><![CDATA['.$Step_action.']]></actions>'.PHP_EOL;
                    $xml .= '<expectedresults><![CDATA['.$Step_result.']]></expectedresults>'.PHP_EOL;
                    $xml .= '<execution_type><![CDATA['.$TC_execution_type_num.']]></execution_type>'.PHP_EOL;
                    $xml .= '</step>'.PHP_EOL;
                    $step_num++;
                }
            }

            if ($sw_TC){
                $xml .= '</steps>'.PHP_EOL;
                $xml .= '</testcase>'.PHP_EOL;
            }
            if ($sw_TS){
                $xml .= '</testsuite>'.PHP_EOL;
            }
        }


        $fileName_xml = str_replace('.xlsx','.xml',$fileName);
        file_put_contents($fileName_xml, $xml);
        $file_check = array('status_ok' => 1, 'msg' => 'ok');
    }

    return $file_check;
}
function convText($in_str){
    if ($in_str == null){
        return '';
    }
    $in_arr = explode("\n",$in_str);
    $out_str = '';
    $cnt = count($in_arr);
    $ii = 0;
    foreach ($in_arr as $line_x){
        $ii++;
        if ($line_x == ''){
            if ($ii < $cnt){
                $out_str .= PHP_EOL;
            }
        }else{
            $out_str .= "<p>".htmlspecialchars($line_x, ENT_QUOTES)."</p>".PHP_EOL;
        }
    }
    return $out_str;
}
function check_excel_tc_tsuite($fileName,$recursiveMode)
{
    $list_status = [
        'ドラフト' => '1',
        'レビュー待ち' => '2',
        'レビュー中' => '3',
        'やり直し' => '4',
        '廃止' => '5',
        '先送り' => '6',
        '完了' => '7',
    ];
    $list_importance = [
        '高' => '1',
        '中' => '2',
        '低' => '3',
    ];
    $list_execution_type = [
        '手動' => '1',
        '自動' => '2',
    ];

    $file_check = array('status_ok' => 0, 'msg' => 'Excelファイルの書式が誤っています');
    if (file_exists($fileName)) {
        $reader = new PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $spreadsheet = $reader->load($fileName);
        $sheet = $spreadsheet->getActiveSheet();
        $head1 = $sheet->getCell("A1")->getValue();
        $head1a = $sheet->getCell("A2")->getValue();
        $head1b = $sheet->getCell("A6")->getValue();
        $head2 = $sheet->getCell("C1")->getValue();
        $head3 = $sheet->getCell("K1")->getValue();
        if (($head1 == 'テストスイート名') && ($head1a == 'テストストート要約') && ($head1b == '入力項目/設定値')){
            $test_1 = $sheet->getCell("E6")->getValue();
            if ($test_1 != '1'){
                $file_check = array('status_ok' => 0, 'msg' => 'ディシジョン形式になっていません');
                return $file_check;
            }
            $file_check = array('status_ok' => 1, 'msg' => 'ok');
        }
        else if (($head1 == 'テストスイート') && ($head2 == 'テストケース') && ($head3 == 'ステップ')){
            $highestRow = $sheet->getHighestRow();
            if ($highestRow > 2){
                for ($row = 3; $row <= $highestRow; $row++){
                    $TC_name = $sheet->getCell("C".$row)->getValue();
                    $TC_status = $sheet->getCell("F".$row)->getValue();
                    $TC_importance = $sheet->getCell("G".$row)->getValue();
                    $TC_execution_type = $sheet->getCell("H".$row)->getValue();
                    if ($TC_name != ''){
                        if ($TC_status == ''){
                            $file_check = array('status_ok' => 0, 'msg' => 'status is blank');
                            return $file_check;
                        }
                        if ($TC_importance == ''){
                            $file_check = array('status_ok' => 0, 'msg' => 'importance is blank');
                            return $file_check;
                        }
                        if ($TC_execution_type == ''){
                            $file_check = array('status_ok' => 0, 'msg' => 'execution_type is blank');
                            return $file_check;
                        }
                        if (!array_key_exists($TC_status, $list_status)){
                            $file_check = array('status_ok' => 0, 'msg' => 'status not valid');
                            return $file_check;
                        }
                        if (!array_key_exists($TC_importance, $list_importance)){
                            $file_check = array('status_ok' => 0, 'msg' => 'importance not valid');
                            return $file_check;
                        }
                        if (!array_key_exists($TC_execution_type, $list_execution_type)){
                            $file_check = array('status_ok' => 0, 'msg' => 'execution_type not valid');
                            return $file_check;
                        }
                    }
                }
                $file_check = array('status_ok' => 1, 'msg' => 'ok');
            }
        }
    }
    return $file_check;
}

