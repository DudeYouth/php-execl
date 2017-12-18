<?php
require 'PHPExcel\Classes\PHPExcel.php';

$arr = [
    [
        'title' =>'测试sheet',
        'style'=>[

        ],
        'header'=>[
            'fields'=>[
                [
                'value'=>"字段名",
                'field'=>'name',
            ],
            [
                'value'=>"字段名1",
                'field'=>'name1',
            ],
            [
                'value'=>"字段名2",
                'field'=>'name2',
            ]],
            'style'=>[

            ],
        ],
        'body'=>[
                'data'=>[
                    
                        [
                            'name'=>[
                                'value'=>"字段名",
                                'fontBold'=>true,
                            ],
                            'name1'=>[
                                'value'=>"字段名8",
                                'fontBold'=>true,
                            ],
                            'name2'=>[
                                'value'=>"字段名9",
                                'fontBold'=>true,
                            ],
                        ],
                        [
                            'name'=>[
                                'value'=>"字段名",
                                'fontBold'=>true,
                            ],
                            'name1'=>[
                                'value'=>"字段名8",
                                'fontBold'=>true,
                            ],
                            'name2'=>[
                                'value'=>"字段名9",
                                'fontBold'=>true,
                            ],
                        ],
                        [
                            'name'=>[
                                'value'=>"字段名",
                                'fontBold'=>true,
                            ],
                            'name1'=>[
                                'value'=>"字段名8",
                                'fontBold'=>true,
                            ],
                            'name2'=>[
                                'value'=>"字段名9",
                                'fontBold'=>true,
                            ],
                        ]
                    
                ],
                'style'=>[

                ],    
        ],

    ]
];
class Execl{
    // 表格的列id生成
    private function numberToLetter($number){
        if( $number<=26 ){
            $str = chr(64+$number);
        }else{
            $str = chr(64+$number%26);
            $new_number = floor($number/26);
            $str = $this ->numberToLetter($new_number).$str;
            
        }
        return $str;
    }
    // 设置字体
    private function setRowCellStyle($sheet,$fun,$method,$row,$ncell,$v){
            $cell= $this ->numberToLetter($ncell);
            
            $obj = $sheet ->getStyle($cell.$row)->$fun(); 
            $this ->list_exec_function($obj,$method,$v);

    }
    // 遍历执行方法
    private function list_exec_function($obj,$method,$args){
        if( is_array($method) ){
            foreach( $method as $key=>$fun ){
                $obj = $obj ->$fun($args[$key]);
            }
        }else{
            $obj ->$method($args);
        }
    }
    // 设置垂直方向布局
    private function setVertical($sheet,$row,$ncell,$v){
        $value = PHPExcel_Style_Alignment::VERTICAL_JUSTIFY;
        switch( $v ){
            case 'top':
                $value = PHPExcel_Style_Alignment::VERTICAL_TOP;
                break;
            case 'bottom':
                $value = PHPExcel_Style_Alignment::VERTICAL_BOTTOM;
                break;
            case 'center':
                $value = PHPExcel_Style_Alignment::VERTICAL_CENTER;
                break;
        }
        $this ->setRowCellStyle($sheet,'getAlignment','setVertical',$row,$ncell,$value);
        
    }
    // 设置水平方向布局
    private function setHorizontal($sheet,$row,$ncell,$v){
        $value = PHPExcel_Style_Alignment::HORIZONTAL_JUSTIFY;
        switch( $v ){
            case 'right':
                $value = PHPExcel_Style_Alignment::HORIZONTAL_RIGHT;
                break;
            case 'left':
                $value = PHPExcel_Style_Alignment::HORIZONTAL_LEFT;
                break;
            case 'center':
                $value = PHPExcel_Style_Alignment::VERTICAL_CENTER;
                break;
            case 'justify':
                $value = PHPExcel_Style_Alignment::HORIZONTAL_JUSTIFY;
                break;
        }
        $this ->setRowCellStyle($sheet,'getAlignment','setHorizontal',$row,$ncell,$value,$type);
    }
    // 设置样式
    public function setStyle($sheet,$row,$ncell,$key,$v){
        $cell = $this ->numberToLetter($ncell);
        $row = $row+1;
        switch( $key ){
            case 'cell':
                if( is_array($v) ){
                    $x = $v[0]-1;
                    $y = $v[1]-1;
                }else{
                    $x = $v[0];
                    $y = 0;
                }
                $mcell= $this ->numberToLetter($ncell+$x);
                $sheet ->mergeCells($cell.$row.':'.$mcell.($row+$y));//合并单元格
                break;
            case 'fontBold':
                $this ->setRowCellStyle($sheet,'getFont','setBold',$row,$ncell,$v);
                break;
            case 'fontSize':
                $this ->setRowCellStyle($sheet,'getFont','setSize',$row,$ncell,$v);
                break;
            case 'fontFamily':
                $this ->setRowCellStyle($sheet,'getFont','setName',$row,$ncell,$v);
                break;
            case 'fontColor':
                $this ->setRowCellStyle($sheet,'getFont',['getColor','setColor'],$row,$ncell,[$v]);
            case 'height':
                $sheet ->getRowDimension($row)->setRowHeight($v);
                break;
            case 'horizontal':
                $this ->setHorizontal($sheet,$row,$ncell,$v);
                break;
            case 'vertical':
                $this ->setVertical($sheet,$row,$ncell,$v);
                break;
            case 'borderTop':
                    $border = $sheet->getBorders()->getTop();
                    $border ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                    $border ->getColor()->setARGB($v); 
                break;
            case 'borderLeft':
                $border = $sheet->getBorders()->getLeft();
                $border ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $border ->getColor()->setARGB($v); 
                break;
            case 'borderRight':
                $border = $sheet->getBorders()->getRight();
                $border ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $border ->getColor()->setARGB($v); 
                break;
            case 'borderBottom':
                $border = $sheet->getBorders()->getBottom();
                $border ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $border ->getColor()->setARGB($v); 
                break;
            case 'width':
                if( $v=='auto' ){
                    $sheet ->getColumnDimension($cell)->setAutoSize(true);
                }else{
                    $sheet ->getColumnDimension($cell)->setWidth($v);
                }
                break;
            case 'value':
                $sheet ->setCellValue($cell.$row,$v );
                break;

        }
        return $sheet;
    }
    private function merge($origin,$sub){
        if( empty($origin)&&empty($sub) ){
            return [];
        }
        if( !$origin ){
            $origin = [];
        }
        if( !$sub ){
            $sub = [];
        }
        return array_merge($origin,$sub);
    }
    private function setValue($sheet,$data,$row,$cell){
        foreach( $data as $key=>$value ){
            $this ->setStyle($sheet,$row,$cell,$key,$value);
        }
    }
    public function create($arr){
        $objPHPExcel = new \PHPExcel();
        // 遍历sheet
        foreach( $arr as $index=>$data ){
            $objPHPExcel->createSheet();   
            $objPHPExcel->setActiveSheetIndex($index);  
            $sheet = $objPHPExcel->getActiveSheet();
            $sheet ->setTitle($data['title']);
            $styles = $data['style'];
            $header_style = $this ->merge($styles,$data['header']['style']);
            $header_index= 0;
            $this ->keys = [];
            // 头部标题设置
            foreach( $data['header']['fields'] as $hk => $hv ){
                if( gettype($hv)=='string' ){
                    $value_arr = $header_style;
                    $value_arr['value'] = $hv;
                }else{
                    $value_arr = $this ->merge($header_style,$hv);
                }
                $this ->keys[] = $hv['field'];
                $header_index++;
                $this ->setValue($sheet,$value_arr,0,$header_index);
            }
            $body_style = $this ->merge($styles,$data['body']['style']);
            // 内容设置
            foreach( $data['body']['data'] as $key=>$value ){
                //echo json_encode([$body_style,$value['style']]);die;
                if( is_array($value) ){
                    if( isset($value['fields']) ){
                        $row_styles =  $this ->merge($body_style,$value['style']);
                        $fields = $value['fields'];
                    }else{
                        $row_styles = $body_style;
                        $fields = $value;
                    }
                    foreach( $this ->keys as $k =>$v ){
                        $vv = $fields[$v];
                        if( gettype($vv)=='string' ){
                            $value_arr = $row_styles;
                            $value_arr['value'] = $vv;
                        }else{
                            $value_arr = $this ->merge($row_styles,$vv);
                        }
                        $this ->setValue($sheet,$value_arr,$key+1,$k+1);
                    }
                }

            }
        }
        return $objPHPExcel;
 
    }
    public function echoExecl($expTitle,$data){
        $xlsTitle = iconv('utf-8', 'gb2312', $expTitle);//文件名称 
        $fileName = $xlsTitle;//$_SESSION['account'].date('_YmdHis');//or $xlsTitle 文件名称可根据自己情况设定

        $execl = $this ->create($data);
        header('pragma:public');
        header('Content-type:application/vnd.ms-excel;charset=utf-8;name="'.$xlsTitle.'.xls"');
        header("Content-Disposition:attachment;filename=$fileName.xls");//attachment新窗口打印inline本窗口打印      
        $objWriter = \PHPExcel_IOFactory::createWriter($execl, 'Excel5');  
        $objWriter->save('php://output'); 
        exit;  
    }


}
$e = new Execl();
echo $e ->echoExecl('test01',$arr);
?>