<?php
require 'PHPExcel\Classes\PHPExcel.php';

$arr = [
    [
        'title' =>'测试sheet',
        'data'=>[
                    [
                        'rows'=>[
                            [
                                'value'=>"字段名",
                            ],
                            [
                                'value'=>"字段名1",
                            ],
                            [
                                'value'=>"字段名2",
                            ],
                        ],
                        'height'=>'30',
                        'fontSize'=>'30'
                    ],
                    [
                        'rows'=>[
                            [
                                'value'=>"测试",
                            ],
                            [
                                'value'=>"测试1",
                            ],
                            [
                                'value'=>"测试2",
                            ],
                        ],
                        'height'=>'30'
                    ],     
                ]
    ]
];
class Execl{
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
    private function setFont($sheet,$method,$row,$ncell,$v,$type){
        if( $type=='row' ){
            for( $i=1;$i<=$ncell;$i++ ){
                $cell= $this ->numberToLetter($i);
                $sheet ->getStyle($cell.$row)->getFont()->$method($v); 
            }
        }elseif( $type=='cell' ){
            $cell= $this ->numberToLetter($ncell);
            for( $i=1;$i<=$row;$i++ ){
                $sheet ->getStyle($cell.$i)->getFont()->$method($v); 
            }
        }else{
            $cell= $this ->numberToLetter($ncell);
            $sheet ->getStyle($cell.$row)->getFont()->$method($v); 
        }
    }
    public function setStyle($sheet,$row,$ncell,$key,$v,$type=''){
        $cell = $this ->numberToLetter($ncell);
        $row = $row+1;
        switch( $key ){
            case 'cell':
                $mcell= $this ->numberToLetter($ncell+$v);
                $sheet ->mergeCells($cell.$row.':'.$mcell.$row);//合并单元格
                break;
            case 'row':
                $sheet ->mergeRows($cell.$row.':'.$cell.($key+$v));//合并单元格
                break;
            case 'fontBold':
                $this ->setFont($sheet,'setBold',$row,$ncell,$v,$type);
            case 'fontSize':
                $this ->setFont($sheet,'setSize',$row,$ncell,$v,$type);
            case 'fontFamily':
                $this ->setFont($sheet,'setName',$row,$ncell,$v,$type);
                break;
            case 'height':
                $sheet ->getRowDimension($row)->setRowHeight($v);
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
    public function create($arr){
        $objPHPExcel = new \PHPExcel();
        // 遍历sheet
        foreach( $arr as $index=>$data ){
            $objPHPExcel->createSheet();   
            $objPHPExcel->setActiveSheetIndex($index);  
            $sheet = $objPHPExcel->getActiveSheet();
            $sheet ->setTitle($data['title']);
            // 遍历行
            foreach( $data['data'] as $key=>$value ){
                // 设置行的样式
                foreach( $value as $k=>$v ){
                    $this ->setStyle($sheet,$key,count($value['rows']),$k,$v,'row');
                }
                // 遍历列
                foreach( $value['rows'] as $cell=>$vv ){
                    // 设置列的样式
                    foreach( $vv as $kkk=>$vvv ){
                        $this ->setStyle($sheet,$key,$cell+1,$kkk,$vvv);
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