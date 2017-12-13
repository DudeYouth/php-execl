<?php
function numberToLetter($number){
    if( $number<=26 ){
        $str = chr(64+$number);
    }else{
        $str = chr(64+$number%26);
        $new_number = floor($number/26);
        $str = numberToLetter($new_number).$str;
        
    }
    return $str;
}

echo numberToLetter(58);