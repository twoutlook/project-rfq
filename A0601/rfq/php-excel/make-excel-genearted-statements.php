<?php 


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:310 function: makePercentFormat
//

// makePercentFormat 0.00% ok, FORMAT_PERCENTAGE_00 為什麼不行?
$objPHPExcel->getActiveSheet()->getStyle('C34:H34')->getNumberFormat()->setFormatCode("0.00%");
$objPHPExcel->getActiveSheet()->getStyle('C47:H47')->getNumberFormat()->setFormatCode("0.00%");
$objPHPExcel->getActiveSheet()->getStyle('C57:H57')->getNumberFormat()->setFormatCode("0.00%");
$objPHPExcel->getActiveSheet()->getStyle('C68:H68')->getNumberFormat()->setFormatCode("0.00%");
$objPHPExcel->getActiveSheet()->getStyle('C89:H89')->getNumberFormat()->setFormatCode("0.00%");
$objPHPExcel->getActiveSheet()->getStyle('C106:H106')->getNumberFormat()->setFormatCode("0.00%");


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:249 function: makeSum
//

// RMB

// USD


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:329 function: makeMoneyStyle
//
$objPHPExcel->getActiveSheet()->getStyle('C25:H25')->getNumberFormat()->setFormatCode("$#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C116:H116')->getNumberFormat()->setFormatCode("$#,##0.00");


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:179 function: makeCell32
//

$objPHPExcel->getActiveSheet() 
->setCellValue('C32', '=C30*C31/1000') 
->setCellValue('D32', '=D30*D31/1000') 
->setCellValue('E32', '=E30*E31/1000') 
->setCellValue('F32', '=F30*F31/1000') 
->setCellValue('G32', '=G30*G31/1000') 
->setCellValue('H32', '=H30*H31/1000') 
;

//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:209 function: extendCell34X
//


// --- extendCell34X(34,=100*C31*C16/(C31*C16+C33)/100) ---

$objPHPExcel->getActiveSheet() 
->setCellValue('C34', '=100*C31*C16/(C31*C16+C33)/100')
->setCellValue('D34', '=100*D31*D16/(D31*D16+D33)/100')
->setCellValue('E34', '=100*E31*E16/(E31*E16+E33)/100')
->setCellValue('F34', '=100*F31*F16/(F31*F16+F33)/100')
->setCellValue('G34', '=100*G31*G16/(G31*G16+G33)/100')
->setCellValue('H34', '=100*H31*H16/(H31*H16+H33)/100')
;


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:209 function: extendCell34X
//


// --- extendCell34X(36,=(C30-C35)*C33/1000/C16) ---

$objPHPExcel->getActiveSheet() 
->setCellValue('C36', '=(C30-C35)*C33/1000/C16')
->setCellValue('D36', '=(D30-D35)*D33/1000/D16')
->setCellValue('E36', '=(E30-E35)*E33/1000/E16')
->setCellValue('F36', '=(F30-F35)*F33/1000/F16')
->setCellValue('G36', '=(G30-G35)*G33/1000/G16')
->setCellValue('H36', '=(H30-H35)*H33/1000/H16')
;


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:209 function: extendCell34X
//


// --- extendCell34X(37,=(C31+C33)*C30*0.02/1000/C16) ---

$objPHPExcel->getActiveSheet() 
->setCellValue('C37', '=(C31+C33)*C30*0.02/1000/C16')
->setCellValue('D37', '=(D31+D33)*D30*0.02/1000/D16')
->setCellValue('E37', '=(E31+E33)*E30*0.02/1000/E16')
->setCellValue('F37', '=(F31+F33)*F30*0.02/1000/F16')
->setCellValue('G37', '=(G31+G33)*G30*0.02/1000/G16')
->setCellValue('H37', '=(H31+H33)*H30*0.02/1000/H16')
;


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:209 function: extendCell34X
//


// --- extendCell34X(38,=IF(ISNA(C32+C36+C37),0,(C32+C36+C37))) ---

$objPHPExcel->getActiveSheet() 
->setCellValue('C38', '=IF(ISNA(C32+C36+C37),0,(C32+C36+C37))')
->setCellValue('D38', '=IF(ISNA(D32+D36+D37),0,(D32+D36+D37))')
->setCellValue('E38', '=IF(ISNA(E32+E36+E37),0,(E32+E36+E37))')
->setCellValue('F38', '=IF(ISNA(F32+F36+F37),0,(F32+F36+F37))')
->setCellValue('G38', '=IF(ISNA(G32+G36+G37),0,(G32+G36+G37))')
->setCellValue('H38', '=IF(ISNA(H32+H36+H37),0,(H32+H36+H37))')
;


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:209 function: extendCell34X
//


// --- extendCell34X(43,=3600/C42) ---

$objPHPExcel->getActiveSheet() 
->setCellValue('C43', '=3600/C42')
->setCellValue('D43', '=3600/D42')
->setCellValue('E43', '=3600/E42')
->setCellValue('F43', '=3600/F42')
->setCellValue('G43', '=3600/G42')
->setCellValue('H43', '=3600/H42')
;


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:209 function: extendCell34X
//


// --- extendCell34X(44,=C41/C43 ) ---

$objPHPExcel->getActiveSheet() 
->setCellValue('C44', '=C41/C43 ')
->setCellValue('D44', '=D41/D43 ')
->setCellValue('E44', '=E41/E43 ')
->setCellValue('F44', '=F41/F43 ')
->setCellValue('G44', '=G41/G43 ')
->setCellValue('H44', '=H41/H43 ')
;


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:209 function: extendCell34X
//


// --- extendCell34X(48,=(C44+C45)*(1+(1-C47))/C16) ---

$objPHPExcel->getActiveSheet() 
->setCellValue('C48', '=(C44+C45)*(1+(1-C47))/C16')
->setCellValue('D48', '=(D44+D45)*(1+(1-D47))/D16')
->setCellValue('E48', '=(E44+E45)*(1+(1-E47))/E16')
->setCellValue('F48', '=(F44+F45)*(1+(1-F47))/F16')
->setCellValue('G48', '=(G44+G45)*(1+(1-G47))/G16')
->setCellValue('H48', '=(H44+H45)*(1+(1-H47))/H16')
;


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:209 function: extendCell34X
//


// --- extendCell34X(52,=(C50/3600)*C51) ---

$objPHPExcel->getActiveSheet() 
->setCellValue('C52', '=(C50/3600)*C51')
->setCellValue('D52', '=(D50/3600)*D51')
->setCellValue('E52', '=(E50/3600)*E51')
->setCellValue('F52', '=(F50/3600)*F51')
->setCellValue('G52', '=(G50/3600)*G51')
->setCellValue('H52', '=(H50/3600)*H51')
;


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:209 function: extendCell34X
//


// --- extendCell34X(56,=(C54/3600)*C55) ---

$objPHPExcel->getActiveSheet() 
->setCellValue('C56', '=(C54/3600)*C55')
->setCellValue('D56', '=(D54/3600)*D55')
->setCellValue('E56', '=(E54/3600)*E55')
->setCellValue('F56', '=(F54/3600)*F55')
->setCellValue('G56', '=(G54/3600)*G55')
->setCellValue('H56', '=(H54/3600)*H55')
;


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:209 function: extendCell34X
//


// --- extendCell34X(59,=C56*(1+(1-C57))) ---

$objPHPExcel->getActiveSheet() 
->setCellValue('C59', '=C56*(1+(1-C57))')
->setCellValue('D59', '=D56*(1+(1-D57))')
->setCellValue('E59', '=E56*(1+(1-E57))')
->setCellValue('F59', '=F56*(1+(1-F57))')
->setCellValue('G59', '=G56*(1+(1-G57))')
->setCellValue('H59', '=H56*(1+(1-H57))')
;


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:209 function: extendCell34X
//


// --- extendCell34X(63,=(C61/3600)*C62) ---

$objPHPExcel->getActiveSheet() 
->setCellValue('C63', '=(C61/3600)*C62')
->setCellValue('D63', '=(D61/3600)*D62')
->setCellValue('E63', '=(E61/3600)*E62')
->setCellValue('F63', '=(F61/3600)*F62')
->setCellValue('G63', '=(G61/3600)*G62')
->setCellValue('H63', '=(H61/3600)*H62')
;


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:209 function: extendCell34X
//


// --- extendCell34X(64,=(C61/3600)*C62) ---

$objPHPExcel->getActiveSheet() 
->setCellValue('C64', '=(C61/3600)*C62')
->setCellValue('D64', '=(D61/3600)*D62')
->setCellValue('E64', '=(E61/3600)*E62')
->setCellValue('F64', '=(F61/3600)*F62')
->setCellValue('G64', '=(G61/3600)*G62')
->setCellValue('H64', '=(H61/3600)*H62')
;


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:209 function: extendCell34X
//


// --- extendCell34X(69,=IF(ISNA((C66/3600)*C67 * (1 + (1 - C68))),0,(C66/3600)*C67 * (1 + (1 - C68)))) ---

$objPHPExcel->getActiveSheet() 
->setCellValue('C69', '=IF(ISNA((C66/3600)*C67 * (1 + (1 - C68))),0,(C66/3600)*C67 * (1 + (1 - C68)))')
->setCellValue('D69', '=IF(ISNA((D66/3600)*D67 * (1 + (1 - D68))),0,(D66/3600)*D67 * (1 + (1 - D68)))')
->setCellValue('E69', '=IF(ISNA((E66/3600)*E67 * (1 + (1 - E68))),0,(E66/3600)*E67 * (1 + (1 - E68)))')
->setCellValue('F69', '=IF(ISNA((F66/3600)*F67 * (1 + (1 - F68))),0,(F66/3600)*F67 * (1 + (1 - F68)))')
->setCellValue('G69', '=IF(ISNA((G66/3600)*G67 * (1 + (1 - G68))),0,(G66/3600)*G67 * (1 + (1 - G68)))')
->setCellValue('H69', '=IF(ISNA((H66/3600)*H67 * (1 + (1 - H68))),0,(H66/3600)*H67 * (1 + (1 - H68)))')
;


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:209 function: extendCell34X
//


// --- extendCell34X(73,=IF(ISNA((C71/3600)*C72),0,(C71/3600)*C72)) ---

$objPHPExcel->getActiveSheet() 
->setCellValue('C73', '=IF(ISNA((C71/3600)*C72),0,(C71/3600)*C72)')
->setCellValue('D73', '=IF(ISNA((D71/3600)*D72),0,(D71/3600)*D72)')
->setCellValue('E73', '=IF(ISNA((E71/3600)*E72),0,(E71/3600)*E72)')
->setCellValue('F73', '=IF(ISNA((F71/3600)*F72),0,(F71/3600)*F72)')
->setCellValue('G73', '=IF(ISNA((G71/3600)*G72),0,(G71/3600)*G72)')
->setCellValue('H73', '=IF(ISNA((H71/3600)*H72),0,(H71/3600)*H72)')
;


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:209 function: extendCell34X
//


// --- extendCell34X(77,=IF(ISNA(C76),0,(C75/3600)*C76)) ---

$objPHPExcel->getActiveSheet() 
->setCellValue('C77', '=IF(ISNA(C76),0,(C75/3600)*C76)')
->setCellValue('D77', '=IF(ISNA(D76),0,(D75/3600)*D76)')
->setCellValue('E77', '=IF(ISNA(E76),0,(E75/3600)*E76)')
->setCellValue('F77', '=IF(ISNA(F76),0,(F75/3600)*F76)')
->setCellValue('G77', '=IF(ISNA(G76),0,(G75/3600)*G76)')
->setCellValue('H77', '=IF(ISNA(H76),0,(H75/3600)*H76)')
;


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:209 function: extendCell34X
//


// --- extendCell34X(83,=C80*C81*C82) ---

$objPHPExcel->getActiveSheet() 
->setCellValue('C83', '=C80*C81*C82')
->setCellValue('D83', '=D80*D81*D82')
->setCellValue('E83', '=E80*E81*E82')
->setCellValue('F83', '=F80*F81*F82')
->setCellValue('G83', '=G80*G81*G82')
->setCellValue('H83', '=H80*H81*H82')
;


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:209 function: extendCell34X
//


// --- extendCell34X(91,=C86*(C87+C88)*(1+(1-C89))*C90) ---

$objPHPExcel->getActiveSheet() 
->setCellValue('C91', '=C86*(C87+C88)*(1+(1-C89))*C90')
->setCellValue('D91', '=D86*(D87+D88)*(1+(1-D89))*D90')
->setCellValue('E91', '=E86*(E87+E88)*(1+(1-E89))*E90')
->setCellValue('F91', '=F86*(F87+F88)*(1+(1-F89))*F90')
->setCellValue('G91', '=G86*(G87+G88)*(1+(1-G89))*G90')
->setCellValue('H91', '=H86*(H87+H88)*(1+(1-H89))*H90')
;


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:209 function: extendCell34X
//


// --- extendCell34X(95,=C94) ---

$objPHPExcel->getActiveSheet() 
->setCellValue('C95', '=C94')
->setCellValue('D95', '=D94')
->setCellValue('E95', '=E94')
->setCellValue('F95', '=F94')
->setCellValue('G95', '=G94')
->setCellValue('H95', '=H94')
;


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:209 function: extendCell34X
//


// --- extendCell34X(99,=C98) ---

$objPHPExcel->getActiveSheet() 
->setCellValue('C99', '=C98')
->setCellValue('D99', '=D98')
->setCellValue('E99', '=E98')
->setCellValue('F99', '=F98')
->setCellValue('G99', '=G98')
->setCellValue('H99', '=H98')
;


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:209 function: extendCell34X
//


// --- extendCell34X(104,=C102+C103) ---

$objPHPExcel->getActiveSheet() 
->setCellValue('C104', '=C102+C103')
->setCellValue('D104', '=D102+D103')
->setCellValue('E104', '=E102+E103')
->setCellValue('F104', '=F102+F103')
->setCellValue('G104', '=G102+G103')
->setCellValue('H104', '=H102+H103')
;


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:209 function: extendCell34X
//


// --- extendCell34X(107,=C105*C106) ---

$objPHPExcel->getActiveSheet() 
->setCellValue('C107', '=C105*C106')
->setCellValue('D107', '=D105*D106')
->setCellValue('E107', '=E105*E106')
->setCellValue('F107', '=F105*F106')
->setCellValue('G107', '=G105*G106')
->setCellValue('H107', '=H105*H106')
;


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:209 function: extendCell34X
//


// --- extendCell34X(110,=C108+C109) ---

$objPHPExcel->getActiveSheet() 
->setCellValue('C110', '=C108+C109')
->setCellValue('D110', '=D108+D109')
->setCellValue('E110', '=E108+E109')
->setCellValue('F110', '=F108+F109')
->setCellValue('G110', '=G108+G109')
->setCellValue('H110', '=H108+H109')
;


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:372 function: makeColorFillStyle
//


// --- makeColorFillStyle(A, {"A9BCF5":[15,28,39,50,54,61,66,71,75,78,85,96,100,104]})---
$objPHPExcel->getActiveSheet()->getStyle('A15')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('A28')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('A39')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('A50')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('A54')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('A61')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('A66')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('A71')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('A75')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('A78')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('A85')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('A96')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('A100')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('A104')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:372 function: makeColorFillStyle
//


// --- makeColorFillStyle(B, {"A9BCF5":[15,28,39,50,54,61,66,71,75,78,85,96,100,104]})---
$objPHPExcel->getActiveSheet()->getStyle('B15')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('B28')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('B39')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('B50')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('B54')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('B61')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('B66')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('B71')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('B75')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('B78')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('B85')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('B96')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('B100')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('B104')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:372 function: makeColorFillStyle
//


// --- makeColorFillStyle(B, {"BE6E6E6":[39,49,53,60,65,70,74,78,84,95,99,103,108]})---
$objPHPExcel->getActiveSheet()->getStyle('B39')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('B49')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('B53')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('B60')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('B65')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('B70')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('B74')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('B78')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('B84')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('B95')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('B99')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('B103')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('B108')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:372 function: makeColorFillStyle
//


// --- makeColorFillStyle(C, {"BE6E6E6":[39,49,53,60,65,70,74,78,84,95,99,103,108]})---
$objPHPExcel->getActiveSheet()->getStyle('C39')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('C49')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('C53')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('C60')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('C65')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('C70')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('C74')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('C78')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('C84')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('C95')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('C99')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('C103')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('C108')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:372 function: makeColorFillStyle
//


// --- makeColorFillStyle(D, {"BE6E6E6":[39,49,53,60,65,70,74,78,84,95,99,103,108]})---
$objPHPExcel->getActiveSheet()->getStyle('D39')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('D49')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('D53')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('D60')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('D65')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('D70')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('D74')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('D78')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('D84')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('D95')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('D99')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('D103')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('D108')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:372 function: makeColorFillStyle
//


// --- makeColorFillStyle(E, {"BE6E6E6":[39,49,53,60,65,70,74,78,84,95,99,103,108]})---
$objPHPExcel->getActiveSheet()->getStyle('E39')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('E49')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('E53')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('E60')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('E65')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('E70')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('E74')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('E78')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('E84')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('E95')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('E99')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('E103')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('E108')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:372 function: makeColorFillStyle
//


// --- makeColorFillStyle(F, {"BE6E6E6":[39,49,53,60,65,70,74,78,84,95,99,103,108]})---
$objPHPExcel->getActiveSheet()->getStyle('F39')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('F49')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('F53')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('F60')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('F65')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('F70')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('F74')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('F78')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('F84')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('F95')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('F99')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('F103')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('F108')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:372 function: makeColorFillStyle
//


// --- makeColorFillStyle(G, {"BE6E6E6":[39,49,53,60,65,70,74,78,84,95,99,103,108]})---
$objPHPExcel->getActiveSheet()->getStyle('G39')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('G49')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('G53')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('G60')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('G65')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('G70')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('G74')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('G78')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('G84')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('G95')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('G99')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('G103')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('G108')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:372 function: makeColorFillStyle
//


// --- makeColorFillStyle(H, {"BE6E6E6":[39,49,53,60,65,70,74,78,84,95,99,103,108]})---
$objPHPExcel->getActiveSheet()->getStyle('H39')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('H49')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('H53')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('H60')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('H65')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('H70')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('H74')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('H78')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('H84')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('H95')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('H99')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('H103')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('H108')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:372 function: makeColorFillStyle
//


// --- makeColorFillStyle(C, {"F9E79F":[10,12,30,41,66,71,75,80,86]})---
$objPHPExcel->getActiveSheet()->getStyle('C10')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('C12')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('C30')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('C41')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('C66')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('C71')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('C75')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('C80')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('C86')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:372 function: makeColorFillStyle
//


// --- makeColorFillStyle(D, {"F9E79F":[10,12,30,41,66,71,75,80,86]})---
$objPHPExcel->getActiveSheet()->getStyle('D10')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('D12')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('D30')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('D41')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('D66')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('D71')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('D75')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('D80')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('D86')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:372 function: makeColorFillStyle
//


// --- makeColorFillStyle(E, {"F9E79F":[10,12,30,41,66,71,75,80,86]})---
$objPHPExcel->getActiveSheet()->getStyle('E10')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('E12')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('E30')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('E41')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('E66')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('E71')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('E75')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('E80')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('E86')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:372 function: makeColorFillStyle
//


// --- makeColorFillStyle(F, {"F9E79F":[10,12,30,41,66,71,75,80,86]})---
$objPHPExcel->getActiveSheet()->getStyle('F10')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('F12')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('F30')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('F41')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('F66')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('F71')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('F75')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('F80')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('F86')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:372 function: makeColorFillStyle
//


// --- makeColorFillStyle(G, {"F9E79F":[10,12,30,41,66,71,75,80,86]})---
$objPHPExcel->getActiveSheet()->getStyle('G10')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('G12')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('G30')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('G41')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('G66')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('G71')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('G75')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('G80')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('G86')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:372 function: makeColorFillStyle
//


// --- makeColorFillStyle(H, {"F9E79F":[10,12,30,41,66,71,75,80,86]})---
$objPHPExcel->getActiveSheet()->getStyle('H10')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('H12')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('H30')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('H41')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('H66')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('H71')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('H75')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('H80')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('H86')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:372 function: makeColorFillStyle
//


// --- makeColorFillStyle(C, {"F4D03F":[11,16]})---
$objPHPExcel->getActiveSheet()->getStyle('C11')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C16')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:372 function: makeColorFillStyle
//


// --- makeColorFillStyle(D, {"F4D03F":[11,16]})---
$objPHPExcel->getActiveSheet()->getStyle('D11')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D16')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:372 function: makeColorFillStyle
//


// --- makeColorFillStyle(E, {"F4D03F":[11,16]})---
$objPHPExcel->getActiveSheet()->getStyle('E11')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E16')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:372 function: makeColorFillStyle
//


// --- makeColorFillStyle(F, {"F4D03F":[11,16]})---
$objPHPExcel->getActiveSheet()->getStyle('F11')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F16')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:372 function: makeColorFillStyle
//


// --- makeColorFillStyle(G, {"F4D03F":[11,16]})---
$objPHPExcel->getActiveSheet()->getStyle('G11')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G16')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:372 function: makeColorFillStyle
//


// --- makeColorFillStyle(H, {"F4D03F":[11,16]})---
$objPHPExcel->getActiveSheet()->getStyle('H11')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H16')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:372 function: makeColorFillStyle
//


// --- makeColorFillStyle(A, {"cccccc":[24,25,109,115,116,117]})---
$objPHPExcel->getActiveSheet()->getStyle('A24')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('A25')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('A109')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('A115')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('A116')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('A117')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:372 function: makeColorFillStyle
//


// --- makeColorFillStyle(B, {"cccccc":[24,25,109,115,116,117]})---
$objPHPExcel->getActiveSheet()->getStyle('B24')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('B25')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('B109')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('B115')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('B116')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('B117')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:372 function: makeColorFillStyle
//


// --- makeColorFillStyle(C, {"cccccc":[24,25,109,115,116,117]})---
$objPHPExcel->getActiveSheet()->getStyle('C24')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('C25')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('C109')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('C115')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('C116')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('C117')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:372 function: makeColorFillStyle
//


// --- makeColorFillStyle(D, {"cccccc":[24,25,109,115,116,117]})---
$objPHPExcel->getActiveSheet()->getStyle('D24')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('D25')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('D109')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('D115')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('D116')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('D117')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:372 function: makeColorFillStyle
//


// --- makeColorFillStyle(E, {"cccccc":[24,25,109,115,116,117]})---
$objPHPExcel->getActiveSheet()->getStyle('E24')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('E25')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('E109')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('E115')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('E116')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('E117')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:372 function: makeColorFillStyle
//


// --- makeColorFillStyle(F, {"cccccc":[24,25,109,115,116,117]})---
$objPHPExcel->getActiveSheet()->getStyle('F24')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('F25')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('F109')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('F115')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('F116')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('F117')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:372 function: makeColorFillStyle
//


// --- makeColorFillStyle(G, {"cccccc":[24,25,109,115,116,117]})---
$objPHPExcel->getActiveSheet()->getStyle('G24')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('G25')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('G109')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('G115')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('G116')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('G117')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');


//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:372 function: makeColorFillStyle
//


// --- makeColorFillStyle(H, {"cccccc":[24,25,109,115,116,117]})---
$objPHPExcel->getActiveSheet()->getStyle('H24')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('H25')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('H109')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('H115')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('H116')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('H117')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');


//[[A0601]] 顯示版本信息

//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:372 function: makeColorFillStyle
//


// --- makeColorFillStyle(A, {"98AFC7":[7,8,9]})---
$objPHPExcel->getActiveSheet()->getStyle('A7')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('98AFC7');
$objPHPExcel->getActiveSheet()->getStyle('A8')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('98AFC7');
$objPHPExcel->getActiveSheet()->getStyle('A9')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('98AFC7');


//[[A0601]] 加方格線

//
// file:D:\xampp\htdocs\project-rfq\A0601\rfq\php-excel\make-excel-helper.php line:350 function: makeCellsBorderColRowFromTo
//


$styleArr = array( 'borders' => array( 'allborders' => array( 'style' => 'thin' )));
$objPHPExcel->getActiveSheet()->getStyle('A3')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A4')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A5')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A6')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A7')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A8')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A9')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A10')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A11')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A12')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A13')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A14')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A15')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A16')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A17')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A18')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A19')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A20')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A21')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A22')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A23')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A24')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A25')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A26')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A27')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A28')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A29')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A30')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A31')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A32')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A33')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A34')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A35')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A36')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A37')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A38')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A39')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A40')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A41')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A42')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A43')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A44')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A45')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A46')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A47')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A48')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A49')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A50')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A51')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A52')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A53')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A54')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A55')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A56')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A57')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A58')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A59')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A60')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A61')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A62')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A63')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A64')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A65')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A66')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A67')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A68')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A69')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A70')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A71')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A72')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A73')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A74')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A75')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A76')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A77')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A78')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A79')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A80')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A81')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A82')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A83')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A84')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A85')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A86')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A87')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A88')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A89')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A90')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A91')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A92')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A93')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A94')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A95')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A96')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A97')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A98')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A99')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A100')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A101')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A102')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A103')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A104')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A105')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A106')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A107')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A108')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A109')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A110')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A111')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A112')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A113')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A114')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A115')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A116')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A117')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A118')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('A119')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B3')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B4')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B5')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B6')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B7')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B8')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B9')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B10')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B11')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B12')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B13')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B14')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B15')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B16')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B17')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B18')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B19')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B20')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B21')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B22')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B23')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B24')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B25')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B26')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B27')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B28')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B29')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B30')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B31')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B32')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B33')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B34')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B35')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B36')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B37')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B38')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B39')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B40')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B41')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B42')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B43')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B44')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B45')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B46')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B47')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B48')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B49')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B50')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B51')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B52')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B53')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B54')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B55')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B56')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B57')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B58')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B59')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B60')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B61')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B62')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B63')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B64')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B65')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B66')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B67')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B68')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B69')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B70')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B71')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B72')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B73')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B74')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B75')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B76')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B77')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B78')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B79')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B80')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B81')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B82')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B83')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B84')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B85')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B86')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B87')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B88')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B89')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B90')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B91')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B92')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B93')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B94')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B95')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B96')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B97')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B98')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B99')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B100')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B101')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B102')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B103')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B104')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B105')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B106')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B107')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B108')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B109')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B110')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B111')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B112')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B113')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B114')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B115')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B116')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B117')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B118')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('B119')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C3')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C4')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C5')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C6')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C7')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C8')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C9')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C10')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C11')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C12')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C13')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C14')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C15')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C16')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C17')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C18')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C19')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C20')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C21')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C22')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C23')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C24')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C25')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C26')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C27')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C28')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C29')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C30')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C31')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C32')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C33')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C34')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C35')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C36')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C37')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C38')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C39')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C40')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C41')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C42')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C43')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C44')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C45')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C46')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C47')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C48')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C49')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C50')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C51')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C52')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C53')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C54')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C55')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C56')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C57')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C58')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C59')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C60')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C61')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C62')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C63')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C64')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C65')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C66')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C67')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C68')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C69')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C70')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C71')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C72')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C73')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C74')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C75')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C76')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C77')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C78')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C79')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C80')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C81')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C82')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C83')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C84')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C85')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C86')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C87')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C88')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C89')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C90')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C91')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C92')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C93')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C94')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C95')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C96')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C97')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C98')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C99')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C100')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C101')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C102')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C103')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C104')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C105')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C106')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C107')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C108')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C109')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C110')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C111')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C112')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C113')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C114')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C115')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C116')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C117')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C118')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('C119')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E3')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E4')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E5')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E6')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E7')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E8')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E9')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E10')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E11')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E12')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E13')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E14')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E15')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E16')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E17')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E18')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E19')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E20')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E21')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E22')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E23')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E24')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E25')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E26')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E27')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E28')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E29')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E30')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E31')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E32')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E33')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E34')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E35')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E36')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E37')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E38')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E39')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E40')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E41')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E42')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E43')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E44')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E45')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E46')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E47')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E48')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E49')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E50')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E51')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E52')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E53')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E54')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E55')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E56')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E57')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E58')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E59')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E60')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E61')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E62')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E63')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E64')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E65')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E66')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E67')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E68')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E69')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E70')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E71')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E72')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E73')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E74')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E75')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E76')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E77')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E78')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E79')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E80')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E81')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E82')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E83')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E84')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E85')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E86')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E87')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E88')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E89')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E90')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E91')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E92')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E93')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E94')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E95')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E96')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E97')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E98')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E99')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E100')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E101')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E102')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E103')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E104')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E105')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E106')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E107')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E108')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E109')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E110')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E111')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E112')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E113')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E114')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E115')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E116')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E117')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E118')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('E119')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D3')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D4')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D5')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D6')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D7')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D8')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D9')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D10')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D11')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D12')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D13')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D14')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D15')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D16')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D17')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D18')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D19')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D20')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D21')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D22')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D23')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D24')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D25')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D26')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D27')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D28')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D29')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D30')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D31')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D32')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D33')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D34')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D35')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D36')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D37')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D38')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D39')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D40')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D41')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D42')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D43')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D44')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D45')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D46')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D47')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D48')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D49')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D50')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D51')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D52')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D53')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D54')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D55')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D56')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D57')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D58')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D59')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D60')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D61')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D62')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D63')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D64')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D65')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D66')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D67')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D68')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D69')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D70')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D71')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D72')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D73')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D74')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D75')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D76')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D77')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D78')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D79')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D80')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D81')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D82')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D83')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D84')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D85')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D86')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D87')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D88')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D89')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D90')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D91')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D92')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D93')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D94')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D95')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D96')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D97')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D98')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D99')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D100')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D101')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D102')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D103')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D104')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D105')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D106')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D107')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D108')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D109')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D110')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D111')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D112')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D113')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D114')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D115')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D116')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D117')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D118')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('D119')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F3')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F4')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F5')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F6')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F7')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F8')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F9')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F10')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F11')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F12')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F13')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F14')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F15')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F16')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F17')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F18')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F19')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F20')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F21')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F22')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F23')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F24')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F25')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F26')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F27')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F28')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F29')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F30')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F31')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F32')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F33')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F34')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F35')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F36')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F37')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F38')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F39')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F40')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F41')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F42')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F43')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F44')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F45')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F46')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F47')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F48')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F49')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F50')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F51')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F52')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F53')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F54')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F55')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F56')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F57')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F58')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F59')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F60')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F61')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F62')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F63')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F64')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F65')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F66')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F67')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F68')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F69')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F70')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F71')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F72')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F73')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F74')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F75')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F76')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F77')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F78')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F79')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F80')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F81')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F82')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F83')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F84')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F85')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F86')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F87')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F88')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F89')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F90')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F91')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F92')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F93')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F94')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F95')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F96')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F97')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F98')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F99')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F100')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F101')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F102')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F103')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F104')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F105')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F106')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F107')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F108')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F109')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F110')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F111')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F112')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F113')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F114')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F115')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F116')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F117')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F118')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('F119')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G3')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G4')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G5')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G6')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G7')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G8')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G9')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G10')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G11')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G12')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G13')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G14')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G15')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G16')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G17')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G18')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G19')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G20')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G21')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G22')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G23')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G24')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G25')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G26')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G27')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G28')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G29')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G30')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G31')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G32')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G33')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G34')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G35')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G36')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G37')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G38')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G39')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G40')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G41')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G42')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G43')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G44')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G45')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G46')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G47')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G48')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G49')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G50')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G51')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G52')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G53')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G54')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G55')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G56')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G57')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G58')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G59')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G60')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G61')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G62')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G63')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G64')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G65')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G66')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G67')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G68')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G69')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G70')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G71')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G72')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G73')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G74')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G75')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G76')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G77')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G78')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G79')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G80')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G81')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G82')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G83')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G84')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G85')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G86')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G87')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G88')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G89')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G90')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G91')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G92')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G93')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G94')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G95')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G96')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G97')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G98')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G99')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G100')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G101')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G102')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G103')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G104')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G105')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G106')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G107')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G108')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G109')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G110')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G111')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G112')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G113')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G114')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G115')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G116')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G117')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G118')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('G119')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H3')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H4')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H5')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H6')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H7')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H8')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H9')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H10')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H11')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H12')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H13')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H14')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H15')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H16')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H17')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H18')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H19')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H20')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H21')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H22')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H23')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H24')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H25')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H26')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H27')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H28')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H29')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H30')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H31')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H32')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H33')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H34')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H35')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H36')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H37')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H38')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H39')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H40')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H41')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H42')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H43')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H44')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H45')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H46')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H47')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H48')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H49')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H50')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H51')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H52')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H53')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H54')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H55')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H56')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H57')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H58')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H59')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H60')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H61')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H62')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H63')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H64')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H65')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H66')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H67')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H68')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H69')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H70')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H71')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H72')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H73')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H74')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H75')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H76')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H77')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H78')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H79')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H80')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H81')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H82')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H83')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H84')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H85')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H86')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H87')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H88')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H89')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H90')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H91')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H92')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H93')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H94')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H95')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H96')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H97')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H98')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H99')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H100')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H101')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H102')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H103')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H104')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H105')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H106')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H107')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H108')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H109')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H110')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H111')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H112')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H113')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H114')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H115')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H116')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H117')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H118')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H119')->applyFromArray($styleArr);