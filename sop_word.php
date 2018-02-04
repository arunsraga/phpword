<?php
include_once 'Sample_Header.php';

echo "filname is = ";
echo "$argv[1]\n";
echo "destination filname is = ";
echo "$argv[2]\n";

$input_file  = $argv[1];
$output_file = $argv[2];

$json = file_get_contents($input_file);
//Decode JSON
$json_data = json_decode($json,true);
$header_counter = 0;

// Content variables
$sop_data = $json_data["sop"];

$departments = implode(", ",$sop_data["departments"]);
echo $departments;
//$sop_data["departments"] = implode(",", $sop_data["departments"]); 
$procedure = $sop_data["procedures"]; 
$roles = $sop_data["roles_responsibility"];
$references = $sop_data["references"];
$annexures = $sop_data["annexures"];

// Local resources
$resources = '/opt/app/drl_sop_backend/phpword/resources/';

// New Word Document
echo date('H:i:s'), ' Create new PhpWord object', EOL;
$phpWord = new \PhpOffice\PhpWord\PhpWord();

//$section1 = $phpWord->addSection();

$section1 = $phpWord->addSection(
	array('marginLeft' => 50, 'marginRight' => 600,
	 'marginTop' => 600, 'marginBottom' => 600)
  );

// New portrait section
$section = $phpWord->addSection(array('marginTop' => 2500, 'headerHeight'=>100, 'marginBottom' => 200));


// Add first page header
$header = $section->addHeader();
//$header->firstPage();
$section1->addImage($resources . 'first_page_log.png', 
	/* array('width' => 200, 'height' => 50,
		 'posHorizontal'    => \PhpOffice\PhpWord\Style\Image::POSITION_HORIZONTAL_LEFT,
		 'posHorizontalRel' => \PhpOffice\PhpWord\Style\Image::POS_RELTO_LMARGIN,
		 'marginLeft'       => \PhpOffice\PhpWord\Shared\Converter::cmToPixel(2),
		 'marginTop'        => \PhpOffice\PhpWord\Shared\Converter::cmToPixel(-10),
	)*/
	array('width' => 200, 'height' => 30,
		 //'marginLeft'		=> 100,
		 'marginRight'		=> 200,
		 'marginTop'		=> -50,
		 //'marginBottom'		=> 30,
		 //'alignment' 		=> \PhpOffice\PhpWord\SimpleType\Jc::START,
		 //'positioning'      => \PhpOffice\PhpWord\Style\Image::POSITION_RELATIVE,
		 //'posHorizontal'    => \PhpOffice\PhpWord\Style\Image::POSITION_HORIZONTAL_LEFT,
         'posHorizontalRel' => \PhpOffice\PhpWord\Style\Image::POS_RELTO_LMARGIN,
        //'posVerticalRel'   => \PhpOffice\PhpWord\Style\Image::POS_RELTO_TMARGIN,
        //'posVerticalTop'   => \PhpOffice\PhpWord\Style\Image::POS_RELTO_TMARGIN,
        'marginLeft'       => \PhpOffice\PhpWord\Shared\Converter::cmToPixel(5),
        //'marginRight'       => \PhpOffice\PhpWord\Shared\Converter::cmToPixel(5),
        //'marginBottom'       => 10,
       // 'marginTop'        => \PhpOffice\PhpWord\Shared\Converter::cmToPixel(1.55),
    )
);

/*$table = $section1->addTable(array("borderSize" => 1, "borderColor" => "FFFFFF"));
$table->addRow();
$table->addCell(4500)->addImage($resources . 'first_page_log.png', array('width'=>200, 'height'=>50, 'align'=>'left'));*/
 
//$header->addWatermark('/opt/drl_sop_backend/phpword/resources/TLTCF.png', array('marginTop' => 500, 'marginLeft' => 55));
$fancyTableStyle = array('borderSize' => 1, 'borderColor' => '999999', 'exactHeight' => false);
$fancyTableStyle1 = array('borderSize' => 1, 'borderColor' => '999999', 'line'=>'dash', 'exactHeight' => false);
$gspgTable = array('borderSize' => 1, 'borderColor' => '999999', 'cellMarginTop'=> 100, 'cellMarginBottom'=> 100, 'cellMarginRight'=> 100, 'cellMarginLeft'=> 100);
$infoList = array('name'=>'Arial', 'size'=>'10', 'color'=>'000', 'italic'=> true);
$gspgList = array('listType'=>\PhpOffice\PhpWord\Style\ListItem::TYPE_BULLET_FILLED);
$cellRowSpan = array('vMerge' => 'restart', 'valign' => 'center');
$cellRowSpan1 = array('vMerge' => 'restart', 'valign' => 'center', 'bgcolor'=>'6337ae');
$cellRowContinue = array('vMerge' => 'continue');
$cellColSpan = array('gridSpan' => 2, 'valign' => 'center');
$cellHCentered = array('alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER);
$cellVCentered = array('valign' => 'center', 'color'=> '666');
$cellVCenteredHeader = array('valign' => 'center', 'color'=> '666');
$titlePagecell1 = array('valign' => 'center', 'bgColor'=>'6337ae');
$noSpace = array('spaceAfter' => 0);
$headingStyle_num = array('name'=>'Arial', 'size'=>'16', 'color'=>'5225B5', 'bold'=> true);
$headingStyle = array('name'=>'Arial', 'size'=>'16', 'color'=>'5225B5', 'bold'=> true,
					  'underline'=> \PhpOffice\PhpWord\Style\Font:: UNDERLINE_SINGLE );
					  
$subHeadingStyle = array('name'=>'Arial', 'size'=>'14', 'color'=>'5225B5', 'bold'=> true);
$subHeadingStyle_l2 = array('name'=>'Arial', 'size'=>'14', 'color'=>'5225B5', 'bold'=> true);
$subTitleStyle = array('name'=>'Arial', 'size'=>'12', 'color'=>'000', 'italic'=> true, 'bold'=> true);
$normalTextTyle = array('name'=>'Arial', 'size'=>'12', 'color'=>'000000', 'bold'=> false);
$normalTextStyle_cell = array('name'=>'Arial', 'size'=>'12', 'color'=>'000000', 'bold'=> false);
$gspg = array('name'=>'Arial', 'size'=>'12', 'color'=>'000000', 'bgColor' => 'FAF9FA','bold'=> true);
// E9E9E9
$HeadingparagraphStyle = array('spaceBefore'=>300, 'spaceAfter'=>200);
$subHeadingparagraphStyle = array('spaceBefore'=>150, 'spaceAfter'=>100);
$contentparagraphStyle = array('spaceBefore'=>100, 'spaceAfter'=>100);

$indentStyle = array('indentation' => array('left' => 100));

$phpWord->addNumberingStyle(
    'multilevel',
    array(
        'type' => 'multilevel',
        'levels' => array(
            array('format' => 'decimal', 'font'=> 'Arial', 'fontSize'=> 45,  'text' => '%1.', 'left' => 360, 'hanging' => 360, 'tabPos' => 360),
            array('format' => 'upperLetter', 'text' => '%2.', 'fontSize'=> 45, 'left' => 720, 'hanging' => 360, 'tabPos' => 720),
        )
    )
);

$subsequent = $section->addHeader();
$subsequent->addWatermark($resources.'draft.png', array('marginTop' => 500, 'marginLeft' => 55));
$spanTableStyleName = 'Colspan Rowspan';
$phpWord->addTableStyle($spanTableStyleName, $fancyTableStyle);
$table = $subsequent->addTable($fancyTableStyle);
$titleRow = array('color'=>'6337ae', 'bold'=> true);
$headerTextRun = array('color'=> '666', 'bold'=> true);
$footertitleRow = array('color'=>'000', 'bold'=> true);
$sentanceRow = array('color'=>'939393', 'bold'=> true, 'size'=>'9');
$normal = array('color'=>'939393', 'bold'=> false, 'size'=>'9');

/* SOP Header */
//ROW1
$table->addRow(700, array("exactHeight"=>false));
$cell1 = $table->addCell(3000, array('vMerge' => 'restart', 'valign' => 'center', 'bgColor' => '5225B5' ));

//$cell1 = $table->addCell(5000, $cellRowSpan1);
$textrun1 = $cell1->addTextRun($cellHCentered);
$textrun1->addImage($resources.'logo.png', 	
	array( 'width' => 156, 'height' => 97,
		'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START,
		'marginTop' => 20, 'marginLeft' => -50,
		'wrappingStyle' => 'square',
		'positioning' => 'absolute',
		'posHorizontal'    => \PhpOffice\PhpWord\Style\Image::POSITION_HORIZONTAL_LEFT,
		'posHorizontalRel' => \PhpOffice\PhpWord\Style\Image::POSITION_RELATIVE_TO_COLUMN,
		'marginLeft'       => round(\PhpOffice\PhpWord\Shared\Converter::cmToPixel(-5)),
	)); 


$cell2 = $table->addCell(10000, $cellColSpan);
$textrun2 = $cell2->addTextRun(array('alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START));
$textrun2->addText(" ".$sop_data["sop_title"], $titleRow);

// ROW2
$table->addRow(500, array("exactHeight"=>true));
$table->addCell(null, $cellRowContinue, $noSpace);
$cell3 = $table->addCell(5000, $cellVCentered);
$textrun3 = $cell3->addTextRun($headerTextRun);
$textrun3->addText(' Document No: ', $sentanceRow);
$textrun3->addText($sop_data["sop_no"], $normal);
$cell4 = $table->addCell(5000, $cellVCentered);
$textrun4 = $cell4->addTextRun($headerTextRun);
$textrun4->addText(' Revision Number: ', $sentanceRow);
$textrun4->addText($sop_data["version"], $normal);

// ROW3
$table->addRow(500, array("exactHeight"=>true));
$table->addCell(null, $cellRowContinue, $noSpace);
$cell5 = $table->addCell(5000, $cellVCentered);
$textrun5 = $cell5->addTextRun($headerTextRun);

$textrun5->addText(' Business Unit: ', $sentanceRow);
$textrun5->addText($departments, $normal);

/*foreach ($sop_data["departments"] as $value) {
	# code...
	$textrun5->addText($value, $normal);
} */

$cell6 = $table->addCell(5000, $cellVCentered);
$textrun6 = $cell6->addTextRun($headerTextRun);
$textrun6->addText(' Effective Date: ', $sentanceRow);
$textrun6->addText($sop_data["effective_from"], $normal);

/* SOP Footer*/
$footer = $section->addFooter();
$spanFooterTableStyleName = 'Colspan Rowspan';
$phpWord->addTableStyle($spanFooterTableStyleName, $fancyTableStyle1);
//$footer->addTextBreak(2);

$table = $footer->addTable($spanFooterTableStyleName);

$table->addRow(300, array("exactHeight"=>false));
$cell1 = $table->addCell(2000, $cellRowSpan);
$textrun1 = $cell1->addTextRun(array('alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START));
$textrun1->addText("", $footertitleRow);

$cell2 = $table->addCell(2500, $cellRowSpan);
$textrun1 = $cell2->addTextRun(array('alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START));
$textrun1->addText("Prepared By", $footertitleRow);

$cell3 = $table->addCell(2500, $cellRowSpan);
$textrun1 = $cell3->addTextRun(array('alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START));
$textrun1->addText("Reviewed By", $footertitleRow);

$cell4 = $table->addCell(2500, $cellRowSpan);
$textrun1 = $cell4->addTextRun(array('alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START));
$textrun1->addText("Approved By", $footertitleRow);

$table->addRow(400, array("exactHeight"=>false));
$cell1 = $table->addCell(2000, $cellRowSpan);
$textrun1 = $cell1->addTextRun(array('alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START));
$textrun1->addText("Signature & Date", $footertitleRow);

$cell2 = $table->addCell(2500, $cellRowSpan);
$textrun1 = $cell2->addTextRun(array('alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START));
$textrun1->addText("", $footertitleRow);

$cell3 = $table->addCell(2500, $cellRowSpan);
$textrun1 = $cell3->addTextRun(array('alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START));
$textrun1->addText("", $footertitleRow);

$cell4 = $table->addCell(2500, $cellRowSpan);
$textrun1 = $cell4->addTextRun(array('alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START));
$textrun1->addText("", $footertitleRow);


/* TITLE PAGE */
$textrun = $section1->addTextRun();
$section1->addTextBreak(10);
//$section->addText('TITLE PAGE HERE', array(), array('shading' => array('fill' => '6337ae')));
////$section->addTextBreak(2);

//$textrun->addText('TITLE PAGE HERE');*/
$titlePagecell1 = array('valign' => 'center', 'bgColor'=>'6337ae', 'borderColor'=>'6337ae');
$titlePagecell_txt1 = array( 'size' => '16', 'color'=> 'ffffff');
$titlePagecell2 = array('valign' => 'center', 'bgColor'=>'6337ae');
$titlePagecell_txt2 = array( 'size' => '26', 'color' => 'ffffff', 'bold' => true);
$titlePagecell3 = array('valign' => 'center', 'bgColor'=>'6337ae');
$titlePagecell_txt3 = array( 'size' => '16', 'color' => 'ffffff');
$titlePagecell_txt3_1 = array( 'size' => '16', 'color' => 'ffffff', 'bold'=>true);

$styleFirstRow = array('borderBottomSize'=>18, 'borderBottomColor'=>'6337ae', 'bgColor'=>'6337ae');
$tableStyle = array('borderColor' => 'ff0000', 'exactHeight' => false);
$table = $section1->addTable($tableStyle); 
$cell = $table->addRow(500, $styleFirstRow)->addCell(10500, $titlePagecell1 );
$cell = $table->addRow(500, $styleFirstRow)->addCell(10500, $titlePagecell1 );
$cell = $table->addRow(500,$styleFirstRow)->addCell(10500, $titlePagecell1 );
$cell->addText('Standard Operating Procedure', $titlePagecell_txt1, array('indentation' => array('left' => 200)));
$cell = $table->addRow(1200,$styleFirstRow)->addCell(10500, $titlePagecell2 );
$cell->addText($sop_data["sop_title"], $titlePagecell_txt2, array('indentation' => array('left' => 200)));
$cell = $table->addRow(900,$styleFirstRow)->addCell(10500, $titlePagecell3 );
$cell->addText('Document No:  ' . $sop_data["sop_no"] . '  '. 'Business Unit:  '  , $titlePagecell_txt3, array('indentation' => array('left' => 200)));
$cell->addText($departments, $titlePagecell_txt3);

/*foreach ($sop_data["departments"] as $value) {
	# code...
	$cell->addText($value, $titlePagecell_txt3);
} */

$cell = $table->addRow(900,$styleFirstRow)->addCell(9000, $titlePagecell3 );
$cell->addText('Revision Number:  ' . $sop_data["version"] . '  '. 'Effective Date:  '. $sop_data["effective_from"], $titlePagecell_txt3, array('indentation' => array('left' => 200)));
//$section->addPageBreak();

/* Objective */
$textrun = $section->addTextRun($HeadingparagraphStyle);
$textrun->addText(++$header_counter . '. ', $headingStyle_num);
$textrun->addText('Objective', $headingStyle);
$section->addText($sop_data["objective"], $normalTextTyle, $contentparagraphStyle);
////$section->addTextBreak(2);

/* Scope */
$textrun = $section->addTextRun($HeadingparagraphStyle);
$textrun->addText(++$header_counter . '. ', $headingStyle_num);
$textrun->addText('Scope', $headingStyle);
$section->addText($sop_data["scope"], $normalTextTyle, $contentparagraphStyle);

/* Scope Table */
$table2 = $section->addTable($fancyTableStyle);

$table2->addRow(500, array("exactHeight"=>false));
$cell7 = $table2->addCell( 3000, $cellVCentered);
$cell7->addText('Department', array('bold'=> true, 'name'=>'Arial', 'size'=>'12'), $indentStyle);

$cell8 = $table2->addCell(7000, $cellVCentered);
$cell8->addText($departments, $normalTextStyle_cell, $indentStyle);
/*foreach ($sop_data["departments"] as $value) {
	# code...
	$cell8 = $table2->addCell(7000, $cellVCentered);
	$cell8->addText($value, $normalTextTyle, $indentStyle);
}*/
if(isset($sop_data["area"])){
	$table2->addRow(500, array("exactHeight"=>false));
	$cell7 = $table2->addCell( 3000, $cellVCentered);
	$cell7->addText('Area', array('bold'=> true, 'name'=>'Arial', 'size'=>'12'), $indentStyle);
	$cell8 = $table2->addCell(7000, $cellVCentered);
	$cell8->addText($sop_data["area"], $normalTextStyle_cell, $indentStyle);
}

if(isset($sop_data["activity"])){
	$table2->addRow(500, array("exactHeight"=>false));
	$cell7 = $table2->addCell( 3000, $cellVCentered);
	$cell7->addText('Activity', array('bold'=> true, 'name'=>'Arial', 'size'=>'12'), $indentStyle);
	$cell8 = $table2->addCell(7000, $cellVCentered);
	$cell8->addText($sop_data["activity"], $normalTextStyle_cell, $indentStyle);
}	

if(isset($sop_data["remarks"])){
	$table2->addRow(500, array("exactHeight"=>false));
	$cell7 = $table2->addCell( 3000, $cellVCentered);
	$cell7->addText('Remarks', array('bold'=> true, 'name'=>'Arial', 'size'=>'12'),  $indentStyle);
	$cell8 = $table2->addCell(7000, $cellVCentered);
	$cell8->addText($sop_data["remarks"], $normalTextStyle_cell, $indentStyle);
}

/* Responsibilities  */
//$section->addTextBreak(2);
$textrun = $section->addTextRun($HeadingparagraphStyle);
$textrun->addText(++$header_counter . '. ', $headingStyle_num);
$textrun->addText('Responsibilities', $headingStyle);

//$section->addTextBreak(1);
/* Responsibilities THead */
$table2 = $section->addTable($fancyTableStyle);
$table2->addRow(500, array("exactHeight"=>false));
$cell7 = $table2->addCell( 3000, $cellVCentered);
$cell7->addText('Department', array('bold'=> true, 'name'=>'Arial', 'size'=>'12'), $indentStyle );
$cell8 = $table2->addCell(3000, $cellVCentered);
$cell8->addText('Role', array('bold'=> true, 'name'=>'Arial', 'size'=>'12'), $indentStyle);
$cell9 = $table2->addCell(4000, $cellVCentered);
$cell9->addText('Responsibility', array('bold'=> true, 'name'=>'Arial', 'size'=>'12'), $indentStyle);

$rolesStyle = array('color'=>'666', 'name'=>'Arial', 'size'=>'12');

/* Responsibilities TBody */
foreach ($roles as $key => $value) {
	$table2->addRow(500, array("exactHeight"=>false));
	$cell7 = $table2->addCell( 3000, $cellVCentered);
	$cell7->addText($value["department"], $rolesStyle, $indentStyle);
	$cell8 = $table2->addCell(3000, $cellVCentered);
	$cell8->addText($value["role"], $rolesStyle, $indentStyle);
	$cell9 = $table2->addCell(4000, $cellVCentered);
	//if(count($value["responsibility"])>0){
	if(isset($value['responsibility'])) {
		$cell9->addText($value["responsibility"], $rolesStyle, 
		$indentStyle);
	}
	
}

/* General Safety And Guildelines */
if(count($sop_data["safety_warnings"])>0){
	$textrun = $section->addTextRun($HeadingparagraphStyle);
	$textrun->addText(++$header_counter . '. ', $headingStyle_num);
	$textrun->addText('General Safety Precautions and Guidelines', $headingStyle);

	//$section->addTextBreak(1);
	$table2 = $section->addTable($gspgTable);
	$table2->addRow(1000);
	$cell7 = $table2->addCell( 1000, array('vMerge' => 'restart', 'valign'=>'center', 'bgColor'=>'F99819'));
	$cell7->addImage($resources.'gspg.png', array('width' => 34, 'height' => 34, 'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER));

	foreach ($sop_data["safety_warnings"] as $key => $value) {
		# code...
		if($key == 0){
			//$cell8 = $table2->addCell(8000, $cellVCentered);
			$cell8 = $table2->addCell(8000, array('valign' => 'center', 'bgColor'=> 'FAF9FA', 'borderColor'=>'FAF9FA', 'borderSize'=>0));
			$cell8->addListItem($value, 0, array('name'=>'Arial', 'size'=>'12', 'color'=>'000000', 'bgColor' => 'FAF9FA','bold'=> true), $gspgList);	
		} else {
			$table2->addRow();
			$table2->addCell(null, $cellRowContinue, $noSpace);
			$cell8 = $table2->addCell(8000, array('valign' => 'center', 'bgColor'=> 'FAF9FA', 'borderColor'=>'FAF9FA', 'borderSize'=>0));
			$cell8->addListItem($value, 0, array('name'=>'Arial', 'size'=>'12', 'color'=>'000000', 'bgColor' => 'FAF9FA','bold'=> true), $gspgList);
		}
	}
}

/*  Procedure  */
//$section->addText(++$header_counter . '. Procedures', $headingStyle, $HeadingparagraphStyle);
$textrun = $section->addTextRun($HeadingparagraphStyle);
$textrun->addText(++$header_counter . '. ', $headingStyle_num);
$textrun->addText('Procedures', $headingStyle);

foreach ($procedure as $key => $value) {
	if($key !=0){
		//$section->addTextBreak(2);
		$section->addText("", $subHeadingStyle, $subHeadingparagraphStyle);
	}
	$count = $header_counter . '.' . ($key+1);
	$proc_title = $section->addTable($fancyTableStyle);
	$proc_title->addRow(200, array("exactHeight"=>false));
	
	$cell1 = $proc_title->addCell(1500, array('valign' => 'center', 'bgColor'=> 'FAF9FA', 'borderColor'=>'FAF9FA', 'borderSize'=>0));
	$cell1->addText( $count, array('name'=>'Arial', 'size'=>'14', 'color'=>'5225B5', 'bold'=> true), array('indentation' => array('left' => 100), 'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER));
	
	$cell2 = $proc_title->addCell(8500,array('valign' => 'center', 'bgColor'=> '5225B5', 'borderColor'=>'5225B5', 'borderSize'=>0));
	$cell2 ->addText($value["name"], 
			array('name'=>'Arial', 'size'=>'14', 'color'=>'FFFFFF', 'bold'=> true), 
			array('indentation' => array('left' => 100)));
	  
	$section->addText($value["description"], $normalTextTyle, $contentparagraphStyle);
	$dir = $value["flow_imgs_dir_path"];
	$files1 = array_slice(scandir($dir), 2);// rm . and .. dir
	sort($files1);
	echo is_array($files1) ? 'Array' : 'not an Array';
	echo "\n";
	foreach($files1 as $image)
	{
	  if (strpos($image, 'tile') !== false) {
		//echo  $image;
		$section->addImage($value["flow_imgs_dir_path"]. "/" . $image);
	  }	
	}

	//$section->addTextBreak(1);
	foreach ($value["steps"] as $key1 => $value1) {
		if($key1 !=0){
			//$section->addTextBreak(2);
			$section->addText("", $subHeadingStyle, $subHeadingparagraphStyle);
		}
		
		$header_set = false; 
		
		$step_count = $count . '.' . $value1["seq_id"]  . ' '; 
		$step_name =  $value1["name"]; 
		
		$table = $section->addTable($spanTableStyleName);
		$table->addRow(400, array("cantSplit" => true, "exactHeight"=> false));

		$cell2 = $table->addCell(700,  array('gridSpan' => 2, 'valign' => 'center', 'borderColor'=>'FFFFFF', 'borderSize'=>0));
		$textrun1 = $cell2->addTextRun($cellHCentered);
		$textrun1->addImage($resources.'step_title.png', array('width' => 30, 'height' => 30,
		'marginBottom'     => 100,	
		 'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER ));
		$textrun1->addTextBreak(1.5, array('name'=>'Arial', 'size'=>'14'));
		$textrun1->addText($step_count, 
		array('name'=>'Arial', 'size'=>'9', 'color'=>'FFFFFF', 'bold'=> true, 'bgColor'=>'5225B5')
		,array('spaceBefore'=> 500)
		);

		$cell1 = $table->addCell(9300, array('vMerge' => 'restart', 'valign' => 'center', 'borderColor'=>'FFFFFF', 'borderSize'=>0));
		$textrun1 = $cell1->addTextRun();
		$textrun1->addText($step_name, array('name'=>'Arial', 'size'=>'12', 'color'=>'5225B5', 'bold'=> true));

		//Information
		if(isset($value1["information"][0]["value"]) && (strlen($value1["information"][0]["value"]) > 0)){
			$section->addText('Information', $subTitleStyle);
			//$section->addTextBreak(1);
			$header_set = true;
			$info_data  = $value1["information"][0]["value"];
			$info_data_1 = str_replace("<br>","", $info_data);
			//echo "&*CONTENT =====>>>>". $info_data_1;
			\PhpOffice\PhpWord\Shared\Html::addHtml($section, $info_data_1);
		}
	
		// Preacaution and warning
		if(count($value1["precaution_warnings"])>0){

			if(!$header_set){
				$section->addText('Information', $subTitleStyle);
				//$section->addTextBreak(1);
				$header_set = true;
			}

			$table2 = $section->addTable($gspgTable);
			$table2->addRow();
			$cell7 = $table2->addCell( 1000, array('vMerge' => 'restart', 'valign'=>'center', 'bgColor'=>'F99819', 'borderColor'=>'FFFFFF', 'borderSize'=>0));
			$textrun1 = $cell7->addTextRun($cellHCentered);
			$textrun1->addImage($resources.'warn.png', array('width' => 34, 'height' => 34, 'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START));
			foreach ($value1["precaution_warnings"] as $key2 => $value2) {
				if($key2 == 0){
					$cell8 = $table2->addCell(8000, array('valign' => 'center', 'color'=> '666', 'borderColor'=>'FFFFFF', 'borderSize'=>0));
					$cell8->addListItem($value2, 0, $gspg, $gspgList);	
				} else {
					$table2->addRow();
					$table2->addCell(null, array('vMerge' => 'continue', 'borderColor'=>'FFFFFF', 'borderSize'=>0), $noSpace);
					$cell8 = $table2->addCell(8000, array('valign' => 'center', 'color'=> '666', 'borderColor'=>'FFFFFF', 'borderSize'=>0));
					$cell8->addListItem($value2, 0, $gspg, $gspgList);
				}

			}
		}

 		// acceptance_criteria 
		if(count($value1["acceptance_criteria"])>0){
			if(!$header_set){
				$section->addText('Information', $subTitleStyle);
				//$section->addTextBreak(1);
				$header_set = true;
			}

	  		$section->addTextBreak(1);
			$table2 = $section->addTable($gspgTable);
			$table2->addRow();
			$cell7 = $table2->addCell( 1000, array('vMerge' => 'restart', 'valign'=>'center', 'bgColor'=>'00A26D', 'borderColor'=>'FFFFFF', 'borderSize'=>0));
			$textrun1 = $cell7->addTextRun($cellHCentered);
			//$textrun1->addText("test");
			$textrun1->addImage($resources.'check.png', array('width' => 34, 'height' => 34, 'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START));
			//$section->addText($text, [$fontStyle]);
			foreach ($value1["acceptance_criteria"] as $key3 => $value3) {
				# code...
				if($key3 == 0){
					$cell8 = $table2->addCell(8000,array('valign' => 'center', 'color'=> '666', 'borderColor'=>'FFFFFF', 'borderSize'=>0));
					$cell8->addListItem($value3, 0, $gspg, $gspgList);	
				} else {
					$table2->addRow();
					$table2->addCell(null, array('valign' => 'center', 'color'=> '666', 'borderColor'=>'FFFFFF', 'borderSize'=>0), $noSpace);
					$cell8 = $table2->addCell(8000, $cellVCentered);
					$cell8->addListItem($value3, 0, $gspg, $gspgList);
				}
			}	
		}	
	}
}

/* Definitions */
if(count($sop_data["definitions"])>0){
	// $section->addText(++$header_counter . '. Definitions', $headingStyle, $HeadingparagraphStyle);

	$textrun = $section->addTextRun($HeadingparagraphStyle);
	$textrun->addText(++$header_counter . '. ', $headingStyle_num);
	$textrun->addText('Definitions', $headingStyle);

	foreach ($sop_data["definitions"] as $key => $value) {
		$table2 = $section->addTable($fancyTableStyle);
		$table2->addRow(500, array("exactHeight"=>false));
		$cell7 = $table2->addCell( 3000, $cellVCentered);
		$cell7->addText($value["key"], array('bold'=> true, 'name'=>'Arial', 'size'=>'12'), $indentStyle);
		$cell8 = $table2->addCell(7000, $cellVCentered);
		$cell8->addText($value["value"], $normalTextStyle_cell, $indentStyle);
	}
} else {
	echo "Definitions Array is NULL";
}

/* Abbreviations */
if(count($sop_data["abbreviations"])>0){
	//$section->addTextBreak(2);
	//$section->addText(++$header_counter . '. Abbreviations', $headingStyle, $HeadingparagraphStyle);

	$textrun = $section->addTextRun($HeadingparagraphStyle);
	$textrun->addText(++$header_counter . '. ', $headingStyle_num);
	$textrun->addText('Abbreviations', $headingStyle);

	foreach ($sop_data["abbreviations"] as $key => $value) {
		# code...

		$table2 = $section->addTable($fancyTableStyle);
		$table2->addRow(500, array("exactHeight"=>false));
		$cell7 = $table2->addCell( 3000, $cellVCentered);
		$cell7->addText($value["key"], array('bold'=> true, 'name'=>'Arial', 'size'=>'12'), $indentStyle);
		$cell8 = $table2->addCell(7000, $cellVCentered);
		$cell8->addText($value["value"],  $normalTextStyle_cell, $indentStyle);
	}
} else {
	echo "Abbreviations Array is NULL";
}

/* References */
if(count($references)>0){
	//$section->addTextBreak(2);
	//$section->addTextBreak(1);
	
	$textrun = $section->addTextRun($HeadingparagraphStyle);
	$textrun->addText(++$header_counter . '. ', $headingStyle_num);
	$textrun->addText('References', $headingStyle);

	foreach ($references as $key => $value) {
		# code...
		$table2 = $section->addTable($fancyTableStyle);
		$table2->addRow(500, array("exactHeight"=>false));
		$cell7 = $table2->addCell( 3000, $cellVCentered);
		$cell7->addText($value["doc_title"], array('bold'=> true, 'name'=>'Arial', 'size'=>'12'), $indentStyle);
		$cell8 = $table2->addCell(7000, $cellVCentered);
		$cell8->addLink($value["doc_link"], $value["doc_no"], $normalTextStyle_cell, $indentStyle);
	}
} else {
	echo "References are null";
}

/* Annexures */
if(count($annexures)>0){
	//$section->addTextBreak(2);
	//$section->addText(++$header_counter .'. Annexures', $headingStyle, $HeadingparagraphStyle);

		
	$textrun = $section->addTextRun($HeadingparagraphStyle);
	$textrun->addText(++$header_counter . '. ', $headingStyle_num);
	$textrun->addText('Annexures', $headingStyle);

	foreach ($annexures as $key => $value) {

		$table2 = $section->addTable($fancyTableStyle);
		$table2->addRow(500, array("exactHeight"=>false));
		$cell7 = $table2->addCell( 3000, $cellVCentered);
		$cell7->addText($value["doc_title"], array('bold'=> true, 'name'=>'Arial', 'size'=>'12'), $indentStyle);
		$cell8 = $table2->addCell(7000, $cellVCentered);
		$cell8->addLink($value["doc_link"], $value["doc_no"], $normalTextStyle_cell, $indentStyle);
	}
} else {
	echo "Annexures are null";
}

$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
try {
    /*$objWriter->save('./results/drl_sop.docx');*/
$objWriter->save($output_file);
} catch (Exception $e) {
return ['ERROR'=>'could not save','e'=>$e];
}
//echo write($phpWord, basename(__FILE__, '.php'), $writers);
//if (!CLI) {
//    include_once 'Sample_Footer.php';
//}

/*
switch ($extension) {
	case 'pdf':
		$writer = PhpWord\IOFactory::createWriter($phpWord, 'PDF');
		break;
	case 'docx':
		$writer = PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
		break;
	case 'odt':
		$writer = PhpWord\IOFactory::createWriter($phpWord, 'ODText');
		break;
}*/