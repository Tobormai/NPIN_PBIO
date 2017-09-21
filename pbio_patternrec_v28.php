
<?php
// the following line prevents the browser from parsing this as HTML.
header('Content-Type: text/plain');
//('Content-type: text/html; charset= iso-8859-1');  
//header('Content-type: text/html; charset= UTF-8');  
require_once 'PHPExcel.php';
ini_set('max_execution_time', 800); //300 seconds = 5 minutes

class ExcelRecord
{
	//PROPERTIES
	private static $row = 1;
	private $col;
	
	public $activeUrl;
	
	public $arrayUrlMatches = array();	
	public $arrayWriteToExcel = array();
	public $arrayStringMatches = array();
	public $arrayActiveUrl = array();
	public $arrayUrlContent = array();
	public $arrayColumnHeaders = array();
	
	public $element;
	public $searchUrl;
	public $matches = array();
	public $recordUrl;
	public $regexPattern;
	public $searchCriteria;	
	

	//CONSTRUCTORS
	
	public function __construct()
	  {
		  echo 'The class "', __CLASS__, '" was initiated!'. "<br>";
	  }
	
	//GETTERS
	public function getSearchCritera()
	{
		return $this->searchCriteria;
	}	
	
	public function getUrlMatches()
	{
		// file_get_contents necessary to convert page contents to string
		$subject = file_get_contents($this->searchUrl);

		// escape special characters in the query
		$regex = preg_quote($this->searchCriteria, '/');

		// finalise the regular expression, matching the whole line
		$regex = "/^.*$regex.*\$/m";

		// search, and store all matching occurences in $arrayStringMatches
		if(preg_match_all($regex, $subject, $this->arrayUrlMatches)){
		   echo "Found Url Matches:\n";
		   //echo implode("\n", $this->arrayUrlMatches[0]);
		}
		else{
			echo "No matches found";
		}		
		return $this->arrayUrlMatches;
	}	
	
	/* FUNCTION DEFINITION: 
	* subtok(string,chr,pos,len)
	* chr = chr used to seperate tokens
	* pos = starting postion
	* len = length, if negative count back from right
	*/
	public function subtok($string,$chr,$pos,$len=NULL) 
	{
		return implode($chr,array_slice(explode($chr,$string),$pos,$len));
	}
	
	/*public function getIndexNum_new()
	{
		$a = $this->arrayUrlMatches[0];
		foreach ($a as $value)
		{
			echo $this->getStringBetweenCriteria($this->subtok($value,'/',6,1))."<br>";
		}
		//arrayUrlMatches unchanged after this method completes
	}*/
	
		
	/* FUNCTION DEFINITION: 
	* getStringBetweenCriteria($ui_firstChar,$ui_secondChar)
	* input = string within to search
	* output = output[0] text that matched the full pattern, output[1] text that matched first subpattern
	*/
	public function getStringBetweenCriteria($ui_FirstBound,$ui_SecondBound)
	{		
		echo "FUNCTION: getStringBetweenCriteria";
		$a = $this->arrayUrlMatches[0];
		foreach ($a as &$value)
		{
			$input = $value;			
			$arrayElement = preg_match('~'.$ui_FirstBound.'(.*?)'.$ui_SecondBound.'~', $input, $output);
			//echo $output[0]; //outputs full string 
			//echo $output[1]; //outputs string between criteria
			$this->arrayWriteToExcel[] = $output;
		}
		//print_r($this->arrayWriteToExcel);
		return $this->arrayWriteToExcel;
	}
	
	public function getIndexNum()
	{		
		$this->matches = $this->arrayUrlMatches;
		$a = $this->arrayUrlMatches[0];
		foreach ($a as $value)
		{
			 $this->setIndexNumber($value);
		}
		//arrayUrlMatches unchanged after this method completes
		//print_r($this->arrayUrlMatches);
		//print_r($this->matches);
	}
		
	public function getBetweenTags($regexPattern,$arrayActiveUrl)
	{
		print_r($this->matches);		
	}
	
	public function writeToExcel()
	{
		
	}
	
	//SETTERS
	public function setSearchCriteria($ui_SearchCriteria)
	{
		$this->searchCriteria = $ui_SearchCriteria; 
	}
	
	public function setSearchUrl($ui_SearchUrl)
	{
		$this->searchUrl = $ui_SearchUrl; 
	}
		
	public function setIndexNumber($value)
	{
		$startIndex = strpos($value, ".html") - 6;
		//echo "$startIndex <br>";
		$indexNum = substr($value, $startIndex, 6);	
		echo "Index Number: "."$indexNum" ."<br>";	
		$this->element = $indexNum;
	}	
	
	public function setRecordUrl()
	{
		$this->recordUrl = 'http://www.plantsystematics.org/reveal/pbio/digitalimages/'.$this->element.'.html';
		return $this->recordUrl;
	}
	
	public function setRegex($pattern)
	{
		$this->regexPattern = $pattern; 		
	}
	
	public function setArrayExcelColumnHeaders($ui_ArrayColumnHeaders)
	{
		$this->arrayColumnHeaders = $ui_ArrayColumnHeaders;
		echo "Printing array of column headers: ";
		print_r($this->arrayColumnHeaders);
	}
	
	public function setArrayRecordUrl($found)
	{
		$this->matches = $found;
		$a = $found[0];
		foreach ($a as $value)
		{			 
			 // Get record number from string and store as element used to build record URL
			 $this->setIndexNumber($value);	
			 
			// Set each Webpage URL	using record number		
			$this->setRecordUrl();
			echo "Record URL: ".$this->recordUrl."<br>";
			$value = $this->recordUrl;
			echo "URLBASE: ".$value."<br>";
			
			// Store each Record Url into array
			$found[]= $value;
			$this->arrayActiveUrl = $found;			
		}
		print_r($found);		
		return $this->arrayActiveUrl;			
	}
	
	public function writeExcelRecord($found)
	{
		// Create PHPExcel object which represents Excel Workbook
		$excel = new PHPExcel();
	
		static $col = 'A';
		static $row = '2';
			
		$this->matches = $found;
		$a = $found[0];
		echo "Found Url Regex content to write to Excel \n";
		   echo implode("\n", $found[0]);
		
		foreach ($a as $value)
		{
			 //echo "Record Number (indexNum LOOP): ";
			 $this->setIndexNumber($value);
						
			// Write contents from Webpage to empty cell
			$excel->setActiveSheetIndex(0)
			->setCellValueByColumnAndRow($col,$row, $this->element);

		// Increment Cell Value by 1 row
		$row++;
			
		}
		// Save Excel 2007 file
		$file = PHPExcel_IOFactory::createWriter($excel, 'Excel2007');
		$file->save('PBIO Raw Data_RecordNums.xlsx');
	}
	
	public function writeExcelSource($ui_ColInit,$ui_RowInit,$ui_SheetName,$ui_FileName)
	{
		echo "Printing Excel Write Url Content Regex matches:\n";
		print_r ($this->arrayWriteToExcel);
		$headerArray = $this->arrayColumnHeaders;
		$rowArray = $this->arrayWriteToExcel;
		
		$columnArray = array_chunk(array_column($rowArray,'1'),1); //used for 1D array as column values vs row
				
		
		// Create PHPExcel object which represents Excel Workbook
		$objPHPExcel = new PHPExcel();
		$objWriter = new PHPExcel_Writer_CSV($objPHPExcel);
		
		
		$fileType = 'Excel2007';
		$fileName = $ui_FileName;

		// Read the file
		$objReader = PHPExcel_IOFactory::createReader($fileType);
		$objPHPExcel = $objReader->load($fileName);

		// Fill worksheet with header values from array
		$objPHPExcel->getActiveSheet()->fromArray($headerArray, null, 'A1');
		
		// Fill worksheet from values in array
		//$objPHPExcel->getActiveSheet()->fromArray($rowArray, null, '$ui_ColInit.$ui_RowInit2');	//2D array		
		$objPHPExcel->getActiveSheet()->fromArray($columnArray,NULL,$ui_ColInit.$ui_RowInit);	//1D array	

		// Name worksheet
		$objPHPExcel->getActiveSheet()->setTitle($ui_SheetName);	

		// Save Excel 2007 file
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
        $objWriter->save($ui_FileName.'.xlsx');
		//$file_put_contents($file,$urlCurrent);		
	}
}
//Instantiate the WriteExcelRecords class.
$records = new ExcelRecord();

//Set webpage on which to search based on criteria - WHERE TO SEARCH
$records->setSearchUrl("http://www.plantsystematics.org/reveal/pbio/digitalimages/digslideindexI.html");
print 'Search Url Address: ' .$records->searchUrl. "<br>";

//Set search criteria - WHAT TO SEARCH FOR
$records->setSearchCriteria("Iberis");
print 'Search Criteria: ' .$records->getSearchCritera(). "<br>";

//Store search results in arrayStringMatches - RESULTS OF SEARCH
print_r($records->getUrlMatches());

//Extract record number from URL string and set record number as element for each URL
$records->getIndexNum();
//print_r($records->getIndexNum_new());


//WRITE CONTENTS TO EXCEL

//Set Array of Excel Column Headers
$ui_ArrayColumnHeaders = array('scientific_name','common_name','accession_date','image_date','slideindex','city','county','state','locationnotes');
$records->setArrayExcelColumnHeaders($ui_ArrayColumnHeaders);

//Search for substring based on boundaries - GET SUBSTRING (scientific_name)
$ui_FirstBound = '<i>';
$ui_SecondBound = '</i>';
print_r($records->getStringBetweenCriteria($ui_FirstBound,$ui_SecondBound));


//Search for substring based on boundaries - GET SUBSTRING (common_name)
$ui_FirstBound = 'http://www.plantsystematics.org/reveal/pbio/digitalimages/';
$ui_SecondBound = '.html';
print_r($records->getStringBetweenCriteria($ui_FirstBound,$ui_SecondBound));

/*
//Search for substring based on boundaries - GET SUBSTRING (accession_date)
$ui_FirstBound = 'http://www.plantsystematics.org/reveal/pbio/digitalimages/';
$ui_SecondBound = '.html';
print_r($records->getStringBetweenCriteria($ui_FirstBound,$ui_SecondBound));

//Search for substring based on boundaries - GET SUBSTRING (slideindex)
$ui_FirstBound = 'http://www.plantsystematics.org/reveal/pbio/digitalimages/';
$ui_SecondBound = '.html';
print_r($records->getStringBetweenCriteria($ui_FirstBound,$ui_SecondBound));

//Search for substring based on boundaries - GET SUBSTRING (city)
$ui_FirstBound = 'http://www.plantsystematics.org/reveal/pbio/digitalimages/';
$ui_SecondBound = '.html';
print_r($records->getStringBetweenCriteria($ui_FirstBound,$ui_SecondBound));

//Search for substring based on boundaries - GET SUBSTRING (county)
$ui_FirstBound = 'http://www.plantsystematics.org/reveal/pbio/digitalimages/';
$ui_SecondBound = '.html';
print_r($records->getStringBetweenCriteria($ui_FirstBound,$ui_SecondBound));

//Search for substring based on boundaries - GET SUBSTRING (state)
$ui_FirstBound = 'http://www.plantsystematics.org/reveal/pbio/digitalimages/';
$ui_SecondBound = '.html';
print_r($records->getStringBetweenCriteria($ui_FirstBound,$ui_SecondBound));

//Search for substring based on boundaries - GET SUBSTRING (locationnotes)
$ui_FirstBound = 'http://www.plantsystematics.org/reveal/pbio/digitalimages/';
$ui_SecondBound = '.html';
print_r($records->getStringBetweenCriteria($ui_FirstBound,$ui_SecondBound));

*/

//Open File, Set Column and Row position, Sheet name, and File name
$ui_ColInit =  "A";
$ui_RowInit = "6";
$ui_SheetName = "Ia_Iz";
$ui_FileName = "PBIO Digital Images Index_Ia-Iz";
$records->writeExcelSource($ui_ColInit,$ui_RowInit,$ui_SheetName,$ui_FileName);



// Set criteria to return contents only between specified HTML tags
//$records->setRegex("blockquote");

//Writes each record number for matching search criteria to Excel file
//$records->writeExcelRecord($records->arrayStringMatches);

//
//print_r($records->getBetweenTags($records->searchCriteria,$records->matches));

//Writes source code for each url based on record number inserted into URL
//$records->writeExcelSource($records->arrayWriteToExcel);


?>

