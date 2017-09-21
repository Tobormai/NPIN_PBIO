
<?php
// the following line prevents the browser from parsing this as HTML.
header('Content-Type: text/plain');
//('Content-type: text/html; charset= iso-8859-1');  
//header('Content-type: text/html; charset= UTF-8');  
require_once 'PHPExcel.php';
require_once 'PHPExcel/IOFactory.php';
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
	public $arrayUrlSource = array();
	public $arrayColumnHeaders = array();
	
	public $arraySetLocationCriteria = array();
	
	public $element;
	public $searchUrl;
	public $matches = array();
	public $recordUrl;
	public $regexPattern;
	public $searchCriteria;	
	public $regexHtmlTag;


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
		echo "Function subtok";
		echo "String passed to function: " .$string;
		return implode($chr,array_slice(explode($chr,$string),$pos,$len));
	}	
		
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
			preg_match('~'.$ui_FirstBound.'(.*?)'.$ui_SecondBound.'~', $input, $output);
			//echo $output[0]; //outputs full string 
			//echo $output[1]; //outputs string between criteria
			$this->arrayWriteToExcel[] = $output;
		}
		//print_r($this->arrayWriteToExcel);
		return $this->arrayWriteToExcel;
	}
	
	public function getLocationDetail($ui_FirstBound,$ui_SecondBound,$ui_LocationSpecifier)
	{
		echo "FUNCTION: getLocationDetail";
		$a = $this->arrayUrlMatches[0];
		foreach ($a as &$value)
		{
			$input = $value;
			echo $input;			
			preg_match('~'.$ui_FirstBound.'(.*?)'.$ui_SecondBound.'~', $input, $output);
						
			//echo $output[0]; //outputs full string 
			//echo $output[1]; //outputs string between criteria
			$this->arrayWriteToExcel[] = $output[0];
		
		}
		if(isset($ui_LocationSpecifier))
		{		
			$this->setLocationCriteria($ui_LocationSpecifier);
			$this->arrayWriteToExcel = $this->arraySetLocationCriteria;
			return $this->arrayWriteToExcel;
		}
		else
		{
			$this->arrayWriteToExcel[] = $output;			
		}			
	}
	
	public function getUrlContent($newTag)
	{
		echo "FUNCTION: getUrlContent";
		$a = $this->arrayWriteToExcel;
		print_r($a);
		
		foreach ($a as &$value)
		{
			$urlBase = $value[0];
			echo "urlBase: ".$urlBase;
			$urlCurrent = file_get_contents($urlBase);
			$input = $urlCurrent;
			$regex = "~<blockquote\b[^>]*>(?:[^<]+|(?R)|<(?!/(?:blockquote|p)>))*</blockquote>~";
			$regex = "~<$newTag\b[^>]*>(?:[^<]+|(?R)|<(?!/(?:$newTag|p)>))*</$newTag>~";
			preg_match($regex, $input, $output);
			$this->arrayUrlSource[] = $output;
		}	
		return $this->arrayUrlSource;
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
		//print_r($this->arrayColumnHeaders);
	}
	
	public function setLocationCriteria($ui_LocationSpecifier)
	{
		//Set City, County, State Condition
		$locationSpecifier = $ui_LocationSpecifier;	
		echo "FUNCTION: setLocationCriteria. Printing array: ";
		
		$a = $this->arrayWriteToExcel;
		//print_r($a);
		
		foreach ($a as $value)
		{
			$input = $value;
			if(substr_count($value,',')== 2)
			{
				$output = $this->subtok($input,',',2,3);
			}
			//$output = $this->subtok($input,',',1,1);
			echo "Printing output".$output."<br>";
			$this->arraySetLocationCriteria[] = $output;
		}	
		//print_r($this->arraySetLocationCriteria);
		return $this->arraySetLocationCriteria;		
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
	
	public function writeExcelSource($matches, $ui_ColInit,$ui_RowInit,$ui_SheetName,$ui_FileName,$ui_FileExtension)
	{
		//echo "Printing Excel Write Url Content Regex matches:\n";
		//print_r ($this->arrayWriteToExcel);

		$headerArray = $this->arrayColumnHeaders;
		//$rowArray = $this->arrayWriteToExcel;	
		$rowArray = $matches;	
		$columnArray1d = array_chunk($rowArray,1);
		$columnArray2d = array_chunk(array_column($rowArray,'0'),1); //modify array_column column_key to output specified column for 2d array
				
		
		$filename = $ui_FileName.$ui_FileExtension;
		if (file_exists($filename)) 
		{			
			// Load existing Excel file
			$objPHPExcel = PHPExcel_IOFactory::load($ui_FileName.$ui_FileExtension);
			echo $filename ." found";
		} 
		else 
		{
			echo  $filename. " not found. Creating new Excel Workbook";
			// Create PHPExcel object which represents Excel Workbook
			$objPHPExcel = new PHPExcel();
			$objWriter = new PHPExcel_Writer_CSV($objPHPExcel);		
		}

		// Fill worksheet with header values from array
		$objPHPExcel->getActiveSheet()->fromArray($headerArray, null, 'A1');
		
		// Fill worksheet from values in array
		//$objPHPExcel->getActiveSheet()->fromArray($rowArray, null, $ui_ColInit.$ui_RowInit);	//2D array		
		$objPHPExcel->getActiveSheet()->fromArray($columnArray2d,NULL,$ui_ColInit.$ui_RowInit);	//1D array	

		// Rename worksheet
		$objPHPExcel->getActiveSheet()->setTitle($ui_SheetName);	
	
		$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
		$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');	
		$objWriter->save($ui_FileName.$ui_FileExtension);		
	}
	
	public function saveExcel($ui_FileName)
	{
				// Save Excel 2007 file
        
        
		//$file_put_contents($file,$urlCurrent);
	}	
}
//Instantiate the WriteExcelRecords class.
$records = new ExcelRecord();

//Set webpage on which to search based on criteria - WHERE TO SEARCH
$records->setSearchUrl("http://www.plantsystematics.org/reveal/pbio/digitalimages/digslideindexM.html");
print 'Search Url Address: ' .$records->searchUrl. "<br>";

//Set search criteria - WHAT TO SEARCH FOR
$records->setSearchCriteria("Malva");
print 'Search Criteria: ' .$records->getSearchCritera(). "<br>";

//Store search results in arrayStringMatches
print_r($records->getUrlMatches());

//Extract record number from URL string and set record number as element for each URL
$records->getIndexNum();
//print_r($records->getIndexNum_new());


//WRITE CONTENTS TO EXCEL - DEFINE CRITERIA

//Set Array of Excel Column Headers
$ui_ArrayColumnHeaders = array('scientific_name','common_name','accession_date','image_date','slideindex','city','county','state','locationnotes');
$records->setArrayExcelColumnHeaders($ui_ArrayColumnHeaders);

/* 
//Search for substring based on boundaries - GET SUBSTRING (A:scientific_name)
$ui_ColInit =  "A";
$ui_FirstBound = '<i>';
$ui_SecondBound = '</i>';
print_r($records->getStringBetweenCriteria($ui_FirstBound,$ui_SecondBound));
*/


/* 
//Search for substring based on boundaries - GET SUBSTRING (B:common_name)
$ui_ColInit =  "B";
$ui_FirstBound = 'http://www.plantsystematics.org/reveal/pbio/digitalimages/';
$ui_SecondBound = '.html';
print_r($records->getStringBetweenCriteria($ui_FirstBound,$ui_SecondBound));
 */


//Search for substring based on boundaries - GET SUBSTRING (C:accession_date)
$ui_ColInit =  "C";
$newTag = "h5";

$ui_FirstBound = 'http://www.plantsystematics.org/reveal/pbio/digitalimages/';
$ui_SecondBound = '.html';
print_r($records->getStringBetweenCriteria($ui_FirstBound,$ui_SecondBound));
$matches = $records->getUrlContent($newTag);
echo "Printing accession_date array urls";
print_r($records->arrayUrlSource);
$ui_RowInit = "2";
$ui_SheetName = "Ma_Mz";
$ui_FileName = "PBIO Digital Images Index_Ma-Mz";
$ui_FileExtension = ".xlsx";
$records->writeExcelSource($matches,$ui_ColInit,$ui_RowInit,$ui_SheetName,$ui_FileName,$ui_FileExtension);

//$ui_FirstBound = '<h5>';
//$ui_SecondBound = '</h5';
//print_r($records->getStringBetweenCriteria($ui_FirstBound,$ui_SecondBound));

 
/* 
//Search for substring based on boundaries - GET SUBSTRING (D:image_date)
$ui_ColInit =  "D";
$ui_FirstBound = '<h5>';
$ui_SecondBound = ';';
print_r($records->getStringBetweenCriteria($ui_FirstBound,$ui_SecondBound));
*/ 
 
/* 
//Search for substring based on boundaries - GET SUBSTRING (E:slideindex)
$ui_ColInit =  "E";
$ui_FirstBound = 'http://www.plantsystematics.org/reveal/pbio/digitalimages/';
$ui_SecondBound = '.html';
print_r($records->getStringBetweenCriteria($ui_FirstBound,$ui_SecondBound));
*/
 
 
/* 
//Search for substring based on boundaries - GET SUBSTRING (F:city)
$ui_ColInit =  "F";
$ui_FirstBound = '</a>';
$ui_SecondBound = '<br>';
$ui_LocationSpecifier = "city";
print_r($records->getLocationDetail($ui_FirstBound,$ui_SecondBound,$ui_LocationSpecifier));
 */
 
/*   
//Search for substring based on boundaries - GET SUBSTRING (G:county)
$ui_ColInit =  "G";
$ui_FirstBound = '</a>';
$ui_SecondBound = '<br>';
$ui_LocationSpecifier = "county";
print_r($records->getLocationDetail($ui_FirstBound,$ui_SecondBound,$ui_LocationSpecifier));
*/
 
/* 
//Search for substring based on boundaries - GET SUBSTRING (H:state)
$ui_ColInit =  "H";
$ui_FirstBound = '</a>';
$ui_SecondBound = '<br>';
$ui_LocationSpecifier = "state";
print_r($records->getLocationDetail($ui_FirstBound,$ui_SecondBound,$ui_LocationSpecifier));
*/ 
 
/* 
//Search for substring based on boundaries - GET SUBSTRING (I:locationnotes)
$ui_ColInit =  "I";
$ui_FirstBound = 'http://www.plantsystematics.org/reveal/pbio/digitalimages/';
$ui_SecondBound = '.html';
print_r($records->getStringBetweenCriteria($ui_FirstBound,$ui_SecondBound));

*/


//Open File, Set Column and Row position, Sheet name, and File name - SET EXCEL WORKSHEET

$ui_RowInit = "2";
$ui_SheetName = "Ma_Mz";
$ui_FileName = "PBIO Digital Images Index_Ma-Mz";
$ui_FileExtension = ".xlsx";
//$records->writeExcelSource($ui_ColInit,$ui_RowInit,$ui_SheetName,$ui_FileName,$ui_FileExtension);


// Set criteria to return contents only between specified HTML tags
//$records->setRegex("blockquote");

//Writes each record number for matching search criteria to Excel file
//$records->writeExcelRecord($records->arrayStringMatches);

//
//print_r($records->getBetweenTags($records->searchCriteria,$records->matches));

//Writes source code for each url based on record number inserted into URL
//$records->writeExcelSource($records->arrayWriteToExcel);


?>

