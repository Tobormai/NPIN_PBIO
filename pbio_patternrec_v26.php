
<?php
// the following line prevents the browser from parsing this as HTML.
//header('Content-Type: text/plain');
('Content-type: text/html; charset= iso-8859-1');  
header('Content-type: text/html; charset= UTF-8');  
require_once 'PHPExcel.php';
ini_set('max_execution_time', 800); //300 seconds = 5 minutes

class ExcelRecord
{
	//PROPERTIES
	private static $row = 1;
	private $col;
	
	public $activeUrl;
	public $arrayStringarrayStringMatches = array();
	public $arrayActiveUrl = array();
	public $arrayUrlContent = array();
	public $element;
	public $indexUrl;
	public $recordUrl;
	public $regexPattern;
	public $searchfor;	
	

	//CONSTRUCTORS
	public function __construct()
	  {
		  echo 'The class "', __CLASS__, '" was initiated!'. PHP_EOL;
	  }
	
	//GETTERS
	public function getSearchCritera()
	{
		return $this->searchfor;
	}
	
	/*public function getArrayMatches($found)
	{
		$this->arrayStringMatches = $found;
		$a = $found[0];
		array_walk($a,array($this, 'setRecordNumber'),".html pos: ");
	}*/
	
	public function getArrayElement($found)
	{
		$this->arrayStringMatches = $found;
		$a = $found[0];
		foreach ($a as $value)
		{
			 //echo "Record NUMBER: ";
			 $this->setRecordNumber($value);
			// echo "Value: $value<br />\n";
		}
	}
	
	public function getMatch($newSearch,$urlIndexBase)
	{
		$this->searchfor = $newSearch;
		$this->indexUrl =$urlIndexBase;
	
		// file_get_contents secessary to convert page contents to string$urlCurrent
		$subject = $urlCurrent= file_get_contents($urlIndexBase);

		// escape special characters in the query
		$regex = preg_quote($newSearch, '/');

		// finalise the regular expression, matching the whole line
		$regex = "/^.*$regex.*\$/m";

		// search, and store all matching occurences in $arrayStringMatches
		if(preg_match_all($regex, $subject, $found)){
		   echo "Found matches:\n";
		   echo implode("\n", $found[0]);
		}
		else{
			echo "No matches found";
		}
		
		$this->arrayStringMatches = $found;
		return $this->arrayStringMatches;
	}
	
	public function getBetweenTags($regexPattern,$arrayActiveUrl)
	{
				
		$this->setArrayRecordUrl($this->arrayStringMatches);
		$pattern = $this->regexPattern;
		echo "Search Criteria for REGEX: ".$this->regexPattern ."<br>";
		
		//$urlCurrent = $this->recordUrl;
		$found = $this->arrayActiveUrl;
		echo "imploding found array:";
		echo implode("\n", $found[0]);
		
		$this->matches = $found;
		$a = $found[0];
		foreach ($a as $value)
		{
			echo "Begin Value: ".$value."End of Value"."<br>";
			echo "Record Number (SOURCE LOOP): ";
			$this->setRecordNumber($value);	
	
			// Set Webpage URL				
			$urlBase = $this->setRecordUrl();
			echo "Record URL within LOOP: ".$this->recordUrl."<br>";
			$urlCurrent = file_get_contents($urlBase);
		
			$this->activeUrl = $urlCurrent;
			// escape special characters in the query
				//$regex = preg_quote($this->searchfor, '/');

				// finalise the regular expression, matching the whole line
				$regex = "~<blockquote\b[^>]*>(?:[^<]+|(?R)|<(?!/(?:blockquote|p)>))*</blockquote>~";

			// search, and store all matching occurences in $arrayStringMatches
			if(preg_match_all($regex, $urlCurrent, $foundRegex))
			{
			   echo "Found Regex matches:\n";
			   echo implode("\n", $foundRegex[0]);
			   //echo "Printing foundRegex inside of LOOP: ";
			  // print_r($foundRegex);		 
			}
			else
			{
				echo "No matches found";
			}
			$this->arrayUrlContent[] = $foundRegex;
			echo "Array Url Regex Content:\n";
			print_r( $this->arrayUrlContent);							
		}
			//return $this->arrayUrlContent;		
	}
	
	public function writeToExcel()
	{
		
	}
	
	//SETTERS
	public function setSearchCriteria($searchfor)
	{
		$this->searchfor = $searchfor; 
	}
	
	public function setRegex($pattern)
	{
		$this->regexPattern = $pattern; 		
	}

	public function setSearchUrl($urlIndexBase)
	{
		$this->indexUrl = $urlIndexBase; 
	}
	
	public function setRecordUrl()
	{
		$this->recordUrl = 'http://www.plantsystematics.org/reveal/pbio/digitalimages/'.$this->element.'.html';
		return $this->recordUrl;
	}
		
	public function setRecordNumber($value)
	{
		$startIndex = strpos($value, ".html") - 6;
		//echo "$startIndex <br>";
		$recordnum = substr($value, $startIndex, 6);	
		echo "Record NUMBER: "."$recordnum" ."<br>";	
		$this->element = $recordnum;
	}	
	
		public function setArrayRecordUrl($found)
	{
		$this->arrayStringMatches = $found;
		$a = $found[0];
		foreach ($a as $value)
		{			 
			 // Get record number from string and store as element used to build record URL
			 $this->setRecordNumber($value);	
			 
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
			
		$this->arrayStringMatches = $found;
		$a = $found[0];
		echo "Found Url Regex content to write to Excel \n";
		   echo implode("\n", $found[0]);
		
		foreach ($a as $value)
		{
			 //echo "Record Number (RECORDNUM LOOP): ";
			 $this->setRecordNumber($value);
						
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
	
	public function writeExcelSource($arrayUrlContent)
	{
		echo "Printing Excel Write Url Content Regex matches:\n";
		print_r ($arrayUrlContent);
		//$rowArray = $arrayUrlContent;
		$rowArray = array('red','white','blue');
		$columnArray = array_chunk($rowArray,1);
				
		
		// Create PHPExcel object which represents Excel Workbook
		$objPHPExcel = new PHPExcel();
		$objWriter = new PHPExcel_Writer_CSV($objPHPExcel);

		// Fill worksheet from values in array
		$objPHPExcel->getActiveSheet()->fromArray($columnArray, null, 'A2');				

		// Rename worksheet
		$objPHPExcel->getActiveSheet()->setTitle('Ia-Iz');
				
 
        // Save Excel 2007 file
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
        $objWriter->save('MyExcel.xlsx');
		//$file_put_contents($file,$urlCurrent);
	}
}
//Instantiate the WriteExcelRecords class.
$records = new ExcelRecord();

//Set serach criteria
$records->setSearchCriteria("Iberis");
//Set webpage on which to search
$records->setSearchUrl("http://www.plantsystematics.org/reveal/pbio/digitalimages/digslideindexI.html");
print 'Value of search critieria: ' .$records->getSearchCritera(). PHP_EOL;
print_r($records->getMatch($records->searchfor,$records->indexUrl));
//print_r($records->getArrayMatches($records->arrayStringMatches));
print_r($records->getArrayElement($records->arrayStringMatches));
//print_r($records->getBetweenTags($records->arrayStringMatches));
$records->setRegex("blockquote");

//Writes each record number for matching search criteria to Excel file
$records->writeExcelRecord($records->arrayStringMatches);

print_r($records->getBetweenTags($records->searchfor,$records->arrayActiveUrl));

//Writes source code for each url based on record number inserted into URL
$records->writeExcelSource($records->arrayUrlContent);


?>

