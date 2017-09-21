
<?php
// the following line prevents the browser from parsing this as HTML.
header('Content-Type: text/plain');
('Content-type: text/html; charset= iso-8859-1');  
//header('Content-type: text/html; charset= UTF-8');  
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
				
		$pattern = $this->regexPattern;
		echo "Search Criteria for REGEX: ".$this->regexPattern ."<br>";
		
		//$urlCurrent = $this->recordUrl;
		$found = $this->arrayActiveUrl;
		echo "imploding found array:";
		echo implode("\n", $found[0]);
		print_r($found);
		
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
			if(preg_match_all($regex, $urlCurrent, $found))
			{
			   echo "Found Regex matches:\n";
			   echo implode("\n", $found);
			}
			else
			{
				echo "No matches found";
			}		
		}
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
	
	public function setArrayRecordUrl()
	{
		
	}
	
	public function writeExcelRecord($found)
	{
		// Create PHPExcel object which represents Excel Workbook
		$excel = new PHPExcel();
	
		static $col = 'A';
		static $row = '2';
			
		$this->arrayStringMatches = $found;
		$a = $found[0];
		
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
	
	public function writeExcelSource($found)
	{
		$this->arrayStringMatches = $found;
		$a = $found[0];
		foreach ($a as $value)
		{			 
			 // Get record number from string and store as element used to build record URL
			 //$this->setRecordNumber($value);	
			 
			// Set each Webpage URL	using record number		
			$this->setRecordUrl();
			echo "Record URL: ".$this->recordUrl."<br>";
			$value = $this->recordUrl;
			echo "URLBASE: ".$value."<br>";
			
			// Store each Record Urls to array
			$found[]= $value;
			$this->arrayActiveUrl = $found;			
		}
		print_r($found);		
		return $this->arrayActiveUrl;	
		
	}


}
//Instantiate the WriteExcelRecords class.
$records = new ExcelRecord();

//Set serach criteria
$records->setSearchCriteria("Mal");
//Set webpage on which to search
$records->setSearchUrl("http://www.plantsystematics.org/reveal/pbio/digitalimages/digslideindexM.html");
print 'Value of search critieria: ' .$records->getSearchCritera(). PHP_EOL;
print_r($records->getMatch($records->searchfor,$records->indexUrl));
//print_r($records->getArrayMatches($records->arrayStringMatches));
print_r($records->getArrayElement($records->arrayStringMatches));
//print_r($records->getBetweenTags($records->arrayStringMatches));
$records->setRegex("blockquote");



//Writes each record number for matching search criteria to Excel file
$records->writeExcelRecord($records->arrayStringMatches);


//Writes source code for each url based on record number inserted into URL
$records->writeExcelSource($records->arrayStringMatches);

print_r($records->getBetweenTags($records->searchfor,$records->arrayActiveUrl));

?>

