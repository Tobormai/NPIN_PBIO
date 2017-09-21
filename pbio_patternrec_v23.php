
<?php
// the following line prevents the browser from parsing this as HTML.
header('Content-Type: text/plain');
//header('Content-type: text/html; charset= iso-8859-1');  
//header('Content-type: text/html; charset= UTF-8');  
require_once 'PHPExcel.php';
ini_set('max_execution_time', 800); //300 seconds = 5 minutes

class ExcelRecord
{
	//PROPERTIES
	private static $row = 1;
	private $col;
	public $matches = array();
	public $searchfor;
	public $indexUrl;
	public $recordUrl;
	public $activeUrl;
	public $element;

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
		$this->matches = $found;
		$a = $found[0];
		array_walk($a,array($this, 'setStartPos'),".html pos: ");
	}*/
	
	public function getArrayElement($found)
	{
		$this->matches = $found;
		$a = $found[0];
		foreach ($a as $value)
		{
			 echo "Record Number: ";
			 $this->setStartPos($value);
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

		// search, and store all matching occurences in $matches
		if(preg_match_all($regex, $subject, $found)){
		   echo "Found matches:\n";
		   echo implode("\n", $found[0]);
		}
		else{
			echo "No matches found";
		}
		
		$this->matches = $found;
		return $this->matches;
	}
	
	public function getBetweenTags($newTag,$activeUrl,$found)
	{
		$this->searchfor = $newTag;
		$urlCurrent = $this->activeUrl;
		echo "Search Criteria for REGEX: ".$this->searchfor ."<br>";
		
		$this->matches = $found;
		$a = $found[0];
		$this->matches = $found;
		$a = $found[0];
		
		foreach ($a as $value)
		{
			 echo "Record Number (SOURCE LOOP): ";
			 $this->setStartPos($value);	
	
		// Set Webpage URL		
		echo "Record URL within LOOP: ".$this->recordUrl;
		$urlBase = $this->setRecordUrl();
		$urlCurrent = file_get_contents($urlBase);
		echo "Active URL: ".$this->activeUrl."<br>";

		// escape special characters in the query
		//$regex = preg_quote($this->searchfor, '/');

		// finalise the regular expression, matching the whole line
		$regex = "~<blockquote\b[^>]*>(?:[^<]+|(?R)|<(?!/(?:blockquote|p)>))*</blockquote>~";

		// search, and store all matching occurences in $matches
		if(preg_match_all($regex, $urlCurrent, $found)){
		   echo "Found matches:\n";
		   echo implode("\n", $found[0]);
		}
		else{
			echo "No matches found";
		}
		
		$this->matches = $found;
		return $this->matches;	
		}
	}
	
	
	//SETTERS
	public function setSearchCriteria($searchfor)
	{
		$this->searchfor = $searchfor; 
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
		
	public function setStartPos($value)
	{
		$startIndex = strpos($value, ".html") - 6;
		//echo "$startIndex <br>";
		$recordnum = substr($value, $startIndex, 6);	
		echo "Record Number: "."$recordnum" ."<br>";	
		$this->element = $recordnum;
	}
	
	public function writeExcelRecord($found)
	{
		// Create PHPExcel object which represents Excel Workbook
		$excel = new PHPExcel();
	
		static $col = 'A';
		static $row = '2';
			
		$this->matches = $found;
		$a = $found[0];
		$this->matches = $found;
		$a = $found[0];
		foreach ($a as $value)
		{
			 echo "Record Number (RECORDNUM LOOP): ";
			 $this->setStartPos($value);
						
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
		// Create PHPExcel object which represents Excel Workbook
		$excel = new PHPExcel();
		$objWriter = new PHPExcel_Writer_CSV($excel);
		

		static $col = 'A';
		static $row = '2';
			
		$this->matches = $found;
		$a = $found[0];
		$this->matches = $found;
		$a = $found[0];
		foreach ($a as $value)
		{
			 echo "Record Number (SOURCE LOOP): ";
			 $this->setStartPos($value);	
	
		// Set Webpage URL		
		
		$urlBase = $this->setRecordUrl();
		echo "Record URL within LOOP: ".$this->recordUrl."<br>";
		$urlCurrent = file_get_contents($urlBase);
		
		$this->activeUrl = $urlCurrent;
		

		//echo "Active URL: ".$this->activeUrl."<br>";
		$urlCurrent = mb_convert_encoding($urlCurrent, 'ISO-8859-1', 'UTF-8');
		$encodedString = html_entity_decode($urlCurrent,ENT_XHTML,'UTF-8');
		//echo $encodedString;
		//$urlCurrent = str_replace('?','&deg',$urlCurrent); 
		//$urlCurrent = str_replace('°', '', $urlCurrent); 
		//$urlCurrent = preg_replace('/\?/', '&deg', $urlCurrent);
		
		// Write contents from Webpage to empty cell
		$excel->setActiveSheetIndex(0)
			->setCellValueByColumnAndRow($col,$row, $encodedString);

		// Increment Cell Value by 1 row
		$row++;
			
		}
		// Save Excel 2007 file
		$file = PHPExcel_IOFactory::createWriter($excel, 'Excel2007');
		$file->save('PBIO Raw Data_Source.xlsx');
		$file->save('PBIO Raw Data_Source.txt');
		
		$objWriter->setUseBOM(true);
		$objWriter->save("PBIO Raw Data_Source.csv");
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
//print_r($records->getArrayMatches($records->matches));
print_r($records->getArrayElement($records->matches));
//print_r($records->getBetweenTags($records->matches));
$records->setSearchCriteria("blockquote");




//Writes each record number for matching search criteria to Excel file
$records->writeExcelRecord($records->matches);


//Writes source code for each url based on record number inserted into URL
$records->writeExcelSource($records->matches);

print_r($records->getBetweenTags($records->searchfor,$records->activeUrl,$records->matches));

?>

