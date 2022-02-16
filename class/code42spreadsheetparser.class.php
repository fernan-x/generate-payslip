<?php

/**
 * Class to help parsing Code 42 payslip spreadsheet
 */
class Code42SpreadsheetParser {

	/**
	 * @var string[] 	List of available worksheets name
	 */
	private $whorsheets;

	private $loadingRules;

	public function __construct()
	{
		// List all worksheets of the spreadsheet
		$this->whorsheets = array(
			1 => 'janvier',
			2 => 'fevrier',
			3 => 'mars',
			4 => 'avril',
			5 => 'mai',
			6 => 'juin',
			7 => 'juillet',
			8 => 'aout',
			9 => 'septembre',
			10 => 'octobre',
			11 => 'novembre',
			12 => 'decembre'
		);

		// List of rules in function of worksheet
		$this->loadingRules = array(
			'janvier' => array(
				'row' => array('from' => 9, 'to' => 39),
				'column' => array('from' => 'B', 'to' => 'E')
			),
			'fevrier' => array(
				'date' => 'J1',
				'year' => 'L1',
				'company' => 'B1',
				'username' => 'B3',
				'row' => array('from' => 9, 'to' => 37),
				'hour_start_m' => 'B',
				'hour_end_m' => 'C',
				'hour_start_a' => 'D',
				'hour_end_a' => 'E',
				'holiday' => 'H',
				'eating_cost' => 'K',
				'driving_cost' => 'M',
				'other_cost' => 'N'
			),
		);
	}

	/**
	 * Get the worksheet to load in function of the month
	 *
	 * @param 	int 		$month				Month to get
	 * @return 	string							Worksheet name
	 */
	public function getWorksheetNameByMonth($month)
	{
		return $this->whorsheets[$month];
	}

	/**
	 * Get the parsing rules of the worksheet in function of its name
	 *
	 * @param 	string 		$name				Worksheet name
	 * @return 	array							Worksheet parsing rules
	 */
	public function getRuleForWorksheet($name) {
		return $this->loadingRules[$name];
	}
}
