<?php

use h2g2\QueryBuilder;
use h2g2\QueryBuilderException;

dol_include_once('/holiday/class/holiday.class.php');
dol_include_once('/h2g2/class/querybuilder.class.php');

class PayslipDataGenerator {
	/**
	 * @var DoliDB 			Database handler
	 */
	private $db;

	/**
	 * @var User 			User
	 */
	private $user;

	/**
	 * @var int 			Reference day to base all datas on (timestamp)
	 */
	private $referenceDay;

	/**
	 * @var int 			Actual month (@example: 2)
	 */
	private $month;

	/**
	 * @var int 			Actual year (@example: 2022)
	 */
	private $year;

	/**
	 * @var int 			First day of the month (timestamp)
	 */
	private $firstMonthDay;

	/**
	 * @var int 			Last day of the month (timestamp)
	 */
	private $lastMonthDay;

	/**
	 * @var array|string[] 	List of working days (from monday to friday by default)
	 */
	private $workingDays;

	/**
	 * @var array			Holiday list for the user
	 */
	private $holidays;

	/**
	 * Construct the generator for the month of the day given
	 *
	 * @param 	DoliDB 		$db					Database handler
	 * @param 	int 		$referenceDay		Day used as reference (timestamp)
	 * @param 	array		$workingDays		List of working days as key (from monday to friday by default)
	 * @param 	User 		$user				User to base data on (global $user by default)
	 */
	public function __construct($db, $referenceDay, $workingDays = array(), $user = null)
	{
		$this->db = $db;
		$this->referenceDay = $referenceDay;

		// Get the working days
		if (empty($workingDays)) {
			$this->workingDays = array('monday', 'tuesday', 'wednesday', 'thursday', 'friday');
		} else {
			$this->workingDays = $workingDays;
		}

		// Get the connected user if user is not defined
		if (!$user) {
			global $user;
		}
		$this->user = $user;

		// Calculate each datas in function of the reference day
		$this->month = intval(dol_print_date($this->referenceDay, '%m')); // Get the date month as integer
		$this->year = intval(dol_print_date($this->referenceDay, '%Y')); // Get the date year as integer
		$this->firstMonthDay = dol_get_first_day($this->year, $this->month); // Get the first day of month
		$this->lastMonthDay = dol_get_last_day($this->year, $this->month); // Get the last day of month

		$this->fetchHolidays();
	}

	/**
	 * Fetch holidays
	 *
	 * @return array|null			List of holidays
	 * @throws Exception
	 */
	private function fetchHolidays()
	{
		global $conf;

		$holidays = null;

		try {
			$holidays = QueryBuilder::table('holiday AS t')
				->select('t.date_debut', 't.date_fin', 't.halfday', 'ht.code')
				->where([
					['t.entity', '=', $conf->entity],
					['t.fk_user', '=', $this->user->id],
					['t.statut', '=', Holiday::STATUS_APPROVED]
				])
				->join('c_holiday_types AS ht', 't.fk_type', 'ht.rowid')
				->get();
		} catch (QueryBuilderException $e) {
			dol_syslog('PayslipDataGenerator::fetchHolidays got an SQL error with the following request : '.$e->getRequest(), LOG_ERR);
		}

		return $holidays;
	}

	public function getDaysData()
	{
		global $conf;

		$outputLangs = new Translate('', $conf);
		$outputLangs->setDefaultLang('en_US');
		$outputLangs->loadLangs(['main']);

		$dayRow = array();

		for ($i = $this->firstMonthDay; $i <= $this->lastMonthDay; $i = strtotime('+1 day', $i)) {
			$row = array();
			$row['key'] = dol_strtolower(dol_print_date($i, '%A', 'auto', $outputLangs));
			$row['number'] = intval(dol_print_date($i, '%d', 'auto', $outputLangs));

			// Set worked day
			if (in_array($row['key'], $this->workingDays)) {
				$row['worked'] = true;
				$row['absenceReason'] = '';
			} else {
				$row['worked'] = false;
				$row['absenceReason'] = 'NWD'; // Not Working Day
			}

			$dayRow[$i] = $row;
		}

		// Add holiday to the worked day
		$holidays = $this->fetchHolidays();
		if ($holidays) {
			foreach ($holidays as $holiday) {
				$holidayStart = strtotime($holiday->date_debut);
				$holidayEnd = strtotime($holiday->date_fin);

				for ($i = $holidayStart; $i <= $holidayEnd; $i = strtotime('+1 day', $i)) { // Loop through each holiday day
					if ($dayRow[$i] && $dayRow[$i]['worked']) {
						$dayRow[$i]['worked'] = false;
						$dayRow[$i]['absenceReason'] = $holiday->code;
					}
				}

			}
		}

		return $dayRow;
//		print '<pre>';
//		var_dump($dayRow);
//		print '</pre>';
	}

	public function getMonth()
	{
		return $this->month;
	}

	public function getYear()
	{
		return $this->year;
	}

	public function getFirstMonthDay()
	{
		return $this->firstMonthDay;
	}
}
