<?php

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

require_once DOL_DOCUMENT_ROOT.'/includes/phpoffice/phpspreadsheet/src/autoloader.php';
require_once DOL_DOCUMENT_ROOT.'/includes/Psr/autoloader.php';
require_once DOL_DOCUMENT_ROOT.'/core/lib/files.lib.php';

class PayslipWriter {
	/**
	 * @var DoliDB 		Database handler
	 */
	private $_db;

	/**
	 * @var Code42SpreadsheetParser		Spreadsheet parser
	 */
	private $_parser;

	/**
	 * @var PayslipData			Payslip datas
	 */
	private $_data;

	/**
	 * @var int 				Actual month (@example: 2)
	 */
	private $_month;

	private $_spreadsheet;

	public function __construct($db, $parser, $data) {
		$this->_db = $db;
		$this->_parser = $parser;
		$this->_data = $data;
		$this->_month = $data->getMonth();
	}

	/**
	 * Method executed before the spreadsheet writing.
	 * This will copy the spreadsheet given as a sample.
	 *
	 * @return string|null			Spreadsheet copy path
	 */
	private function _beforeWriting()
	{
		global $conf;

		// Create a copy of the sample spreadsheet
		$sample = dol_buildpath('/generatepayslip/assets/sample.xls');
		$destinationPath = $conf->generatepayslip->multidir_output[$conf->entity];
		if (empty($destinationPath)) {
			$destinationPath = $conf->generatepayslip->dir_output;
		}

		$filename = 'test.xslx'; // TODO : change the filename
		$sampleCopy = $destinationPath.'/'.$filename;

		$res = dol_copy($sample, $sampleCopy);

		return $res > 0 ? $sampleCopy : null;
	}

	/**
	 * Fill each days on the spreadsheet
	 *
	 * @param 	array 			$rules			Parser rules for the spreadsheet
	 * @param 	Worksheet 		$worksheet		Active worksheet to complete
	 * @return 	void
	 * @throws 	Exception
	 */
	private function _fillDaysData($rules, $worksheet)
	{
		$daysData = $this->_data->getDaysData();
		$row = $rules['row']['from'];
		foreach ($daysData as $day) {
			if ($row <= $rules['row']['to']) { // Avoid going outside of the row setup
				if ($day['worked']) {
					// Fill worked day
					$worksheet->setCellValue($rules['hour_start_m'].$row, '08:00');
					$worksheet->setCellValue($rules['hour_end_m'].$row, '12:00');
					$worksheet->setCellValue($rules['hour_start_a'].$row, '14:00');
					$worksheet->setCellValue($rules['hour_end_a'].$row, '17:00');
				} else if ($day['absenceReason'] != 'NWD') { // We avoid absence of type not working day
					// Fill absence column
					if ($day['absenceReason'] == 'LEAVE_PAID_FR') {
						$reason = 'CP';
					} else {
						$reason = $day['absenceReason'];
					}
					$worksheet->setCellValue($rules['holiday'].$row, $reason);
				}

				// Fill expenses
				if ($day['expenses']) {
					foreach ($day['expenses'] as $key => $value) {
						$worksheet->setCellValue($rules[$key].$row, $value);
					}
				}
			}

			$row++;
		}
	}

	/**
	 * Save the new spredsheet
	 *
	 * @param 	string 		$spreadsheetCopy		New spreadsheet path
	 * @return 	string|null							New spreadsheet path on success or null on error
	 * @throws 	Exception
	 */
	private function _save($spreadsheetCopy)
	{
		$ret = null;
		$objWriter = new Xlsx($this->_spreadsheet);

		try {
			$objWriter->save($spreadsheetCopy);
			$ret = $spreadsheetCopy;
		} catch (Exception $e) {
			dol_syslog('Can\'t save the worksheet'.$e->getMessage(), LOG_ERR);
		}

		return $ret;
	}

	/**
	 * Write the payslip for the month and the user
	 *
	 * @return string|null 			Path of the new spreadsheet or null
	 */
	public function write()
	{
		global $mysoc, $langs;

		$ret = null;

		// Execute actions before writing
		$spreadsheetCopy = $this->_beforeWriting();

		if ($spreadsheetCopy) {
			// Load spreadsheet
			$this->_spreadsheet = IOFactory::load($spreadsheetCopy);

			$name = $this->_parser->getWorksheetNameByMonth($this->_month);
			$worksheet = $this->_spreadsheet->getSheetByName($name);
			$rules = $this->_parser->getRuleForWorksheet($name);

			if ($worksheet && $rules) {
				// Fill date and year datas
				$worksheet->setCellValue($rules['date'], Date::stringToExcel(dol_print_date($this->_data->getFirstMonthDay(), '%Y-%m-%d')));
				$worksheet->setCellValue($rules['year'], $this->_data->getYear());

				// Fill enterprise name
				$worksheet->setCellValue($rules['company'], $mysoc->name);

				// Fill user name
				$loadedUser = $this->_data->getUser();
				if ($loadedUser) {
					$worksheet->setCellValue($rules['username'], $loadedUser->getFullName($langs));
				}

				// Fill days
				$this->_fillDaysData($rules, $worksheet);

				// Save the new worksheet
				$ret = $this->_save($spreadsheetCopy);
			} else {
				dol_syslog('Can\'t load worksheet or rules for '.$name, LOG_ERR);
			}
		} else {
			dol_syslog('Can\'t create a copy of the sample', LOG_ERR);
		}

		return $ret;
	}
}
