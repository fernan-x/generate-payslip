<?php

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\Date;
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

	private function _writeHolidays()
	{

	}

	private function _writeExpenses()
	{

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

		$sampleCopy = $destinationPath.'/test.xlsx';

		$res = dol_copy($sample, $sampleCopy);

		return $res > 0 ? $sampleCopy : null;
	}

	/**
	 * Write the payslip for the month
	 *
	 * @return void
	 */
	public function write()
	{
		global $mysoc, $langs;

		$spreadsheetCopy = $this->_beforeWriting();

		if ($spreadsheetCopy) {
			// Load spreadsheet
			$this->_spreadsheet = IOFactory::load($spreadsheetCopy);

			$name = $this->_parser->getWorksheetNameByMonth($this->_month);
			$worksheet = $this->_spreadsheet->getSheetByName($name);
			$rules = $this->_parser->getRuleForWorksheet($name);

			if ($worksheet && $rules) {
				// Debug date of the worksheet
				$date = $worksheet->getCell($rules['date'])->getValue();
				$date = Date::excelToTimestamp($date);
				var_dump(dol_print_date($date));

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

				$objWriter = new Xlsx($this->_spreadsheet);
				$objWriter->save($spreadsheetCopy);
			} else {
				print "Can't load worksheet ".$name;
			}
		}

	}
}
