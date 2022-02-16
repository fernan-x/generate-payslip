<?php
/* Copyright (C) 2001-2005 Rodolphe Quiedeville <rodolphe@quiedeville.org>
 * Copyright (C) 2004-2015 Laurent Destailleur  <eldy@users.sourceforge.net>
 * Copyright (C) 2005-2012 Regis Houssin        <regis.houssin@inodbox.com>
 * Copyright (C) 2015      Jean-François Ferry	<jfefe@aternatik.fr>
 *
 * This program is free software; you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation; either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program. If not, see <https://www.gnu.org/licenses/>.
 */

/**
 *	\file       generatepayslip/generatepayslipindex.php
 *	\ingroup    generatepayslip
 *	\brief      Home page of generatepayslip top menu
 */

// Load Dolibarr environment
$res = 0;
// Try main.inc.php into web root known defined into CONTEXT_DOCUMENT_ROOT (not always defined)
if (!$res && !empty($_SERVER["CONTEXT_DOCUMENT_ROOT"])) {
	$res = @include $_SERVER["CONTEXT_DOCUMENT_ROOT"]."/main.inc.php";
}
// Try main.inc.php into web root detected using web root calculated from SCRIPT_FILENAME
$tmp = empty($_SERVER['SCRIPT_FILENAME']) ? '' : $_SERVER['SCRIPT_FILENAME']; $tmp2 = realpath(__FILE__); $i = strlen($tmp) - 1; $j = strlen($tmp2) - 1;
while ($i > 0 && $j > 0 && isset($tmp[$i]) && isset($tmp2[$j]) && $tmp[$i] == $tmp2[$j]) {
	$i--; $j--;
}
if (!$res && $i > 0 && file_exists(substr($tmp, 0, ($i + 1))."/main.inc.php")) {
	$res = @include substr($tmp, 0, ($i + 1))."/main.inc.php";
}
if (!$res && $i > 0 && file_exists(dirname(substr($tmp, 0, ($i + 1)))."/main.inc.php")) {
	$res = @include dirname(substr($tmp, 0, ($i + 1)))."/main.inc.php";
}
// Try main.inc.php using relative path
if (!$res && file_exists("../main.inc.php")) {
	$res = @include "../main.inc.php";
}
if (!$res && file_exists("../../main.inc.php")) {
	$res = @include "../../main.inc.php";
}
if (!$res && file_exists("../../../main.inc.php")) {
	$res = @include "../../../main.inc.php";
}
if (!$res) {
	die("Include of main fails");
}

require_once DOL_DOCUMENT_ROOT.'/core/class/html.formfile.class.php';
dol_include_once('/generatepayslip/class/payslipdatagenerator.class.php');

// Load translation files required by the page
$langs->loadLangs(array("generatepayslip@generatepayslip"));

$action = GETPOST('action', 'aZ09');


// Security check
// if (! $user->rights->generatepayslip->myobject->read) {
// 	accessforbidden();
// }
$socid = GETPOST('socid', 'int');
if (isset($user->socid) && $user->socid > 0) {
	$action = '';
	$socid = $user->socid;
}

$max = 5;
$now = dol_now();


/*
 * Actions
 */

// None


/*
 * View
 */

$form = new Form($db);
$formfile = new FormFile($db);

llxHeader("", $langs->trans("GeneratePayslipArea"));

print load_fiche_titre($langs->trans("GeneratePayslipArea"), '', 'generatepayslip.png@generatepayslip');

require_once DOL_DOCUMENT_ROOT.'/core/modules/export/modules_export.php';

$workingDays = array('monday', 'tuesday', 'wednesday', 'thursday', 'friday');
$referenceUser = new User($db);
$user->fetch(2);
$datas = new PayslipDataGenerator($db, dol_now(), $workingDays, $user);

//$datas->generate();
/*$outputLangs = new Translate('', $conf);
$outputLangs->setDefaultLang('en_US');
$outputLangs->loadLangs(['main']);

$dayRow = array();

for ($i = $firstDay; $i < $lastDay; $i = strtotime('+1 day', $i)) {
	$row = array();
	$row['key'] = dol_strtolower(dol_print_date($i, '%A', 'auto', $outputLangs));
	$row['number'] = intval(dol_print_date($i, '%d', 'auto', $outputLangs));
	// Set worked day
	if (in_array($row['key'], $workingDay)) {
		$row['worked'] = true;
	} else {
		$row['worked'] = false;
	}

	// Add holiday to the worked day


	$dayRow[$i] = $row;
}

print '<pre>';
var_dump($dayRow);
print '</pre>';
*/

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\Date;

require_once DOL_DOCUMENT_ROOT.'/includes/phpoffice/phpspreadsheet/src/autoloader.php';
require_once DOL_DOCUMENT_ROOT.'/includes/Psr/autoloader.php';
dol_include_once('/generatepayslip/vendor/autoload.php');

// Create a copy of the spreadsheet
$sample = dol_buildpath('/generatepayslip/assets/sample.xls');
$destinationPath = $conf->generatepayslip->multidir_output[$conf->entity];
if (empty($destinationPath)) {
	$destinationPath = $conf->generatepayslip->dir_output;
}
$sampleCopy = $destinationPath.'/test.xlsx';
require_once DOL_DOCUMENT_ROOT.'/core/lib/files.lib.php';
$res = dol_copy($sample, $sampleCopy);

// Load spreadsheet
$spreadsheet = IOFactory::load($sampleCopy);


// Load the correct worksheet in function of the month
dol_include_once('/generatepayslip/class/code42spreadsheetparser.class.php');
$spreadsheetParser = new Code42SpreadsheetParser();
$month = $datas->getMonth();

$name = $spreadsheetParser->getWorksheetNameByMonth($month);
$worksheet = $spreadsheet->getSheetByName($name);
$rules = $spreadsheetParser->getRuleForWorksheet($name);

if ($worksheet && $rules) {
	// Debug date of the worksheet
	$date = $worksheet->getCell($rules['date'])->getValue();
	$date = Date::excelToTimestamp($date);
	var_dump(dol_print_date($date));

	// Fill date and year datas
	$worksheet->setCellValue($rules['date'], Date::stringToExcel(dol_print_date($datas->getFirstMonthDay(), '%Y-%m-%d')));
	$worksheet->setCellValue($rules['year'], $datas->getYear());

	// Fill enterprise name
	$worksheet->setCellValue($rules['company'], $mysoc->name);

	// Fill user name
	$worksheet->setCellValue($rules['username'],  $user->getFullName($langs));

	$daysData = $datas->getDaysData();
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
		}

		$row++;
	}

	$objWriter = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
	$objWriter->save($sampleCopy);

	$pdf = new \PhpOffice\PhpSpreadsheet\Writer\Pdf\Mpdf($spreadsheet);
	$pdf->setSheetIndex($spreadsheet->getIndex($worksheet));
	$pdf->save($sampleCopy.'.pdf');
	// apt-get install libreoffice
	// https://wiki.ubuntu.com/LibreOffice/fr
	// https://stackoverflow.com/questions/23223491/how-to-convert-xls-to-pdf-via-php
	//libreoffice --headless --convert-to pdf:calc_pdf_Export --outdir ../../documents/generatepayslip/ ../../documents/generatepayslip/test.xls

} else {
	print "Can't load worksheet ".$name;
}

// End of page
llxFooter();
$db->close();
