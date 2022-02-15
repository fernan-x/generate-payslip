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

use PhpOffice\PhpSpreadsheet\IOFactory;
require_once DOL_DOCUMENT_ROOT.'/includes/phpoffice/phpspreadsheet/src/autoloader.php';
require_once DOL_DOCUMENT_ROOT.'/includes/Psr/autoloader.php';
require_once PHPEXCELNEW_PATH.'Spreadsheet.php';
$spreadsheet = IOFactory::load(dol_buildpath('/generatepayslip/assets/sample.xls'));

$now = dol_now();

$month = intval(dol_print_date($now, '%m')); // Get the date month as integer
$year = intval(dol_print_date($now, '%Y')); // Get the date year as integer
$firstDay = dol_get_first_day($year, $month); // Get the first day of month
$lastDay = dol_get_last_day($year, $month); // Get the last day of month

$workingDay = array('monday', 'tuesday', 'wednesday', 'thursday', 'friday'); // Working days

$outputLangs = new Translate('', $conf);
$outputLangs->setDefaultLang('en_US');
$outputLangs->loadLangs(['main']);
var_dump(dol_print_date($firstDay, '%A %d', 'auto', $outputLangs));
var_dump(dol_print_date($lastDay, '%A %d', 'auto', $outputLangs));

$dayRow = array();

for ($i = $firstDay; $i < $lastDay; $i = strtotime('+1 day', $i)) {
	$dayRow[] = dol_print_date($i, '%A %d', 'auto', $outputLangs);
}

var_dump($dayRow);

// Load the correct worsheet in function of the month
dol_include_once('/generatepayslip/class/code42spreadsheetparser.class.php');
$spreadsheetParser = new Code42SpreadsheetParser();
$name = $spreadsheetParser->getWorksheetNameByMonth($month);
$worksheet = $spreadsheet->getSheetByName($name);
$rules = $spreadsheetParser->getRuleForWorksheet($name);

if ($worksheet && $rules) {
	$date = $worksheet->getCell($rules['date'])->getValue();
	$date = \PhpOffice\PhpSpreadsheet\Shared\Date::excelToTimestamp($date);

	$year = $worksheet->getCell($rules['year'])->getValue();

	var_dump($date, $year);
} else {
	print "Can't load worksheet ".$name;
}

// End of page
llxFooter();
$db->close();
