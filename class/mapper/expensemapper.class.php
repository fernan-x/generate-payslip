<?php

/**
 * Class used to map dolibarr expenses type to spreadsheet expenses type
 */
class ExpenseMapper {
	const MAPPER = array(
		'TF_LUNCH' => 'eating_cost',
		'EX_KME' => 'driving_cost',
		'TF_TRIP' => 'other_cost',
		'TF_OTHER' => 'other_cost'
	);

	const AVAILABLE_MAP = array(
		'TF_LUNCH',
		'TF_OTHER',
		'TF_TRIP',
		'EX_KME'
	);

	const FALLBACK_MAP = 'other_cost';

	/**
	 * Map a dolibarr expense code to the spreadsheet rule key
	 *
	 * @param 	string 			$code			Dolibarr code
	 * @return 	string 							Rule key
	 */
	public static function toSpreadSheet($code)
	{
		if (in_array($code, self::AVAILABLE_MAP)) {
			$ret = self::MAPPER[$code];
		} else {
			$ret = self::FALLBACK_MAP;
		}

		return $ret;
	}
}
