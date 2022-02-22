<?php

class PayslipWriter {
	private $db;

	public function __construct($db) {
		$this->db = $db;
	}
}
