<?php

namespace RendyRobbani\Cappuccino;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as XlsxWriter;

class JsonToExcel
{
	private static self $instance;

	public static function getInstance(): self
	{
		if (!isset(self::$instance)) self::$instance = new self();
		return self::$instance;
	}

	public function toExcel(string $fileJSON, string $fileXLSX): void
	{
		$resource = fopen($fileJSON, "r");
		$contents = fread($resource, filesize($fileJSON));
		fclose($resource);

		$spreadsheet = new Spreadsheet();
		$defaultStyle = $spreadsheet->getDefaultStyle();
		$defaultStyle->getFont()->setName("Aptos Narrow");
		$defaultStyle->getFont()->setSize(11);
		$defaultStyle->getAlignment()->setVertical(Alignment::VERTICAL_TOP);

		$worksheet = $spreadsheet->getActiveSheet();

		$r = 0;

		$rows = json_decode($contents, true);
		$max = sizeof($rows);
		for ($i = 0; $i < sizeof($rows); $i++) {
			$now = $i + 1;
			echo "\r";
			echo "Processing : " . ceil($now * 100 / $max) . "%";;

			$row = $rows[$i];
			if ($i == 0) {
				$r++;
				$c = 1;
				foreach (array_keys($row) as $key) {
					$worksheet->getCell([$c, $r])->setValue($key);
					$c++;
				}
			}

			$c = 1;
			$r++;
			foreach ($row as $key => $value) {
				if (is_string($value)) {
					$value = str_replace("\n", " ", $value);
					while (str_contains($value, "  ")) $value = str_replace("  ", " ", $value);
				}
				$worksheet->getCell([$c, $r])->setValue($value);
				$c++;
			}
		}

		$xlsx = new XlsxWriter($spreadsheet);
		$xlsx->save($fileXLSX);
	}
}