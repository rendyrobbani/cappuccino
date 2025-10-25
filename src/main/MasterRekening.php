<?php

namespace RendyRobbani\Cappuccino;

use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as XlsxReader;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as XlsxWriter;

class MasterRekening
{
	private static self $instance;

	public static function getInstance(): self
	{
		if (!isset(self::$instance)) self::$instance = new self();
		return self::$instance;
	}

	public function create(string $from, string $into): void
	{
		$fileSize = filesize($from);
		$resource = fopen($from, "r");
		$contents = fread($resource, $fileSize);
		fclose($resource);

		$contents = json_decode($contents, true);

		$list = [];
		foreach ($contents as $content) {
			$code = $content["kode_akun"];
			$name = $content["nama_akun"];
			while (str_contains($name, "/ ")) $name = str_replace("/ ", "/", $name);
			$level = match (strlen($code)) {
				1 => 1,
				3 => 2,
				6 => 3,
				9 => 4,
				13 => 5,
				default => 6,
			};
			if (str_starts_with($code, "4")) $list[] = ["code" => $code, "name" => $name, "level" => $level];
		}

		usort($list, fn($a, $b) => strcmp("|" . $a["code"], "|" . $b["code"]));

		$spreadsheet = new Spreadsheet();
		$defaultStyle = $spreadsheet->getDefaultStyle();
		$defaultStyle->getFont()->setName("Aptos Narrow");
		$defaultStyle->getFont()->setSize(11);
		$defaultStyle->getAlignment()->setVertical(Alignment::VERTICAL_TOP);

		$worksheet = $spreadsheet->getActiveSheet();
		$worksheet->getRowDimension(1)->setRowHeight(30);
		$worksheet->getColumnDimension("A")->setWidth(10);
		$worksheet->getColumnDimension("B")->setWidth(20);
		$worksheet->getColumnDimension("C")->setWidth(60);
		$worksheet->getColumnDimension("D")->setWidth(20);
		$worksheet->getColumnDimension("E")->setWidth(20);
		$worksheet->getColumnDimension("F")->setWidth(20);
		$worksheet->getColumnDimension("G")->setWidth(20);

		$rowNum = 1;
		$colNum = 1;

		foreach (explode("|", "Level|Kode|Uraian|APBD 2025|Rancangan\nAPBD 2026|Bertambah/\n(Berkurang)|Keterangan") as $colName) {
			$cell = $worksheet->getCell([$colNum, $rowNum]);
			$style = $cell->getStyle();
			$style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
			$style->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
			$style->getAlignment()->setWrapText(true);
			$style->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
			$style->getFill()->setFillType(Fill::FILL_SOLID);
			$style->getFill()->setStartColor(new Color("6d28d9"));
			$style->getFont()->setBold(true);
			$style->getFont()->setColor(new Color(Color::COLOR_WHITE));
			$cell->setValue($colName);
			$colNum++;
		}

		$rowNum++;

		$maps = [];

		foreach ($list as $data) {
			$code = $data["code"];
			$name = $data["name"];
			$level = $data["level"];

//			if ($code == "4.1.01.03") break;
			if ($level <= 4) {
				$rowNum++;
				for ($colNum = 1; $colNum <= 7; $colNum++) $worksheet->getStyle([$colNum, $rowNum])->getFont()->setBold(true);
			}

			if (!isset($maps[$level])) $maps[$level] = [];
			$maps[$level][$code] = $rowNum;

			$worksheet->getCell("A$rowNum")->setValueExplicit($level, DataType::TYPE_NUMERIC);
			$worksheet->getCell("B$rowNum")->setValueExplicit($code, DataType::TYPE_STRING);
			$worksheet->getCell("C$rowNum")->setValueExplicit($name, DataType::TYPE_STRING);
			$worksheet->getCell("D$rowNum")->setValueExplicit(0, DataType::TYPE_NUMERIC);
			$worksheet->getCell("E$rowNum")->setValueExplicit(0, DataType::TYPE_NUMERIC);
			$worksheet->getCell("F$rowNum")->setValueExplicit("=E$rowNum-D$rowNum", DataType::TYPE_FORMULA);
			$rowNum++;
		}

		for ($l = 1; $l < 6; $l++) {
			$data1 = $maps[$l];
			$data2 = $maps[$l + 1];

			foreach ($data1 as $code1 => $rowNum1) {
				$sebelum = [];
				$setelah = [];
				foreach ($data2 as $code2 => $rowNum2) {
					if (str_starts_with($code2, $code1)) {
						$sebelum[] = "D$rowNum2";
						$setelah[] = "E$rowNum2";
					}
				}
				if (sizeof($sebelum) > 0) $worksheet->getCell("D$rowNum1")->setValueExplicit("=" . implode("+", $sebelum), DataType::TYPE_FORMULA);
				if (sizeof($setelah) > 0) $worksheet->getCell("E$rowNum1")->setValueExplicit("=" . implode("+", $setelah), DataType::TYPE_FORMULA);
			}
		}

		for ($rowNum = 2; $rowNum <= $worksheet->getHighestRow(); $rowNum++) {
			for ($colNum = 1; $colNum <= 7; $colNum++) {
				$cell = $worksheet->getCell([$colNum, $rowNum]);
				$style = $cell->getStyle();
				if ($colNum == 1) $style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
				if ($colNum == 3) $style->getAlignment()->setWrapText(true);
				$style->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);

				if (in_array($colNum, [4, 5, 6])) $style->getNumberFormat()->setFormatCode("#,##0.00_);[Red](#,##0.00)");

				$level = $worksheet->getCell("A$rowNum")->getValue();
				if ($level != null && $level <= 4) $style->getFont()->setBold(true);
			}
		}

		$xlsx = new XlsxWriter($spreadsheet);
		$xlsx->save($into);
	}
}