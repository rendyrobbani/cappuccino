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

class Anggaran
{
	private static self $instance;

	public static function getInstance(): self
	{
		if (!isset(self::$instance)) self::$instance = new self();
		return self::$instance;
	}

	private const array HEADER = [
		"Kode SKPD" => 5,
		"Nama SKPD" => 6,
		"Kode Unit SKPD" => 7,
		"Nama Unit SKPD" => 8,
		"Kode Urusan" => 3,
		"Nama Urusan" => 4,
		"Kode Bidang" => 9,
		"Nama Bidang" => 10,
		"Kode Program" => 11,
		"Nama Program" => 12,
		"Kode Kegiatan" => 13,
		"Nama Kegiatan" => 14,
		"Kode Subkegiatan" => 15,
		"Nama Subkegiatan" => 16,
		"Kode Sumber Dana" => 17,
		"Nama Sumber Dana" => 18,
		"Kode Rekening" => 19,
		"Nama Rekening" => 20,
		"Anggaran" => 21,
	];

	private XlsxReader $reader;

	private function __construct()
	{
		$this->reader = new XlsxReader();
	}

	/**
	 * @param string $fileName
	 * @return array
	 * @throws \Exception
	 */
	private function read(string $fileName): array
	{
		if (!file_exists($fileName)) throw new \Exception("File not found");

		$values = [];
		$spreadsheet = $this->reader->load($fileName);
		$worksheet = $spreadsheet->getActiveSheet();
		$cols = array_values(self::HEADER);
		for ($row = 2; $row <= $worksheet->getHighestRow(); $row++) {
			$v = $worksheet->getCell([19, $row])->getValue();
			if ($v != null && strlen($v) > 0) {
				$value = [];
				foreach ($cols as $col) $value[] = $worksheet->getCell([$col, $row])->getValue();
				$values[] = $value;
			}
		}
		return $values;
	}

	private function write(string $fileName, array $values): void
	{
		$spreadsheet = new Spreadsheet();
		$defaultStyle = $spreadsheet->getDefaultStyle();
		$defaultStyle->getFont()->setName("Aptos Narrow");
		$defaultStyle->getFont()->setSize(11);
		$defaultStyle->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);

		$worksheet = $spreadsheet->getActiveSheet();
		$row = 1;
		$col = 1;
		foreach (array_keys(self::HEADER) as $header) {
			$cell = $worksheet->getCell([$col, $row]);
			$style = $cell->getStyle();
			$style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
			$style->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
			$style->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
			$style->getFill()->setFillType(Fill::FILL_SOLID);
			$style->getFill()->setStartColor(new Color("6d28d9"));
			$style->getFont()->setBold(true);
			$style->getFont()->setColor(new Color(Color::COLOR_WHITE));

			$cell->setValue($header);
			$col++;
		}

		foreach ($values as $value) {
			$row++;
			$col = 1;
			foreach ($value as $val) {
				$cell = $worksheet->getCell([$col, $row]);
				switch ($col) {
					case 19:
						$cell->setValueExplicit($val, DataType::TYPE_NUMERIC);
						break;
					default:
						$cell->setValueExplicit($val, DataType::TYPE_STRING);
						break;
				}
				$style = $cell->getStyle();
				$style->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
				$col++;
			}
		}

		$xlsx = new XlsxWriter($spreadsheet);
		$xlsx->save($fileName);
	}

	/**
	 * @param string $fileName
	 * @return void
	 * @throws \Exception
	 */
	public function recap(string $fileName): void
	{
		$fileName = str_replace("\\", "/", $fileName);
		$root = substr($fileName, 0, strrpos($fileName, "/"));

		echo "Read : $fileName";
		echo PHP_EOL;

		$values = $this->read($fileName);
		echo "Write : Laporan Anggaran.xlsx";
		echo PHP_EOL;
		$this->write("$root/Laporan Anggaran.xlsx", $values);
	}
}