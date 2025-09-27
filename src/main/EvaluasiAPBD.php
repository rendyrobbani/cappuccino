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

class EvaluasiAPBD
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
		"Kode Rekening" => 19,
		"Nama Rekening" => 20,
		"Jumlah" => 21,
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
		for ($row = 2; $row <= $worksheet->getHighestRow(); $row++) {
			$v = $worksheet->getCell([19, $row])->getValue();
			if ($v != null && strlen($v) > 0) $values[] = array_map(fn($val) => $worksheet->getCell([$val, $row])->getValue(), self::HEADER);
		}
		return $values;
	}

	/**
	 * @param string $fileName
	 * @param array $values
	 * @param bool $formatted
	 * @return void
	 */
	private function write(string $fileName, array $values, bool $formatted = false): void
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
			if ($header != "Jumlah") $cell->setValue($header);
			else {
				$worksheet->getCell([$col++, $row])->setValue("Sebelum Evaluasi");
				$worksheet->getCell([$col++, $row])->setValue("Setelah Evaluasi");
				$worksheet->getCell([$col++, $row])->setValue("Bertambah/Berkurang");
			}
			$col++;
		}

		if ($formatted) {
			for ($i = 1; $i < sizeof(self::HEADER) + 3; $i++) {
				$cell = $worksheet->getCell([$i, $row]);
				$style = $cell->getStyle();
				$style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
				$style->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
				$style->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
				$style->getFill()->setFillType(Fill::FILL_SOLID);
				$style->getFill()->setStartColor(new Color("6d28d9"));
				$style->getFont()->setBold(true);
				$style->getFont()->setColor(new Color(Color::COLOR_WHITE));
			}
		}

		foreach ($values as $value) {
			$row++;
			$col = 1;

			foreach ($value as $key => $val) {
				$cell = $worksheet->getCell([$col, $row]);
				switch ($key) {
					case "Sebelum Evaluasi":
					case "Setelah Evaluasi":
					case "Bertambah/Berkurang":
						$cell->setValueExplicit($val, DataType::TYPE_NUMERIC);
						break;
					default:
						$cell->setValueExplicit($val, DataType::TYPE_STRING);
						break;
				}
				if ($formatted) {
					$style = $cell->getStyle();
					$style->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
				}
				$col++;
			}
		}

		$xlsx = new XlsxWriter($spreadsheet);
		$xlsx->save($fileName);
	}

	/**
	 * @param string $fileSebelum
	 * @param string $fileSetelah
	 * @param string $fileName
	 * @param bool $formatted
	 * @return void
	 * @throws \Exception
	 */
	public function recap(string $fileSebelum, string $fileSetelah, string $fileName, bool $formatted = false): void
	{
		echo "Read : $fileSebelum";
		echo PHP_EOL;
		$sebelum = $this->read($fileSebelum);

		echo "Read : $fileSetelah";
		echo PHP_EOL;
		$setelah = $this->read($fileSetelah);

		$data = [];
		$step = 0;
		foreach ([$sebelum, $setelah] as $from) {
			$step++;
			foreach ($from as $row) {
				$key = implode("|", [$row["Kode Unit SKPD"], $row["Kode Subkegiatan"], $row["Kode Rekening"]]);

				if (!isset($data[$key])) {
					$temp = [];
					foreach (array_keys(self::HEADER) as $header) {
						if ($header != "Jumlah") $temp[$header] = $row[$header];
						else {
							if ($step == 1) {
								$temp["Sebelum Evaluasi"] = floatval($row["Jumlah"]);
								$temp["Setelah Evaluasi"] = 0.00;
							}
							if ($step == 2) {
								$temp["Sebelum Evaluasi"] = 0.00;
								$temp["Setelah Evaluasi"] = floatval($row["Jumlah"]);
							}
						}
					}
				} else {
					$temp = $data[$key];
					if ($step == 1) $temp["Sebelum Evaluasi"] += floatval($row["Jumlah"]);
					if ($step == 2) $temp["Setelah Evaluasi"] += floatval($row["Jumlah"]);
				}
				$data[$key] = $temp;
			}
		}

		echo "Write : Rekap Rancangan Perubahan APBD (Penyesuaian Hasil Evaluasi).xlsx";
		echo PHP_EOL;
		$this->write("$fileName/Rekap Rancangan Perubahan APBD (Penyesuaian Hasil Evaluasi).xlsx", $data, $formatted);
	}
}