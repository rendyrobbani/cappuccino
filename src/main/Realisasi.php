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

class Realisasi
{
	private static self $instance;

	public static function getInstance(): self
	{
		if (!isset(self::$instance)) self::$instance = new self();
		return self::$instance;
	}

	private const array HEADER = [
		"Kode SKPD",
		"Nama SKPD",
		"Kode Unit SKPD",
		"Nama Unit SKPD",
		"Kode Fungsi",
		"Nama Fungsi",
		"Kode Subfungsi",
		"Nama Subfungsi",
		"Kode Urusan",
		"Nama Urusan",
		"Kode Bidang",
		"Nama Bidang",
		"Kode Program",
		"Nama Program",
		"Kode Kegiatan",
		"Nama Kegiatan",
		"Kode Subkegiatan",
		"Nama Subkegiatan",
		"Kode Rekening",
		"Nama Rekening",
		"Nomor Dokumen",
		"Jenis Dokumen",
		"Jenis Transaksi",
		"Nomor DPT",
		"Tanggal Dokumen",
		"Keterangan Dokumen",
		"Nilai Realisasi",
		"Nilai Setoran",
		"NIP Pegawai",
		"Nama Pegawai",
		"Tanggal Simpan",
		"Nomor SPD",
		"Periode SPD",
		"Nilai SPD",
		"Tahapan SPD",
		"Nama Sub Tahapan Jadwal",
		"Tahapan APBD",
		"Nomor SPP",
		"Tanggal SPP",
		"Nomor SPM",
		"Tanggal SPM",
		"Nomor SP2D",
		"Tanggal SP2D",
		"Tanggal Transfer",
		"Nilai SP2D",
	];

	private XlsxReader $reader;

	private function __construct()
	{
		$this->reader = new XlsxReader();
	}

	private function toNumber(string $from): string
	{
		if ($from == null || $from == "null") return "0";
		return str_replace(",", ".", str_replace(".", "", $from));
	}

	private function toDate(string $from): null|string
	{
		if ($from == null || $from == "" || $from == "null") return null;
		$explode = explode(" ", $from);
		$into = [];
		$into[] = str_pad($explode[2], 4, "0", STR_PAD_LEFT);
		$into[] = str_pad(match ($explode[1]) {
			"Januari" => "01",
			"Februari" => "02",
			"Maret" => "03",
			"April" => "04",
			"Mei" => "05",
			"Juni" => "06",
			"Juli" => "07",
			"Agustus" => "08",
			"September" => "09",
			"Oktober" => "10",
			"November" => "11",
			"Desember" => "12",
		}, 2, "0", STR_PAD_LEFT);
		$into[] = str_pad($explode[0], 2, "0", STR_PAD_LEFT);
		return implode("-", $into);
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
		for ($row = 1; $row <= $worksheet->getHighestRow(); $row++) {
			$cell = $worksheet->getCell([1, $row]);
			if (preg_match("#\\d#", $cell->getValue())) {
				$value = [];
				for ($col = 2; $col <= 46; $col++) {
					$val = $worksheet->getCell([$col, $row]);
					$value[] = match ($col) {
						28, 29, 35, 46 => $this->toNumber($val),
						26, 32, 40, 44, 45 => $this->toDate($val),
						default => $val,
					};
				}
				$values[] = $value;
			}
		}
		return $values;
	}

	/**
	 * @param string $fileName
	 * @param array[] $values
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
		foreach (self::HEADER as $header) {
			$cell = $worksheet->getCell([$col, $row]);
			if ($formatted) {
				$style = $cell->getStyle();
				$style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
				$style->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
				$style->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
				$style->getFill()->setFillType(Fill::FILL_SOLID);
				$style->getFill()->setStartColor(new Color("6d28d9"));
				$style->getFont()->setBold(true);
				$style->getFont()->setColor(new Color(Color::COLOR_WHITE));
			}

			$cell->setValue($header);
			$col++;
		}

		foreach ($values as $value) {
			$row++;
			$col = 1;
			foreach ($value as $val) {
				$cell = $worksheet->getCell([$col, $row]);
				switch ($col) {
					case 28 - 1:
					case 29 - 1:
					case 35 - 1:
					case 46 - 1:
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
	 * @param string $root
	 * @param bool $formatted
	 * @return void
	 * @throws \Exception
	 */
	public function join(string $root, bool $formatted = false): void
	{
		$values = [];
		if (!file_exists($root)) throw new \Exception("File not found");
		$fileNames = scandir($root);
		$fileNames = array_filter($fileNames, fn($fileName) => str_ends_with(strtolower($fileName), ".xlsx"));
		$fileNames = array_filter($fileNames, fn($fileName) => is_file("$root/$fileName"));
		$fileNames = array_filter($fileNames, fn($fileName) => preg_match("#\\d{2}\\. Laporan Realisasi.xlsx#", $fileName));
		foreach ($fileNames as $fileName) {
			echo "Read : " . $fileName;
			echo PHP_EOL;
			$values = array_merge($values, $this->read("$root/$fileName"));
		}

		echo "Write : Laporan Realisasi.xlsx";
		echo PHP_EOL;
		$this->write("$root/Laporan Realisasi.xlsx", $values, $formatted);
	}

	/**
	 * @param string $root
	 * @return void
	 * @throws \Exception
	 */
	public function rename(string $root): void
	{
		if (!file_exists($root)) throw new \Exception("File not found");
		$fileNames = scandir($root);
		$fileNames = array_filter($fileNames, fn($fileName) => str_ends_with($fileName, ".xlsx"));
		$fileNames = array_filter($fileNames, fn($fileName) => preg_match("#Laporan Realisasi.*\\.xlsx#", $fileName));
		sort($fileNames);
		for ($i = 0; $i < sizeof($fileNames); $i++) {
			$fileName = $fileNames[$i];

			$number = 1;
			if (preg_match("#Laporan Realisasi \\((\\d)\\)\\.xlsx#", $fileName, $matches)) $number = intval($matches[1]) + 1;
			$number = str_pad($number, 2, "0", STR_PAD_LEFT);

			$from = "$root/$fileName";
			$into = "$root/$number. Laporan Realisasi.xlsx";
			rename($from, $into);
		}
	}
}