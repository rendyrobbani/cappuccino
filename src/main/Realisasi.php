<?php

namespace RendyRobbani\Cappuccino;

use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as XlsxReader;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as XlsxWriter;

class Realisasi
{
	private const array HEADER = ["Kode SKPD", "Nama SKPD", "Kode Sub SKPD", "Nama Sub SKPD", "Kode Fungsi", "Nama Fungsi", "Kode Sub Fungsi", "Nama Sub Fungsi", "Kode Urusan", "Nama Urusan", "Kode Bidang Urusan", "Nama Bidang Urusan", "Kode Program", "Nama Program", "Kode Kegiatan", "Nama Kegiatan", "Kode Sub Kegiatan", "Nama Sub Kegiatan", "Kode Rekening", "Nama Rekening", "Nomor Dokumen", "Jenis Dokumen", "Jenis Transaksi", "Nomor DPT", "Tanggal Dokumen", "Keterangan Dokumen", "Nilai Realisasi", "Nilai Setoran", "NIP Pegawai", "Nama Pegawai", "Tanggal Simpan", "Nomor SPD", "Periode SPD", "Nilai SPD", "Tahapan SPD", "Nama Sub Tahapan Jadwal", "Tahapan APBD", "Nomor SPP", "Tanggal SPP", "Nomor SPM", "Tanggal SPM", "Nomor SP2D", "Tanggal SP2D", "Tanggal Transfer", "Nilai SP2D"];

	private static Realisasi $instance;

	public static function getInstance(): self
	{
		if (!isset(self::$instance)) self::$instance = new self();
		return self::$instance;
	}

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
	 * @return void
	 */
	private function write(string $fileName, array $values): void
	{
		$spreadsheet = new Spreadsheet();
		$worksheet = $spreadsheet->getActiveSheet();
		$row = 1;
		$col = 1;
		foreach (self::HEADER as $header) {
			$worksheet->getCell([$col, $row])->setValue($header);
			$col++;
		}

		foreach ($values as $value) {
			$row++;
			$col = 1;
			foreach ($value as $val) {
				switch ($col) {
					case 28 - 1:
					case 29 - 1:
					case 35 - 1:
					case 46 - 1:
						$worksheet->getCell([$col, $row])->setValueExplicit($val, DataType::TYPE_NUMERIC);
						break;
					default:
						$worksheet->getCell([$col, $row])->setValueExplicit($val, DataType::TYPE_STRING);
						break;
				}
				$col++;
			}
		}

		$xlsx = new XlsxWriter($spreadsheet);
		$xlsx->save($fileName);
	}

	/**
	 * @param string $root
	 * @return void
	 * @throws \Exception
	 */
	public function join(string $root): void
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
		$this->write("$root/Laporan Realisasi.xlsx", $values);
	}
}