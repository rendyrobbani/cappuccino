<?php

namespace RendyRobbani\Cappuccino;

use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as XlsxWriter;

class TemplatePPAS
{
	private static self $instance;

	public static function getInstance(): self
	{
		if (!isset(self::$instance)) self::$instance = new self();
		return self::$instance;
	}

	private static array $KATEGORI = [
		"BELANJA WAJIB",
		"BELANJA MENGIKAT",
		"BELANJA PRIORITAS",
		"BELANJA PENUNJANG KESEKRETARIATAN",
		"BELANJA PROGRAM KEGIATAN",
	];

	private static string $KEY_SKPD = "1. SKPD";

	private static string $KEY_BIDANG = "2. Bidang";

	private static string $KEY_UNIT = "3. Unit";

	private static string $KEY_KATEGORI = "4. Kategori";

	private static string $KEY_URAIAN = "5. Uraian";

	private static string $KEY_SUMBER = "6. Sumber";

	private static int $COL_ID = 1;

	private static int $COL_REKENING = 2;

	private static int $COL_SUMBER = 3;

	private static int $COL_JENIS = 4;

	private static int $COL_NO = 5;

	private static int $COL_URAIAN = 6;

	private static int $COL_SEBELUM = 7;

	private static int $COL_SETELAH = 8;

	private static int $COL_SELISIH = 9;

	private static int $COL_KETERANGAN = 10;

	private static function readJSON(string $fileName): array
	{
		$resource = fopen($fileName, "r");
		$contents = fread($resource, filesize($fileName));
		fclose($resource);
		return json_decode($contents);
	}

	private static array $HEADER = [
		"ID" => 10,
		"REKENING" => 10,
		"SUMBER" => 20,
		"JENIS" => 10,
		"NO" => 5,
		"URAIAN" => 60,
		"APBD 2025" => 20,
		"PPAS 2026" => 20,
		"BERTAMBAH\n(BERKURANG)" => 20,
		"KETERANGAN" => 20,
	];

	private array $listSKPD;

	private array $listUnit;

	private array $listBidang;

	private function __construct()
	{
		$this->listSKPD = self::readJSON(__DIR__ . "/../../.test/01. SKPD.json");
		$this->listUnit = self::readJSON(__DIR__ . "/../../.test/02. Unit.json");
		$this->listBidang = self::readJSON(__DIR__ . "/../../.test/03. Bidang.json");
	}

	private function styleOfHeader(Worksheet $worksheet, int $rowNum): void
	{
		for ($colNum = 1; $colNum <= sizeof(self::$HEADER); $colNum++) {
			$style = $worksheet->getStyle([$colNum, $rowNum]);
			$style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
			$style->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
			$style->getAlignment()->setWrapText(true);
			$style->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
			$style->getFill()->setFillType(Fill::FILL_SOLID);
			$style->getFill()->setStartColor(new Color("6d28d9"));
			$style->getFont()->setBold(true);
			$style->getFont()->setColor(new Color(Color::COLOR_WHITE));
		}
	}

	private function styleOfSurplusDefisit(Worksheet $worksheet, int $rowNum): void
	{
		for ($colNum = 1; $colNum <= sizeof(self::$HEADER); $colNum++) {
			$style = $worksheet->getStyle([$colNum, $rowNum]);
			if ($colNum == 6) $style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
			if (in_array($colNum, [self::$COL_NO, self::$COL_KETERANGAN])) $style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
			if (in_array($colNum, [self::$COL_URAIAN, self::$COL_KETERANGAN])) $style->getAlignment()->setWrapText(true);
			$style->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
			if ($colNum >= self::$COL_NO) {
				$style->getFill()->setFillType(Fill::FILL_SOLID);
				$style->getFill()->setStartColor(new Color("a7f3d0"));
				$style->getFont()->setBold(true);
			}
			if (in_array($colNum, [self::$COL_SEBELUM, self::$COL_SETELAH, self::$COL_SELISIH])) $style->getNumberFormat()->setFormatCode("#,##0.00_);[Red](#,##0.00)");
		}
	}

	private function styleOfTotalAPBD(Worksheet $worksheet, int $rowNum): void
	{
		for ($colNum = 1; $colNum <= sizeof(self::$HEADER); $colNum++) {
			$style = $worksheet->getStyle([$colNum, $rowNum]);
			if ($colNum == 6) $style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
			if (in_array($colNum, [self::$COL_NO, self::$COL_KETERANGAN])) $style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
			if (in_array($colNum, [self::$COL_URAIAN, self::$COL_KETERANGAN])) $style->getAlignment()->setWrapText(true);
			$style->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
			if ($colNum >= self::$COL_NO) {
				$style->getFill()->setFillType(Fill::FILL_SOLID);
				$style->getFill()->setStartColor(new Color("ffff00"));
				$style->getFont()->setBold(true);
			}
			if (in_array($colNum, [self::$COL_SEBELUM, self::$COL_SETELAH, self::$COL_SELISIH])) $style->getNumberFormat()->setFormatCode("#,##0.00_);[Red](#,##0.00)");
		}
	}

	private function styleOfSKPD(Worksheet $worksheet, int $rowNum): void
	{
		for ($colNum = 1; $colNum <= sizeof(self::$HEADER); $colNum++) {
			$style = $worksheet->getStyle([$colNum, $rowNum]);
			if (in_array($colNum, [self::$COL_NO, self::$COL_KETERANGAN])) $style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
			if (in_array($colNum, [self::$COL_URAIAN, self::$COL_KETERANGAN])) $style->getAlignment()->setWrapText(true);
			$style->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
			if ($colNum >= self::$COL_NO) {
				$style->getFill()->setFillType(Fill::FILL_SOLID);
				$style->getFill()->setStartColor(new Color("5eead4"));
				$style->getFont()->setBold(true);
			}
			if (in_array($colNum, [self::$COL_SEBELUM, self::$COL_SETELAH, self::$COL_SELISIH])) $style->getNumberFormat()->setFormatCode("#,##0.00_);[Red](#,##0.00)");
		}
	}

	private function styleOfUnit(Worksheet $worksheet, int $rowNum): void
	{
		for ($colNum = 1; $colNum <= sizeof(self::$HEADER); $colNum++) {
			$style = $worksheet->getStyle([$colNum, $rowNum]);
			if (in_array($colNum, [self::$COL_NO, self::$COL_KETERANGAN])) $style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
			if (in_array($colNum, [self::$COL_URAIAN, self::$COL_KETERANGAN])) $style->getAlignment()->setWrapText(true);
			$style->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
			if ($colNum >= self::$COL_URAIAN) {
				$style->getFill()->setFillType(Fill::FILL_SOLID);
				$style->getFill()->setStartColor(new Color("ffff00"));
				$style->getFont()->setBold(true);
			}
			if (in_array($colNum, [self::$COL_SEBELUM, self::$COL_SETELAH, self::$COL_SELISIH])) $style->getNumberFormat()->setFormatCode("#,##0.00_);[Red](#,##0.00)");
		}
	}

	private function styleOfBidang(Worksheet $worksheet, int $rowNum): void
	{
		for ($colNum = 1; $colNum <= sizeof(self::$HEADER); $colNum++) {
			$style = $worksheet->getStyle([$colNum, $rowNum]);
			if (in_array($colNum, [self::$COL_NO, self::$COL_KETERANGAN])) $style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
			if (in_array($colNum, [self::$COL_URAIAN, self::$COL_KETERANGAN])) $style->getAlignment()->setWrapText(true);
			$style->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
			if ($colNum >= self::$COL_URAIAN) {
				$style->getFill()->setFillType(Fill::FILL_SOLID);
				$style->getFill()->setStartColor(new Color("f0abfc"));
				$style->getFont()->setBold(true);
			}
			if (in_array($colNum, [self::$COL_SEBELUM, self::$COL_SETELAH, self::$COL_SELISIH])) $style->getNumberFormat()->setFormatCode("#,##0.00_);[Red](#,##0.00)");
		}
	}

	private function styleOfKategori(Worksheet $worksheet, int $rowNum): void
	{
		for ($colNum = 1; $colNum <= sizeof(self::$HEADER); $colNum++) {
			$style = $worksheet->getStyle([$colNum, $rowNum]);
			if (in_array($colNum, [self::$COL_NO, self::$COL_KETERANGAN])) $style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
			if (in_array($colNum, [self::$COL_URAIAN, self::$COL_KETERANGAN])) $style->getAlignment()->setWrapText(true);
			$style->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
			if ($colNum >= self::$COL_URAIAN) {
				$style->getFill()->setFillType(Fill::FILL_SOLID);
				$style->getFill()->setStartColor(new Color("ffc000"));
				$style->getFont()->setBold(true);
			}
			if (in_array($colNum, [self::$COL_SEBELUM, self::$COL_SETELAH, self::$COL_SELISIH])) $style->getNumberFormat()->setFormatCode("#,##0.00_);[Red](#,##0.00)");
		}
	}

	private function styleOfUraian(Worksheet $worksheet, int $rowNum): void
	{
		for ($colNum = 1; $colNum <= sizeof(self::$HEADER); $colNum++) {
			$style = $worksheet->getStyle([$colNum, $rowNum]);
			if (in_array($colNum, [self::$COL_NO, self::$COL_KETERANGAN])) $style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
			if (in_array($colNum, [self::$COL_URAIAN, self::$COL_KETERANGAN])) $style->getAlignment()->setWrapText(true);
			$style->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
			if (in_array($colNum, [self::$COL_SEBELUM, self::$COL_SETELAH, self::$COL_SELISIH])) $style->getNumberFormat()->setFormatCode("#,##0.00_);[Red](#,##0.00)");
		}
	}

	private function styleOfSumber(Worksheet $worksheet, int $rowNum): void
	{
		for ($colNum = 1; $colNum <= sizeof(self::$HEADER); $colNum++) {
			$style = $worksheet->getStyle([$colNum, $rowNum]);
			if (in_array($colNum, [self::$COL_NO, self::$COL_KETERANGAN])) $style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
			if (in_array($colNum, [self::$COL_URAIAN, self::$COL_KETERANGAN])) $style->getAlignment()->setWrapText(true);
			$style->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
			if (in_array($colNum, [self::$COL_SEBELUM, self::$COL_SETELAH, self::$COL_SELISIH])) $style->getNumberFormat()->setFormatCode("#,##0.00_);[Red](#,##0.00)");
			if ($colNum >= self::$COL_URAIAN) $style->getFont()->setColor(new Color("00b050"));
			if ($colNum == self::$COL_URAIAN) {
				$style->getAlignment()->setIndent(2);
				$style->getFont()->setItalic(true);
			}
		}
	}

	private function formulaOfSelisih(int $rowNum): string
	{
		$col1 = chr(64 + self::$COL_SETELAH);
		$col2 = chr(64 + self::$COL_SEBELUM);
		return "=$col1$rowNum-$col2$rowNum";
	}

	private function formulaOfSurplusDefisit(int $colNum, int $rowNum): string
	{
		$ref = chr(64 + self::$COL_JENIS);
		$col = chr(64 + $colNum);
		$row = $rowNum + 1;
		$key = self::$KEY_SKPD;
		return "=$col$row-SUMIFS($col:$col,\$$ref:\$$ref,\"$key\")";
	}

	private function formulaOfSum(int $colNum, int $row1, int $row2): string
	{
		$col = chr(64 + $colNum);
		return "=SUM($col$row1:$col$row2)";
	}

	private function formulaOfSumifs(int $colNum, int $row1, int $row2, string $key): string
	{
		$ref = chr(64 + self::$COL_JENIS);
		$col = chr(64 + $colNum);
		return "=SUMIFS($col$row1:$col$row2,\$$ref$row1:\$$ref$row2,\"$key\")";
	}

	public function create(string $fileName): void
	{
		$spreadsheet = new Spreadsheet();
		$defaultStyle = $spreadsheet->getDefaultStyle();
		$defaultStyle->getFont()->setName("Aptos Narrow");
		$defaultStyle->getFont()->setSize(11);
		$defaultStyle->getAlignment()->setVertical(Alignment::VERTICAL_TOP);

		$worksheet = $spreadsheet->getActiveSheet();
		$worksheet->getRowDimension(1)->setRowHeight(30);

		$rowNum = 1;
		$colNum = 1;
		foreach (self::$HEADER as $title => $width) {
			$worksheet->getColumnDimension(chr(64 + $colNum))->setWidth($width);
			$worksheet->getCell([$colNum, $rowNum])->setValueExplicit($title, DataType::TYPE_STRING);
			if ($colNum >= 5) $worksheet->getCell([$colNum, $rowNum + 1])->setValueExplicit($colNum - 4, DataType::TYPE_NUMERIC);
			$colNum++;
		}

		$this->styleOfHeader($worksheet, 1);
		$this->styleOfHeader($worksheet, 2);
		$this->styleOfUraian($worksheet, 3);

		$rowNum = 4;
		$this->styleOfSurplusDefisit($worksheet, $rowNum);
		$worksheet->getCell([self::$COL_URAIAN, $rowNum])->setValueExplicit("SURPLUS/(DEFISIT)", DataType::TYPE_STRING);
		$worksheet->getCell([self::$COL_SEBELUM, $rowNum])->setValueExplicit($this->formulaOfSurplusDefisit(self::$COL_SEBELUM, $rowNum), DataType::TYPE_FORMULA);
		$worksheet->getCell([self::$COL_SETELAH, $rowNum])->setValueExplicit($this->formulaOfSurplusDefisit(self::$COL_SETELAH, $rowNum), DataType::TYPE_FORMULA);
		$worksheet->getCell([self::$COL_SELISIH, $rowNum])->setValueExplicit($this->formulaOfSelisih($rowNum), DataType::TYPE_FORMULA);

		$rowNum = 5;
		$this->styleOfTotalAPBD($worksheet, $rowNum);
		$worksheet->getCell([self::$COL_URAIAN, $rowNum])->setValueExplicit("APBD KOTA MATARAM", DataType::TYPE_STRING);
		$worksheet->getCell([self::$COL_SEBELUM, $rowNum])->setValueExplicit(0, DataType::TYPE_NUMERIC);
		$worksheet->getCell([self::$COL_SETELAH, $rowNum])->setValueExplicit(0, DataType::TYPE_NUMERIC);
		$worksheet->getCell([self::$COL_SELISIH, $rowNum])->setValueExplicit($this->formulaOfSelisih($rowNum), DataType::TYPE_FORMULA);

		$rowNum++;
		$this->styleOfUraian($worksheet, $rowNum);

		$rowNum++;
		for ($indexOfSKPD = 0; $indexOfSKPD < sizeof($this->listSKPD); $indexOfSKPD++) {
			$skpdRow = $rowNum;
			$skpdNum = $indexOfSKPD + 1;
			$skpdCode = $this->listSKPD[$indexOfSKPD]->kode;
			$skpdName = $this->listSKPD[$indexOfSKPD]->nama;

			$worksheet->getCell([self::$COL_JENIS, $rowNum])->setValueExplicit(self::$KEY_SKPD, DataType::TYPE_STRING);
			$worksheet->getCell([self::$COL_NO, $rowNum])->setValueExplicit($skpdNum, DataType::TYPE_NUMERIC);
			$worksheet->getCell([self::$COL_URAIAN, $rowNum])->setValueExplicit($skpdName, DataType::TYPE_STRING);
			$worksheet->getCell([self::$COL_SELISIH, $rowNum])->setValueExplicit($this->formulaOfSelisih($rowNum), DataType::TYPE_FORMULA);
			$this->styleOfSKPD($worksheet, $rowNum);

			$bidangCodes = [];
			$bidangCodes[] = substr($skpdCode, 0, 4);
			$bidangCodes[] = substr($skpdCode, 5, 4);
			$bidangCodes[] = substr($skpdCode, 10, 4);
			$bidangCodes = array_values(array_filter($bidangCodes, fn($code) => $code != "0.00"));

			$listBidang = array_values(array_filter($this->listBidang, fn($bidang) => in_array($bidang->kode, $bidangCodes)));
			foreach ($listBidang as $bidang) {
				$rowNum++;
				$bidangRow = $rowNum;
				$bidangName = $bidang->nama;
				$worksheet->getCell([self::$COL_JENIS, $rowNum])->setValueExplicit(self::$KEY_BIDANG, DataType::TYPE_STRING);
				$worksheet->getCell([self::$COL_URAIAN, $rowNum])->setValueExplicit($bidangName, DataType::TYPE_STRING);
				$worksheet->getCell([self::$COL_SELISIH, $rowNum])->setValueExplicit($this->formulaOfSelisih($rowNum), DataType::TYPE_FORMULA);
				$this->styleOfBidang($worksheet, $rowNum);

				$listUnit = array_values(array_filter($this->listUnit, fn($unit) => $unit->skpd == $skpdCode));
				foreach ($listUnit as $unit) {
					$rowNum++;
					$unitRow = $rowNum;
					$unitName = $unit->nama;
					$worksheet->getCell([self::$COL_JENIS, $rowNum])->setValueExplicit(self::$KEY_UNIT, DataType::TYPE_STRING);
					$worksheet->getCell([self::$COL_URAIAN, $rowNum])->setValueExplicit($unitName, DataType::TYPE_STRING);
					$worksheet->getCell([self::$COL_SELISIH, $rowNum])->setValueExplicit($this->formulaOfSelisih($rowNum), DataType::TYPE_FORMULA);
					$this->styleOfUnit($worksheet, $rowNum);

					foreach (self::$KATEGORI as $kategori) {
						$rowNum++;
						$kategoriRow = $rowNum;
						$worksheet->getCell([self::$COL_JENIS, $rowNum])->setValueExplicit(self::$KEY_KATEGORI, DataType::TYPE_STRING);
						$worksheet->getCell([self::$COL_URAIAN, $rowNum])->setValueExplicit($kategori, DataType::TYPE_STRING);
						$worksheet->getCell([self::$COL_SELISIH, $rowNum])->setValueExplicit($this->formulaOfSelisih($rowNum), DataType::TYPE_FORMULA);
						$this->styleOfKategori($worksheet, $rowNum);

						for ($i = 0; $i < 2; $i++) {
							$rowNum++;
							$uraianRow = $rowNum;
							$worksheet->getCell([self::$COL_JENIS, $rowNum])->setValueExplicit(self::$KEY_URAIAN, DataType::TYPE_STRING);
							$worksheet->getCell([self::$COL_URAIAN, $rowNum])->setValueExplicit("Uraian JsonToExcel", DataType::TYPE_STRING);
							$worksheet->getCell([self::$COL_SELISIH, $rowNum])->setValueExplicit($this->formulaOfSelisih($rowNum), DataType::TYPE_FORMULA);
							$this->styleOfUraian($worksheet, $rowNum);

							for ($i = 0; $i < 2; $i++) {
								$rowNum++;
								$worksheet->getCell([self::$COL_JENIS, $rowNum])->setValueExplicit(self::$KEY_SUMBER, DataType::TYPE_STRING);
								$worksheet->getCell([self::$COL_URAIAN, $rowNum])->setValueExplicit("Sumber Dana", DataType::TYPE_STRING);
								$worksheet->getCell([self::$COL_SEBELUM, $rowNum])->setValueExplicit(0, DataType::TYPE_NUMERIC);
								$worksheet->getCell([self::$COL_SETELAH, $rowNum])->setValueExplicit(0, DataType::TYPE_NUMERIC);
								$worksheet->getCell([self::$COL_SELISIH, $rowNum])->setValueExplicit($this->formulaOfSelisih($rowNum), DataType::TYPE_FORMULA);
								$this->styleOfSumber($worksheet, $rowNum);
							}

							$worksheet->getCell([self::$COL_SEBELUM, $uraianRow])->setValueExplicit($this->formulaOfSum(self::$COL_SEBELUM, $uraianRow + 1, $rowNum), DataType::TYPE_FORMULA);
							$worksheet->getCell([self::$COL_SETELAH, $uraianRow])->setValueExplicit($this->formulaOfSum(self::$COL_SETELAH, $uraianRow + 1, $rowNum), DataType::TYPE_FORMULA);
						}

						$rowNum++;
						$this->styleOfUraian($worksheet, $rowNum);

						$worksheet->getCell([self::$COL_SEBELUM, $kategoriRow])->setValueExplicit($this->formulaOfSumifs(self::$COL_SEBELUM, $kategoriRow + 1, $rowNum, self::$KEY_URAIAN), DataType::TYPE_FORMULA);
						$worksheet->getCell([self::$COL_SETELAH, $kategoriRow])->setValueExplicit($this->formulaOfSumifs(self::$COL_SETELAH, $kategoriRow + 1, $rowNum, self::$KEY_URAIAN), DataType::TYPE_FORMULA);
					}

					$worksheet->getCell([self::$COL_SEBELUM, $unitRow])->setValueExplicit($this->formulaOfSumifs(self::$COL_SEBELUM, $unitRow + 1, $rowNum, self::$KEY_KATEGORI), DataType::TYPE_FORMULA);
					$worksheet->getCell([self::$COL_SETELAH, $unitRow])->setValueExplicit($this->formulaOfSumifs(self::$COL_SETELAH, $unitRow + 1, $rowNum, self::$KEY_KATEGORI), DataType::TYPE_FORMULA);
				}

				$worksheet->getCell([self::$COL_SEBELUM, $bidangRow])->setValueExplicit($this->formulaOfSumifs(self::$COL_SEBELUM, $bidangRow + 1, $rowNum, self::$KEY_UNIT), DataType::TYPE_FORMULA);
				$worksheet->getCell([self::$COL_SETELAH, $bidangRow])->setValueExplicit($this->formulaOfSumifs(self::$COL_SETELAH, $bidangRow + 1, $rowNum, self::$KEY_UNIT), DataType::TYPE_FORMULA);
			}

			$worksheet->getCell([self::$COL_SEBELUM, $skpdRow])->setValueExplicit($this->formulaOfSumifs(self::$COL_SEBELUM, $skpdRow + 1, $rowNum, self::$KEY_BIDANG), DataType::TYPE_FORMULA);
			$worksheet->getCell([self::$COL_SETELAH, $skpdRow])->setValueExplicit($this->formulaOfSumifs(self::$COL_SETELAH, $skpdRow + 1, $rowNum, self::$KEY_BIDANG), DataType::TYPE_FORMULA);
		}

		$rowNum++;
		$this->styleOfUraian($worksheet, $rowNum);

		$xlsx = new XlsxWriter($spreadsheet);
		$xlsx->save($fileName);
	}
}