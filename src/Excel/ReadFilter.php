<?php


namespace ExcelToCoordinates\Excel;

use PhpOffice\PhpSpreadsheet\Reader\IReadFilter;

class ReadFilter implements IReadFilter
{
    private int $startRow = 0;

    private int $endRow = 0;

    private array $columns = [];

    public function __construct($startRow, $endRow, $columns)
    {
        $this->startRow = $startRow;
        $this->endRow = $endRow;
        $this->columns = $columns;
    }

    public function readCell($column, $row, $worksheetName = ''): bool
    {
        if ($row >= $this->startRow && $row <= $this->endRow) {
            if (in_array($column, $this->columns)) {
                return true;
            }
        }

        return false;
    }
}
