<?php

namespace Langmans\PhpSpreadsheet;

use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class Parser
{
    /**
     * @throws Exception
     */
    public function getCells(Spreadsheet $spreadsheet): array
    {
        $cells = [];
        foreach ($spreadsheet->getWorksheetIterator() as $sheetIndex => $worksheet) {
            foreach ($worksheet->getRowIterator() as $row) {
                foreach ($row->getCellIterator() as $cell) {
                    $value = $cell->getValue();
                    // todo: cell styles? cell formatting? conditional prefix / suffix? how?
                    // match some json paths e.g. @{$.foo} or @{$.foo[0][1]} or @{$.foo[*].bar}
                    if (!is_string($value) || !preg_match_all('/@{([^{}]+)}/', $value, $matches)) {
                        continue;
                    }
                    $replacements = array_combine($matches[0], $matches[1]);
                    $cells[] = [
                        'sheet' => $sheetIndex,
                        'cell' => $cell->getColumn(),
                        'row' => $cell->getRow(),
                        'value' => $value,
                        'replacements' => $replacements,
                    ];
                }
            }
        }
        return $cells;
    }
}