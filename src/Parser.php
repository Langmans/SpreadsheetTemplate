<?php

namespace Langmans\SpreadsheetTemplate;

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
        // match some json paths e.g. @{$.foo} or @{$.foo[0][1]} or @{$.foo[*].bar}
        // settings can be added by appending json object: @{$.foo[*].bar}{"foo":"bar"}
        // see the template class for available settings.
        $regex = '/@{(?<jsonpath>[^{}]+)}(?<settings>\{[^}]+})?/';

        foreach ($spreadsheet->getWorksheetIterator() as $sheetIndex => $worksheet) {
            foreach ($worksheet->getRowIterator() as $row) {
                foreach ($row->getCellIterator() as $cell) {
                    $value = $cell->getValue();
                    if (!is_string($value) || !preg_match_all($regex, $value, $matches)) {
                        continue;
                    }
                    $cells[] = [
                        'sheet' => $sheetIndex,
                        'cell' => $cell->getColumn(),
                        'row' => $cell->getRow(),
                        'value' => $value,
                        'replacements' => $this->getReplacementsFromAllMatches($matches),
                    ];
                }
            }
        }
        return $cells;
    }

    protected function getReplacementsFromAllMatches(array $matches): array
    {
        $replacements = [];
        foreach ($matches[0] as $i => $replacement) {
            $settings = $matches['settings'][$i];
            $replacements[$replacement] = [
                $matches['jsonpath'][$i],
                str_starts_with($settings, '{') && str_ends_with($settings, '}')
                    ? json_decode($settings, true)
                    : [],
            ];
        }
        return $replacements;
    }
}
