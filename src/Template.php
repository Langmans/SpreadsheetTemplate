<?php

namespace Langmans\SpreadsheetTemplate;

use JsonPath\JsonObject;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use Psr\SimpleCache\CacheInterface;

class Template
{
    protected Spreadsheet $spreadsheet;
    protected Parser $parser;

    protected array $formatters = [];

    protected array $cellData;

    public function __construct(protected CacheInterface $cache, Parser $parser = null)
    {
        $this->parser = $parser ?? new Parser();
    }

    public function setFormatter(string $name, callable $formatter): static
    {
        $this->formatters[$name] = $formatter;
        return $this;
    }

    public function load(string $path, array|object $data): Spreadsheet
    {
        $this->parseCells($path);
        $this->setData($data);
        return $this->spreadsheet;
    }

    protected function parseCells(string $path): void
    {
        $this->spreadsheet = IOFactory::load($path);

        // we compare the modification time to the one in the cache, so we can skip parsing if the file hasn't changed.
        $modified = filemtime($path);
        $key = md5(realpath($path));
        if ($this->cache->has($key)) {
            $cached = $this->cache->get($key);
            if ($modified === $cached['modified']) {
                $this->cellData = $cached['cellData'];
                return;
            }
        }

        $this->cellData = $cellData = $this->parser->getCells($this->spreadsheet);
        $this->cache->set($key, compact('cellData', 'modified'));
    }

    protected function setData(array|object $data): void
    {
        $json = new JsonObject($data);
        foreach ($this->cellData as $cellData) {
            $value = $cellData['value'];
            $cell = $cellData['cell'];
            $row = $cellData['row'];

            // jsonpath always returns an array, even if there's only one value. So just loop through the items.
            // when we encounter different lengths for each find, then we should loop over the longest one.
            // we should always have at least one row, so we can find the tags.
            $allReplacements = [];
            $rows = 1;
            foreach ($cellData['replacements'] as $find => [$path, $settings]) {
                $replacements = $json->get($path);
                $rows = max($rows, count($replacements));
                $allReplacements[$find] = [$replacements, $settings];
            }

            for ($i = 0; $i < $rows; $i++) {
                $replacements = [];
                foreach ($allReplacements as $find => [$replacementValues, $settings]) {
                    $replacement = $replacementValues[$i] ?? null;
                    $formatter = $this->formatters[$settings['formatter'] ?? null] ?? null;
                    if (is_callable($formatter)) {
                        $replacement = $formatter($replacement, $i, $settings);
                    }

                    if (!($replacement === null || $replacement === '')) {
                        // example: @{$.foo[*].bar}{"prefixWhenValue":"€ "}
                        if (isset($settings['prefixWhenValue'])) {
                            $replacement = $settings['prefixWhenValue'] . $replacement;
                        }
                        // example: @{$.foo[*].bar}{"suffixWhenValue":"%"}
                        if (isset($settings['suffixWhenValue'])) {
                            $replacement = $replacement . $settings['suffixWhenValue'];
                        }
                    }

                    // these will add a prefix or suffix to the replacement,
                    // regardless of whether the replacement is empty or not.
                    if (isset($settings['prefix'])) {
                        $replacement = $settings['prefix'] . $replacement;
                    }
                    if (isset($settings['suffix'])) {
                        $replacement = $replacement . $settings['suffix'];
                    }

                    $replacements[$find] = $replacement;
                }
                $newValue = str_replace(array_keys($replacements), array_values($replacements), $value);
                $this->spreadsheet
                    ->getSheet($cellData['sheet'])
                    ->setCellValue($cell . ($row + $i), $newValue);
            }
        }
    }
}
