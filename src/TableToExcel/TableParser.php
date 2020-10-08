<?php

namespace Vaxy\TableToExcel;

use DOMDocument;
use DOMElement;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

class TableParser {

    public static $path;

    public static function parse($source)
    {
        $dom = new DOMDocument();
        $dom->loadHTML(mb_convert_encoding($source, 'HTML-ENTITIES', 'UTF-8'));

        $spreadsheet = new Spreadsheet();

        foreach ($dom->getElementsByTagName('table') as $tableIndex => $table) {
            $caption = $table->getElementsByTagName('caption')->item(0);
            if ($tableIndex < 1) {
                $sheet = $spreadsheet->getActiveSheet();
            } else {
                $sheet = $spreadsheet->createSheet();
            }

            if ($caption instanceof DOMElement && $caption->nodeValue) {
                $sheet->setTitle($caption->nodeValue);
            } else {
                $sheet->setTitle('Table'.$tableIndex);
            }

            $rowIndex = 0;
            $rowspans = [];
            foreach ($table->getElementsByTagName('tr') as $tr) {
                $rowIndex++;
                $columnIndex = 0;
                $rowspanStep = 0;
                foreach ($tr->childNodes as $td) {
                    if ($td->nodeName === 'th' || $td->nodeName === 'td') {
                        $columnIndex++;
                        // var_dump($rowIndex, $columnIndex, $rowspans);
                        if (array_key_exists($columnIndex + $rowspanStep, $rowspans)) {
                            foreach ($rowspans[$columnIndex + $rowspanStep] as $rows) {
                                if (in_array($rowIndex, $rows)) {
                                    $rowspanStep++;
                                    break;
                                }
                            }
                        }
                        $cell = $sheet->getCellByColumnAndRow($columnIndex + $rowspanStep, $rowIndex);
                        $cell->setValue($td->nodeValue);
                        $style = $cell->getStyle();
                        $font = $style->getFont();
                        if ($td->nodeName === 'th') {
                            $font->setBold(true);
                            $style->getAlignment()->setVertical('center');
                        }

                        // Formatting
                        if ($td->hasAttribute('number-format')) {
                            $format = $td->getAttribute('number-format');
                            if ($format) {
                                $style->getNumberFormat()->setFormatCode($format);
                            } else {
                                $style->getNumberFormat()->setFormatCode('#,##0');
                            }
                        }

                        // Cascading Style Sheet
                        $css = CssParser::parse($td->getAttribute('style'));
                        if ($css) {
                            self::applyCellCss($sheet, $cell, $css);
                        }

                        // Merge
                        if ($td->hasAttribute('rowspan')) {
                            $rowspan = $td->getAttribute('rowspan') - 1;
                            $sheet->mergeCells($cell->getColumn().$cell->getRow().':'.$cell->getColumn().($cell->getRow() + $rowspan));
                            $mergeStyle = $sheet->getStyle($cell->getColumn().$cell->getRow().':'.$cell->getColumn().($cell->getRow() + $rowspan));
                            self::applyBorder($mergeStyle, $css);
                            if (array_key_exists($columnIndex, $rowspans) === false) {
                                $rowspans[$columnIndex] = [];
                            }
                            $rowspans[$columnIndex][] = range($rowIndex, $rowIndex + $rowspan);
                        }
                        if ($td->hasAttribute('colspan')) {
                            $colspan = $td->getAttribute('colspan') - 1;
                            $sheet->mergeCells($cell->getColumn().$cell->getRow().':'.chr(ord($cell->getColumn()) + $colspan).$cell->getRow());
                            $mergeStyle = $sheet->getStyle($cell->getColumn().$cell->getRow().':'.chr(ord($cell->getColumn()) + $colspan).$cell->getRow());
                            self::applyBorder($mergeStyle, $css);
                            if ($colspan > 0) {
                                $columnIndex += $colspan;
                            }
                        }
                    }
                }
            }

            $sheet->setSelectedCell('A1');
        }

        return $spreadsheet;
    }

    protected static function exportFile(string $path, array $data)
    {
        self::$path = $path;
        ob_start();
        extract($data);
        include self::$path;
        return ob_get_clean();
    }

    public static function parseFromFile($path, $data = [])
    {
        return self::parse(self::exportFile($path, $data));
    }

    public static function pixelToPoint($pixel)
    {
        return $pixel * 0.75;
    }

    public static function applyCellCss($sheet, $cell, $css)
    {
        $style = $cell->getStyle();
        $font = $style->getFont();

        // Font weight
        if ($css->has('font-weight')) {
            // Bold
            if ($css['font-weight'] >= 700 || $css['font-weight'] === 'bold' || $css['font-weight'] === 'bolder') {
                $font->setBold(true);
            }
            // Normal
            else if ($css['font-weight'] <= 400 || $css['font-weight'] === 'normal') {
                $font->setBold(false);
            }
            // Thin
            else {
                $font->setBold(false);
            }
        }

        // Background
        $backgroundColor = null;
        if ($css->has('background')) {
            $backgroundColor = $css['background'];
        }
        if ($css->has('background-color')) {
            $backgroundColor = $css['background-color'];
        }
        if (preg_match('/^#[0-9a-fA-F]{6}/', $backgroundColor) > 0) {
            $style->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setRGB(preg_replace('/^#/', '', $backgroundColor));
        }

        // Height
        if ($css->has('height')) {
            if (preg_match('/px$/', $css['height']) > 0) {
                $height = self::pixelToPoint(preg_replace('/px$/', '', $css['height']));
            }
            if (preg_match('/pt$/', $css['height']) > 0) {
                $height = preg_replace('/pt$/', '', $css['height']);
            }
            $sheet->getRowDimension($cell->getRow())->setRowHeight($height);
        }

        // Width
        if ($css->has('width')) {
            if (preg_match('/px$/', $css['width']) > 0) {
                $width = self::pixelToPoint(preg_replace('/px$/', '', $css['width'])) / 10;
            }
            if (preg_match('/pt$/', $css['width']) > 0) {
                $width = preg_replace('/pt$/', '', $css['width']) / 10;
            }
            $sheet->getColumnDimension($cell->getColumn())->setWidth($width * 1.174);
        }

        // Alignment
        if ($css->has('text-align') || $css->has('vertical-align')) {
            $align = $style->getAlignment();
            if ($css->has('text-align')) {
                $align->setHorizontal($css['text-align']);
            }
            if ($css->has('vertical-align')) {
                if ($css['vertical-align'] === 'middle') {
                    $align->setVertical('center');
                } else {
                    $align->setVertical($css['vertical-align']);
                }
            }
        }

        // Font
        if ($css->has('color')) {
            $font->getColor()->setRGB(preg_replace('/^#/', '', $css['color']));
        }
        if ($css->has('font-family')) {
            $fontFamilies = explode(',', $css['font-family']);
            $fontName = str_replace(['"', "'"], '', array_shift($fontFamilies));
            $font->setName($fontName);
        }
        if ($css->has('font-size')) {
            if (preg_match('/px$/', $css['font-size']) > 0) {
                $font->setSize(self::pixelToPoint(preg_replace('/px$/', '', $css['font-size'])));
            }
            else {
                $font->setSize($css['font-size']);
            }
        }

        // Border
        self::applyBorder($style, $css);
    }

    public static function applyBorder($style, $css)
    {
        if ($css->has('border')) {
            $border = explode(' ', $css['border']);
            while (count($border) > 0) {
                $attribute = array_shift($border);
                // Border width
                if (preg_match('/px$/', $attribute) > 0) {
                }
                // Border color
                else if (preg_match('/^#/', $attribute) > 0) {
                }
                // Border style
                else {
                    if ($attribute === 'solid') {
                        $style->getBorders()->getAllBorders()->setBorderStyle('thin');
                    } else {
                        $style->getBorders()->getAllBorders()->setBorderStyle($attribute);
                    }
                }
            }
        }
    }

}
