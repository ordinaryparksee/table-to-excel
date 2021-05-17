<?php

namespace Vaxy\TableToExcel;

use DOMDocument;
use DOMElement;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

class TableParser {

    public static $path;

    public static function toExcelWidth($width)
    {
        if (preg_match('/px$/', $width) > 0) {
            return self::pixelToPoint(preg_replace('/px$/', '', $width)) / 10 * 1.174;
        }
        else if (preg_match('/pt$/', $width) > 0) {
            return preg_replace('/pt$/', '', $width) / 10 * 1.174;
        }
        else {
            return $width;
        }
    }

    public static function toExcelHeight($height)
    {
        if (preg_match('/px$/', $height) > 0) {
            return self::pixelToPoint(preg_replace('/px$/', '', $height));
        }
        else if (preg_match('/pt$/', $height) > 0) {
            return preg_replace('/pt$/', '', $height);
        }
        else {
            return $height;
        }
    }

    public static function makeTableLayout(DOMElement $table)
    {
        $tableRange = [[1, 1], [1, 1]];

        $tableCss = CssParser::parse($table->getAttribute('style'));

        if ($tableCss->has('margin-top')) {
            $tableRange[0][1]++;
        }
        if ($tableCss->has('margin-left')) {
            $tableRange[0][0]++;
        }

        $rowIndex = $tableRange[0][1] - 1;
        $rowspans = [];
        foreach ($table->getElementsByTagName('tr') as $tr) {
            $rowIndex++;
            if ($tableRange[1][1] < $rowIndex) {
                $tableRange[1][1] = $rowIndex;
            }
            $rowspanStep = 0;
            $columnIndex = $tableRange[0][0] - 1;
            foreach ($tr->childNodes as $td) {
                if ($td->nodeName === 'th' || $td->nodeName === 'td') {
                    $columnIndex++;
                    if (array_key_exists($columnIndex + $rowspanStep, $rowspans)) {
                        foreach ($rowspans[$columnIndex + $rowspanStep] as $rows) {
                            if (in_array($rowIndex, $rows)) {
                                $rowspanStep++;
                                break;
                            }
                        }
                    }

                    // Merge
                    if ($td->hasAttribute('rowspan')) {
                        $rowspan = $td->getAttribute('rowspan') - 1;
                        if (array_key_exists($columnIndex, $rowspans) === false) {
                            $rowspans[$columnIndex] = [];
                        }
                        $rowspans[$columnIndex][] = range($rowIndex, $rowIndex + $rowspan);
                    }
                    if ($td->hasAttribute('colspan')) {
                        $colspan = $td->getAttribute('colspan') - 1;
                        if ($colspan > 0) {
                            $columnIndex += $colspan;
                        }
                    }

                    if ($tableRange[1][0] < $columnIndex + $rowspanStep) {
                        $tableRange[1][0] = $columnIndex + $rowspanStep;
                    }
                }
            }
        }
        return $tableRange;
    }

    public static function parse($source)
    {
        $dom = new DOMDocument();
        $dom->loadHTML(mb_convert_encoding(str_replace('&', '&amp;', $source), 'HTML-ENTITIES', 'UTF-8'));

        $spreadsheet = new Spreadsheet();

        foreach ($dom->getElementsByTagName('table') as $tableIndex => $table) {
            $layout = self::makeTableLayout($table);

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

            $tableRange = [[1, 1], [1, 1]];

            $tableCss = CssParser::parse($table->getAttribute('style'));

            $cssExtend = [];
            if ($tableCss->has('font-family')) {
                $cssExtend['font-family'] = $tableCss['font-family'];
            }
            if ($tableCss->has('font-size')) {
                $cssExtend['font-size'] = $tableCss['font-size'];
            }
            if ($tableCss->has('text-align')) {
                $cssExtend['text-align'] = $tableCss['text-align'];
            }
            if ($tableCss->has('vertical-align')) {
                $cssExtend['vertical-align'] = $tableCss['vertical-align'];
            }
            if ($tableCss->has('color')) {
                $cssExtend['color'] = $tableCss['color'];
            }
            if ($tableCss->has('background')) {
                $cssExtend['background'] = $tableCss['background'];
            }
            if ($tableCss->has('background-color')) {
                $cssExtend['background-color'] = $tableCss['background-color'];
            }

            if ($tableCss->has('margin-top')) {
                $tableRange[0][1]++;
                $marginTop = self::toExcelHeight($tableCss['margin-top']);
                $sheet->getRowDimension(1)->setRowHeight($marginTop);
            }
            if ($tableCss->has('margin-left')) {
                $tableRange[0][0]++;
                $marginLeft = self::toExcelWidth($tableCss['margin-left']);
                $sheet->getColumnDimensionByColumn(1)->setWidth($marginLeft);
            }

            foreach ($table->getElementsByTagName('col') as $colIndex => $col) {
                if ($col->hasAttribute('width')) {
                    if ($tableCss->has('margin-left')) {
                        $sheet->getColumnDimensionByColumn($colIndex + 2)->setWidth($col->getAttribute('width'));
                    } else {
                        $sheet->getColumnDimensionByColumn($colIndex + 1)->setWidth($col->getAttribute('width'));
                    }
                }
            }
    
            $rowIndex = $tableRange[0][1] - 1;
            $rowspans = [];
            foreach ($table->getElementsByTagName('tr') as $tr) {
                $rowIndex++;
                if ($tableRange[1][1] < $rowIndex) {
                    $tableRange[1][1] = $rowIndex;
                }
                $rowspanStep = 0;
                $_cssExtend = $cssExtend;
                if ($tr->hasAttribute('height')) {
                    $sheet->getRowDimension($rowIndex)->setRowHeight($tr->getAttribute('height'));
                }
                if ($tr->hasAttribute('style')) {
                    $rowCss = CssParser::parse($tr->getAttribute('style'), $cssExtend);
                    if ($rowCss->has('font-family')) {
                        $_cssExtend['font-family'] = $rowCss['font-family'];
                    }
                    if ($rowCss->has('font-size')) {
                        $_cssExtend['font-size'] = $rowCss['font-size'];
                    }
                    if ($rowCss->has('text-align')) {
                        $_cssExtend['text-align'] = $rowCss['text-align'];
                    }
                    if ($rowCss->has('vertical-align')) {
                        $_cssExtend['vertical-align'] = $rowCss['vertical-align'];
                    }
                    if ($rowCss->has('color')) {
                        $_cssExtend['color'] = $rowCss['color'];
                    }
                    if ($rowCss->has('background')) {
                        $_cssExtend['background'] = $rowCss['background'];
                    }
                    if ($rowCss->has('background-color')) {
                        $_cssExtend['background-color'] = $rowCss['background-color'];
                    }
                } else {
                    $rowCss = null;
                }
                $columnIndex = $tableRange[0][0] - 1;
                foreach ($tr->childNodes as $td) {
                    if ($td->nodeName === 'th' || $td->nodeName === 'td') {
                        $columnIndex++;
                        if (array_key_exists($columnIndex + $rowspanStep, $rowspans)) {
                            foreach ($rowspans[$columnIndex + $rowspanStep] as $rows) {
                                if (in_array($rowIndex, $rows)) {
                                    $rowspanStep++;
                                    break;
                                }
                            }
                        }
                        $cell = $sheet->getCellByColumnAndRow($columnIndex + $rowspanStep, $rowIndex);

                        if ($td->hasAttribute('explicit')) {
                            $explicit = $td->getAttribute('explicit');
                            $cell->setValueExplicit(preg_replace_callback('/\{\{([^}]+)\}\}/', function($_) use ($cell) {
                                return eval('return '.$_[1].';');
                            }, $td->textContent), $explicit);
                        } else {
                            $cell->setValue(preg_replace_callback('/\{\{([^}]+)\}\}/', function($_) use ($cell) {
                                return eval('return '.$_[1].';');
                            }, $td->textContent));
                        }

                        $style = $cell->getStyle();
                        $style->getAlignment()->setVertical('center');
                        $pre = $td->getElementsByTagName('pre');
                        if ($pre && $pre->length > 0) {
                            $style->getAlignment()->setWrapText(true);
                        }
                        $font = $style->getFont();
                        if ($td->nodeName === 'th') {
                            $font->setBold(true);
                            $style->getAlignment()->setVertical('center')->setHorizontal('center');
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
                        $css = CssParser::parse($td->getAttribute('style'), $_cssExtend);
                        if ($td->hasAttribute('width')) {
                            $css['width'] = $td->getAttribute('width');
                        }
                        if ($td->hasAttribute('height')) {
                            $css['height'] = $td->getAttribute('height');
                        }
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

                        if ($tableRange[1][0] < $columnIndex + $rowspanStep) {
                            $tableRange[1][0] = $columnIndex + $rowspanStep;
                        }
                    }
                }

                if ($rowCss) {
                    $rowStyle = $sheet->getStyleBycolumnAndRow($tableRange[0][0], $rowIndex, $layout[1][0], $rowIndex);
                    self::applyBorder($rowStyle, $rowCss);
                }
            }

            if ($table->hasAttribute('style')) {
                if ($tableCss->has('border') || $tableCss->has('border-style') || $tableCss->has('border-color') || $tableCss->has('border-width')) {
                    $tableStyle = $sheet->getStyleBycolumnAndRow($tableRange[0][0], $tableRange[0][1], $tableRange[1][0], $tableRange[1][1]);
                    self::applyBorder($tableStyle, $tableCss);
                }
            }

            $sheet->setSelectedCell('A1');
        }

        $spreadsheet->setActiveSheetIndex(0);

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
            $height = self::toExcelHeight($css['height']);
            $sheet->getRowDimension($cell->getRow())->setRowHeight($height);
        }

        // Width
        if ($css->has('width')) {
            $width = self::toExcelHeight($css['width']);
            $sheet->getColumnDimension($cell->getColumn())->setWidth($width);
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
            else if (preg_match('/pt$/', $css['font-size']) > 0) {
                $font->setSize(preg_replace('/pt$/', '', $css['font-size']));
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
        $border = [];
        $borderStyle = null;
        $borderColor = null;
        $borderWidth = null;

        if ($css->has('border')) {
            $border = explode(' ', $css['border']);
            $target = $style->getBorders()->getOutline();
        }
        else if ($css->has('border-left')) {
            $border = explode(' ', $css['border-left']);
            $target = $style->getBorders()->getLeft();
        }
        else if ($css->has('border-right')) {
            $border = explode(' ', $css['border-right']);
            $target = $style->getBorders()->getRight();
        }
        else if ($css->has('border-top')) {
            $border = explode(' ', $css['border-top']);
            $target = $style->getBorders()->getTop();
        }
        else if ($css->has('border-bottom')) {
            $border = explode(' ', $css['border-bottom']);
            $target = $style->getBorders()->getBottom();
        }

        if (isset($target)) {
            while (count($border) > 0) {
                $attribute = array_shift($border);
                // Border width
                if (preg_match('/px$/', $attribute) > 0) {
                    $borderWidth = (int)preg_replace('/px$/', '', $attribute);
                }
                // Border color
                else if (preg_match('/^#/', $attribute) > 0) {
                    $borderColor = preg_replace('/^#/', '', $attribute);
                }
                // Border style
                else {
                    $borderStyle = $attribute;
                }
            }
    
            if ($css->has('border-style')) {
                $borderStyle = $css['border-style'];
            }
            if ($css->has('border-color')) {
                $borderColor = $css['border-color'];
            }
            if ($css->has('border-width')) {
                $borderWidth = $css['border-width'];
            }
    
            if ($borderWidth === 2) {
                $target->setBorderStyle('medium');
            }
            else if ($borderWidth > 2) {
                $target->setBorderStyle('thick');
            }
            else if ($borderStyle === 'solid') {
                $target->setBorderStyle('thin');
            }
            else {
                if ($borderStyle) {
                    $target->setBorderStyle($borderStyle);
                }
            }
        }
    }

}
