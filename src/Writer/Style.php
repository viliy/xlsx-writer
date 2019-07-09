<?php

/**
 * Class Style.
 */

namespace Zhaqq\Xlsx\Writer;

use Zhaqq\Xlsx\Support;

trait Style
{
    /**
     * 风格设置.
     *
     * @var array
     */
    protected $cellStyles = [];

    /**
     * @var array
     */
    protected $numberFormats = [];

    protected $borderAllowed = [
        'left', 'right', 'top', 'bottom',
    ];

    protected $borderStyleAllowed = [
        'thin',
        'medium',
        'thick',
        'dashDot',
        'dashDotDot',
        'dashed',
        'dotted',
        'double',
        'hair',
        'mediumDashDot',
        'mediumDashDotDot',
        'mediumDashed',
        'slantDashDot',
    ];

    protected $horizontalAllowed = ['general', 'left', 'right', 'justify', 'center'];

    protected $verticalAllowed = ['bottom', 'center', 'distributed', 'top'];

    protected $defaultFont = ['size' => '10', 'name' => 'Arial', 'family' => '2'];

    /**
     * @param $numberFormat
     * @param $cellStyleString
     *
     * @return false|int|string
     */
    protected function addCellStyle($numberFormat, $cellStyleString)
    {
        $numberFormatIdx = Support::add2listGetIndex($this->numberFormats, $numberFormat);
        $lookupString = $numberFormatIdx.';'.$cellStyleString;

        return Support::add2listGetIndex($this->cellStyles, $lookupString);
    }

    public function styleFontIndexes(array $cellStyles)
    {
        $fills = ['', '']; //2 placeholders for static xml later
        $fonts = ['', '', '', '']; //4 placeholders for static xml later
        $borders = ['']; //1 placeholder for static xml later
        $styleIndexes = [];
        foreach ($cellStyles as $i => $cellStyleString) {
            $semiColonPos = strpos($cellStyleString, ';');
            $numberFormatIdx = substr($cellStyleString, 0, $semiColonPos);
            $styleJsonString = substr($cellStyleString, $semiColonPos + 1);
            $style = @json_decode($styleJsonString, true);
            $styleIndexes[$i] = ['num_fmt_idx' => $numberFormatIdx]; //initialize entry
            if (isset($style['border']) && is_string($style['border'])) { //border is a comma delimited str
                $borderValue['side'] = array_intersect(explode(',', $style['border']), $this->borderAllowed);
                if (isset($style['border-style']) && in_array($style['border-style'], $this->borderStyleAllowed)) {
                    $borderValue['style'] = $style['border-style'];
                }
                if (isset($style['border-color']) && is_string($style['border-color']) && '#' == $style['border-color'][0]) {
                    $v = substr($style['border-color'], 1, 6);
                    $v = 3 == strlen($v) ? $v[0].$v[0].$v[1].$v[1].$v[2].$v[2] : $v; // expand cf0 => ccff00
                    $border_value['color'] = 'FF'.strtoupper($v);
                }
                $styleIndexes[$i]['border_idx'] = Support::add2listGetIndex($borders, json_encode($borderValue));
            }
            if (isset($style['fill']) && is_string($style['fill']) && '#' == $style['fill'][0]) {
                $v = substr($style['fill'], 1, 6);
                $v = 3 == strlen($v) ? $v[0].$v[0].$v[1].$v[1].$v[2].$v[2] : $v; // expand cf0 => ccff00
                $styleIndexes[$i]['fill_idx'] = Support::add2listGetIndex($fills, 'FF'.strtoupper($v));
            }
            if (isset($style['halign']) && in_array($style['halign'], $this->horizontalAllowed)) {
                $styleIndexes[$i]['alignment'] = true;
                $styleIndexes[$i]['halign'] = $style['halign'];
            }
            if (isset($style['valign']) && in_array($style['valign'], $this->verticalAllowed)) {
                $styleIndexes[$i]['alignment'] = true;
                $styleIndexes[$i]['valign'] = $style['valign'];
            }
            if (isset($style['wrap_text'])) {
                $styleIndexes[$i]['alignment'] = true;
                $styleIndexes[$i]['wrap_text'] = (bool) $style['wrap_text'];
            }

            $font = $this->defaultFont;
            if (isset($style['font-size'])) {
                $font['size'] = floatval($style['font-size']); //floatval to allow "10.5" etc
            }
            if (isset($style['font']) && is_string($style['font'])) {
                if ('Comic Sans MS' == $style['font']) {
                    $font['family'] = 4;
                }
                if ('Times New Roman' == $style['font']) {
                    $font['family'] = 1;
                }
                if ('Courier New' == $style['font']) {
                    $font['family'] = 3;
                }
                $font['name'] = strval($style['font']);
            }
            if (isset($style['font-style']) && is_string($style['font-style'])) {
                if (false !== strpos($style['font-style'], 'bold')) {
                    $font['bold'] = true;
                }
                if (false !== strpos($style['font-style'], 'italic')) {
                    $font['italic'] = true;
                }
                if (false !== strpos($style['font-style'], 'strike')) {
                    $font['strike'] = true;
                }
                if (false !== strpos($style['font-style'], 'underline')) {
                    $font['underline'] = true;
                }
            }
            if (isset($style['color']) && is_string($style['color']) && '#' == $style['color'][0]) {
                $v = substr($style['color'], 1, 6);
                $v = 3 == strlen($v) ? $v[0].$v[0].$v[1].$v[1].$v[2].$v[2] : $v; // expand cf0 => ccff00
                $font['color'] = 'FF'.strtoupper($v);
            }
            if ($font != $this->defaultFont) {
                $styleIndexes[$i]['font_idx'] = Support::add2listGetIndex($fonts, json_encode($font));
            }
        }

        return [
            'fills' => $fills,
            'fonts' => $fonts,
            'borders' => $borders,
            'styles' => $styleIndexes,
        ];
    }
}
