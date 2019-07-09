<?php

/**
 * Class XlsxBuilder.
 */

namespace Zhaqq\Xlsx\Writer;

use Zhaqq\Xlsx\Support;

class XlsxBuilder
{
    /**
     * @param string $company
     *
     * @return string
     */
    public static function buildAppXML($company = '')
    {
        $company = Support::xmlSpecialChars($company);

        return <<<EOF
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><TotalTime>0</TotalTime><Company>$company</Company></Properties>
EOF;
    }

    /**
     * @param string $title
     * @param string $subject
     * @param string $author
     * @param array  $keywords
     * @param string $description
     *
     * @return string
     */
    public static function buildCoreXML(
        string $title = '',
        string $subject = '',
        string $author = '',
        array $keywords = [],
        string $description = ''
    ) {
        $title = Support::xmlSpecialChars($title);
        $subject = Support::xmlSpecialChars($subject);
        $author = Support::xmlSpecialChars($author);
        $keywords = Support::xmlSpecialChars(implode(',', $keywords));
        if ($keywords) {
            $keywords = '<cp:keywords>'.$keywords.'</cp:keywords>';
        }
        $description = Support::xmlSpecialChars($description);
        $date = date("Y-m-d\TH:i:s.00\Z");

        return <<<EOF
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dcterms:created xsi:type="dcterms:W3CDTF">$date</dcterms:created>
<dc:title>$title</dc:title><dc:subject>$subject</dc:subject><dc:creator>$author</dc:creator>$keywords<dc:description>$description</dc:description>
<cp:revision>0</cp:revision>
</cp:coreProperties>
EOF;
    }

    /**
     * @return string
     */
    public static function buildRelationshipsXML()
    {
        $relsXml = '';
        $relsXml .= '<?xml version="1.0" encoding="UTF-8"?>'."\n";
        $relsXml .= '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        $relsXml .= '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>';
        $relsXml .= '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>';
        $relsXml .= '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>';
        $relsXml .= "\n";
        $relsXml .= '</Relationships>';

        return $relsXml;
    }

    /**
     * @param Sheet[] $sheets
     *
     * @return string
     */
    public static function buildWorkbookXML(array $sheets)
    {
        $i = 0;
        $workbookXml = <<<'EOF'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><fileVersion appName="Calc"/><workbookPr backupFile="false" showObjects="all" date1904="false"/><workbookProtection/><bookViews><workbookView activeTab="0" firstSheet="0" showHorizontalScroll="true" showSheetTabs="true" showVerticalScroll="true" tabRatio="212" windowHeight="8192" windowWidth="16384" xWindow="0" yWindow="0"/></bookViews><sheets>
EOF;
        foreach ($sheets as $sheetName => $sheet) {
            $sheetname = Support::sanitizeSheetname($sheet->sheetname);
            $workbookXml .= '<sheet name="'.Support::xmlSpecialChars($sheetname).
                '" sheetId="'.($i + 1).'" state="visible" r:id="rId'.($i + 2).'"/>';
            ++$i;
        }
        $workbookXml .= '</sheets><definedNames>';
        foreach ($sheets as $sheetName => $sheet) {
            if ($sheet->autoFilter) {
                $sheetname = Support::sanitizeSheetname($sheet->sheetname);
                $workbookXml .= '<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">\''.
                    Support::xmlSpecialChars($sheetname).'\'!$A$1:'.
                    Support::xlsCell($sheet->rowCount - 1, count($sheet->columns) - 1, true).'</definedName>';
                ++$i;
            }
        }
        $workbookXml .= '</definedNames><calcPr iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.001"/></workbook>';

        return $workbookXml;
    }

    /**
     * @param Sheet[] $sheets
     *
     * @return string
     */
    public static function buildWorkbookRelsXML(array $sheets)
    {
        $i = 0;
        $wkbkrelsXml = <<<'EOF'
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
EOF;
        foreach ($sheets as $sheetName => $sheet) {
            $wkbkrelsXml .= '<Relationship Id="rId'.($i + 2).
                '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/'.
                ($sheet->xmlname).'"/>';
            ++$i;
        }
        $wkbkrelsXml .= "\n".'</Relationships>';

        return $wkbkrelsXml;
    }

    /**
     * @param Sheet[] $sheets
     *
     * @return string
     */
    public static function buildContentTypesXML(array $sheets)
    {
        $contentTypesXml = '';
        $contentTypesXml .= '<?xml version="1.0" encoding="UTF-8"?>'."\n";
        $contentTypesXml .= '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
        $contentTypesXml .= '<Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
        $contentTypesXml .= '<Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
        foreach ($sheets as $sheetName => $sheet) {
            $contentTypesXml .= '<Override PartName="/xl/worksheets/'.($sheet->xmlname).'" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
        }
        $contentTypesXml .= '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
        $contentTypesXml .= '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';
        $contentTypesXml .= '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
        $contentTypesXml .= '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>';
        $contentTypesXml .= "\n";
        $contentTypesXml .= '</Types>';

        return $contentTypesXml;
    }
}
