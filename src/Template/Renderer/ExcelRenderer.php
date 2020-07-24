<?php
namespace Burdock\Template\Renderer;

use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class ExcelRenderer implements IRenderer
{
    /**
     * @param string $template_path テンプレートファイルへのパス
     * @param array $config レンダリング設定
     * @param array $data データ
     * @param string $output_path
     * @return void
     * @throws Exception
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public static function render(string $template_path, array $config, array $data, string $output_path): void
    {
        $reader = IOFactory::createReader('Xlsx');
        $excel = $reader->load($template_path);

        $rows_per_page = self::splitRowsPerPage($config, $data);
        $embedded = [
            '@TOTAL_PAGES' => count($rows_per_page),
            '@PAGE_NUM' => 1,
            '@rownum' => 0
        ];
        foreach ($rows_per_page as $page_num => $rows) {
            $embedded['@PAGE_NUM'] = $page_num + 1;
            $page_cnf = ($page_num == 0) ? $config['pages'][0] : $config['pages'][1];
            $summary  = $data['summary'];
            $rownum = 1;
            if ($page_num === 0) {
                $sheet = $excel->getSheet($page_num);
                $sheet->setTitle('page '.$embedded['@PAGE_NUM']);
                self::renderPage($page_cnf, $sheet, $summary, $rows, $embedded);
            } else {
                $sheet = clone $excel->getSheet(1);
                self::renderPage($page_cnf, $sheet, $summary, $rows, $embedded);
                $sheet->setTitle('page '.$embedded['@PAGE_NUM']);
                $excel->addSheet($sheet);
            }
        }
        if ($excel->getSheetCount() > 1)
            $excel->removeSheetByIndex(1);

        $writer = IOFactory::createWriter($excel, 'Xlsx');
        $writer->save($output_path);
        $excel->disconnectWorksheets();
        unset($excel);
    }

    public static function splitRowsPerPage(array $config, array $data): array
    {
        $details_per_page = [];
        $page_rows = [];
        $page_num = 0;
        $page_cnf = $config['pages'][$page_num]['rows'];

        for ($i = 0; $i < count($data['rows']); $i++) {
            $page_rows[] = $data['rows'][$i];
            if (count($page_rows) == $page_cnf['row_max']) {
                $details_per_page[] = $page_rows;
                $page_num++;
                $page_rows = [];
                $page_cnf = $config['pages'][1]['rows'];
            }
        }

        if (count($page_rows) > 0) {
            $details_per_page[] = $page_rows;
        }

        return $details_per_page;
    }

    /**
     * @param array $page_cnf ぺージ設定
     * @param Worksheet $sheet 操作対象シート
     * @param array $summary サマリデータ
     * @param array $rows 明細データ
     * @param array $embedded 組み込み変数
     * @return void
     */
    public static function renderPage(array $page_cnf, Worksheet $sheet, array $summary, array $rows, array &$embedded): void
    {
        foreach ($page_cnf['summary']['mappings'] as $cell => $_field) {
            list($field, $styles) = is_array($_field) ? $_field : [$_field, null];
            if (preg_match('/{.+}/', $field)) {
                $value = preg_replace_callback('/{(@\w+)}/', function ($m) use ($embedded) {
                    return $embedded[$m[1]];
                }, $field);
                $sheet->setCellValue($cell,  $value);
            } else {
                if (!array_key_exists($field, $summary)) continue;
                $value = $summary[$field];
                self::setValue($sheet, $cell, $styles, $value);
            }
        }
        $tpl_start   = $page_cnf['rows']['start_idx'];
        foreach ($rows as $idx => $row) {
            $embedded['@rownum']++;
            foreach ($page_cnf['rows']['mappings'] as $column => $_field) {
                list($field, $styles) = is_array($_field) ? $_field : [$_field, null];
                if (!array_key_exists($field, $row) && $field !== '@rownum') continue;
                $row_idx  = (int)$tpl_start + (int)$idx;
                $cell = $column . $row_idx;
                $value = ($field === '@rownum') ? $embedded['@rownum'] : $row[$field];
                self::setValue($sheet, $cell, $styles, $value);
            }
        }
    }

    public static function setValue(Worksheet $sheet, string $cell, ?array $styles, $value)
    {
        if (isset($styles['wrap_text'])) {
            $sheet->getStyle($cell)->getAlignment()->setWrapText(true);
            $value = str_replace("\\n", "\n", $value);
        }
        $sheet->setCellValue($cell, $value);
    }
}