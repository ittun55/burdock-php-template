<?php
use Burdock\Template\Renderer\ExcelRenderer;
use PHPUnit\Framework\TestCase;

const DS = DIRECTORY_SEPARATOR;

class ExcelRendererTest extends TestCase
{
    public function testRender()
    {
        $config = [
            'pages' => [
                [
                    'summary' => [
                        'mappings' => [
                            'AM1'  => '{@PAGE_NUM} / {@TOTAL_PAGES}',
                            'AL2'  => 'invoice_no',
                            'AE2'  => 'date',
                            'B3'   => 'customer_postal_code',
                            'B4'   => 'customer_addr1',
                            'B5'   => 'customer_addr2',
                            'B6'   => 'customer_name',
                            'Q29'  => 'subtotal',
                            'AA29' => 'tax',
                            'AJ29' => 'amount',
                        ]
                    ],
                    'rows' => [
                        'start_idx' => 14,
                        'row_max'   => 15,
                        'mappings' => [
                            'A' => 'month',
                            'C' => 'day',
                            'E' => 'site',
                            'L' => 'work_item',
                            'V' => 'roughter_type',
                            'Z' => 'quantity', // amount
                            'AC' => 'work_time',
                            'AG' => 'price',
                        ]
                    ]
                ],
                [
                    'summary' => [
                        'mappings' => [
                            'AM1' => '{@PAGE_NUM} / {@TOTAL_PAGES}',
                            'AL2' => 'invoice_no',
                            'AE2' => 'date',
                            'Q27'  => 'subtotal',
                            'AA27' => 'tax',
                            'AJ27' => 'amount',
                        ]
                    ],
                    'rows' => [
                        'start_idx' => 5,
                        'row_max'   => 22,
                        'mappings' => [
                            'A' => 'month',
                            'C' => 'day',
                            'E' => 'site',
                            'L' => 'work_item',
                            'V' => 'roughter_type',
                            'Z' => 'quantity', // amount
                            'AC' => 'work_time',
                            'AG' => 'price',
                        ]
                    ]
                ]
            ]
        ];
        $json = json_encode($config, JSON_PRETTY_PRINT);
        echo $json . PHP_EOL;
        $data = [
            'summary' => [
                'invoice_no'            => 'ABC123456',
                'date'                  => '令和元年 8 月 27 日',
                'customer_postal_code'  => '〒350-0042',
                'customer_addr1'        => '埼玉県川越市中原町１－９－７',
                'customer_addr2'        => 'アドリーム川越３０３',
                'subtotal'              => '2000000',
            ],
            'rows' => [
                ['month' => '01', 'day' => '01', 'site' => '　現場１', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                [                                'site' => '　現場２', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                [                                'site' => '　現場３', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                ['month' => '01', 'day' => '01', 'site' => '　現場４', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                [                                'site' => '　現場５', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                [                                'site' => '　現場６', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                [                                'site' => '　現場７', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                ['month' => '01', 'day' => '02', 'site' => '　現場８', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                [                                'site' => '　現場９', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                ['month' => '01', 'day' => '03', 'site' => '　現場10', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                [                                'site' => '　現場11', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                [                                'site' => '　現場12', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                ['month' => '01', 'day' => '01', 'site' => '　現場13', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                [                                'site' => '　現場14', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                [                                'site' => '　現場15', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                ['month' => '01', 'day' => '01', 'site' => '　現場16', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                [                                'site' => '　現場17', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                [                                'site' => '　現場18', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                [                                'site' => '　現場19', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                ['month' => '01', 'day' => '02', 'site' => '　現場20', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                [                                'site' => '　現場21', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                ['month' => '01', 'day' => '03', 'site' => '　現場22', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                [                                'site' => '　現場23', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                [                                'site' => '　現場24', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                ['month' => '01', 'day' => '01', 'site' => '　現場25', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                [                                'site' => '　現場26', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                [                                'site' => '　現場27', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                ['month' => '01', 'day' => '01', 'site' => '　現場28', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                [                                'site' => '　現場29', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                [                                'site' => '　現場30', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                [                                'site' => '　現場31', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                ['month' => '01', 'day' => '02', 'site' => '　現場32', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                [                                'site' => '　現場33', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                ['month' => '01', 'day' => '03', 'site' => '　現場34', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                [                                'site' => '　現場35', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                [                                'site' => '　現場36', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                ['month' => '01', 'day' => '03', 'site' => '　現場37', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                [                                'site' => '　現場38', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
                [                                'site' => '　現場39', 'work_item' => '　ＡＢＣ　ＸＹＺ', 'roughter_type' => '１tユニック', 'work_time' => '10:00~17:00', 'price' => 5000],
            ]
        ];
        $output_path = __DIR__ . DS . 'tmp' . DS . Date('YmdHis') . '.xlsx';
        $template_path = __DIR__ . DS . 'tmp' . DS . 'invoice_new.xlsx';
        ExcelRenderer::render($template_path, $config, $data, $output_path);
        $this->assertTrue(file_exists($output_path));
    }
}
