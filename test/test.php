<?php

require '../vendor/autoload.php';

use YExcel\Excel;

$data = [
    ["a1", "a2", "a3", "a4", "a5", "a6", "a7", "a8", "a9", "a10"],
    ["b1", "b2", "b3", "b4", "b5", "b6", "b7", "b8", "b9", "b10"],
    ["c1", "c2", "c3", "c4", "c5", "c6", "c7", "c8", "c9", "c10"],
    [
        function ($sheet, $colNum, $rowNum) {
            $sheet->setCellValueByColumnAndRow($colNum, $rowNum, "d1");
            $fill = $sheet->getStyleByColumnAndRow($colNum, $rowNum)->getFill();
            $fill->setFillType('solid');
            $fill->getStartColor()->setARGB('FFFFEB18');
        },
        function ($sheet, $colNum, $rowNum) {
            $sheet->setCellValueByColumnAndRow($colNum, $rowNum, "d2");
            $fill = $sheet->getStyleByColumnAndRow($colNum, $rowNum)->getFill();
            $fill->setFillType('solid');
            $fill->getStartColor()->setARGB('FFFFEB18');
        },
        function ($sheet, $colNum, $rowNum) {
            $sheet->setCellValueByColumnAndRow($colNum, $rowNum, "d3");
            $fill = $sheet->getStyleByColumnAndRow($colNum, $rowNum)->getFill();
            $fill->setFillType('solid');
            $fill->getStartColor()->setARGB('FFFFEB18');
        },
        function ($sheet, $colNum, $rowNum) {
            $sheet->setCellValueByColumnAndRow($colNum, $rowNum, "d4");
            $fill = $sheet->getStyleByColumnAndRow($colNum, $rowNum)->getFill();
            $fill->setFillType('solid');
            $fill->getStartColor()->setARGB('FFFFEB18');
        },
        function ($sheet, $colNum, $rowNum) {
            $sheet->setCellValueByColumnAndRow($colNum, $rowNum, "d5");
            $fill = $sheet->getStyleByColumnAndRow($colNum, $rowNum)->getFill();
            $fill->setFillType('solid');
            $fill->getStartColor()->setARGB('FFFFEB18');
        },
        function ($sheet, $colNum, $rowNum) {
            $sheet->setCellValueByColumnAndRow($colNum, $rowNum, "d6");
            $fill = $sheet->getStyleByColumnAndRow($colNum, $rowNum)->getFill();
            $fill->setFillType('solid');
            $fill->getStartColor()->setARGB('FFFFEB18');
        },
        function ($sheet, $colNum, $rowNum) {
            $sheet->setCellValueByColumnAndRow($colNum, $rowNum, "d7");
            $fill = $sheet->getStyleByColumnAndRow($colNum, $rowNum)->getFill();
            $fill->setFillType('solid');
            $fill->getStartColor()->setARGB('FFFFEB18');
        },
        function ($sheet, $colNum, $rowNum) {
            $sheet->setCellValueByColumnAndRow($colNum, $rowNum, "d8");
            $fill = $sheet->getStyleByColumnAndRow($colNum, $rowNum)->getFill();
            $fill->setFillType('solid');
            $fill->getStartColor()->setARGB('FFFFEB18');
        },
        function ($sheet, $colNum, $rowNum) {
            $sheet->setCellValueByColumnAndRow($colNum, $rowNum, "d9");
            $fill = $sheet->getStyleByColumnAndRow($colNum, $rowNum)->getFill();
            $fill->setFillType('solid');
            $fill->getStartColor()->setARGB('FFFFEB18');
        },
        function ($sheet, $colNum, $rowNum) {
            $sheet->setCellValueByColumnAndRow($colNum, $rowNum, "d10");
            $fill = $sheet->getStyleByColumnAndRow($colNum, $rowNum)->getFill();
            $fill->setFillType('solid');
            $fill->getStartColor()->setARGB('FFFFEB18');
        }
    ],
];

$data2 = [
    1 => $data
];

$generator = function ($data) {
    foreach ($data as $row) {
        yield $row;
    }
};


Excel::put('test1.xlsx', $data, null, 0);

Excel::put('test2.xlsx', $generator($data), 'template.xlsx', 2);

Excel::put('test3.xlsx', $data, null, 5);

Excel::put('test3.xls', $data, null, 5, 'Excel5');

Excel::put('test4.xls', $data2, null, 5, 'Excel5', true);

Excel::put('test5.xls', $data2, null, 5, 'Excel5', [1 => 'test11']);

$tmp = [];
foreach (Excel::get('test3.xlsx', null, [1, 2], false) as $row) {
    $tmp[] = $row;
}

echo expectTrue(count($tmp) === count($data) + 3, 'skip blank row 0');

$tmp = [];
foreach (Excel::get('test3.xlsx', null, [1], true) as $row) {
    $tmp[] = $row;
}

echo expectTrue(count($tmp) === count($data), 'skip blank row 1');

$tmp = [];
foreach (Excel::get('test3.xls', null, [1], true, 'Excel5') as $row) {
    $tmp[] = $row;
}

echo expectTrue(count($tmp) === count($data), 'skip blank row 1 (xls)');

echo "\n";

$tmp = [];
foreach (Excel::get('test4.xls', null, [1], true, 'Excel5', 1) as $row) {
    $tmp[] = $row;
}

echo expectTrue(count($tmp) === count($data), 'skip blank row 1 (xls) | sheet index 1');

echo "\n";

$tmp = [];
foreach (Excel::get('test5.xls', null, [1], true, 'Excel5', 0) as $row) {
    $tmp[] = $row;
}

echo expectTrue(count($tmp) === 0, 'skip blank row 1 (xls) | sheet index 0');

echo "\n";

foreach (Excel::get('test2.xlsx', null, []) as $row) {
    echo "\n".implode("\t", $row);
}

function expectTrue($flag, $name)
{
    return "\n".($flag ? '√' : '×').' '.$name;
}
