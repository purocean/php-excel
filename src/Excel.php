<?php

namespace YExcel;

use \PHPExcel_IOFactory as IOFactory;
use \PHPExcel_Cell as Cell;
use \PHPExcel as PHPExcel;
use Closure;

/**
* 导入导出
*/
class Excel
{
    /**
     * 从 Excel 获取所有行
     *
     * @param string   $file          xlsx 文件路径
     * @param int|null $highestColumn 列数，为 null 时候自动检测
     * @param array    $skipRows      跳过的行，默认跳过第一行（表头）
     * @param bool     $skipBlankRow  是否跳过空白行，默认为 true
     * @param string   $type          Excel 文件类型，如 Excel2007 或 Excel5
     * @param int      $sheetNum      Sheet index
     *
     * @return generator 可遍历的生成器
     */
    public static function get($file, $highestColumn = null, $skipRows = [1], $skipBlankRow = true, $type = 'Excel2007', $sheetNum = 0)
    {

        $objReader = IOFactory::createReader($type);
        $objPHPExcel = $objReader->load($file);
        $sheet = $objPHPExcel->getSheet($sheetNum);
        $highestRow = $sheet->getHighestRow();
        is_null($highestColumn) and $highestColumn = Cell::columnIndexFromString($sheet->getHighestColumn());

        for ($row = 1; $row <= $highestRow; ++$row) {
            if (in_array($row, $skipRows)) {
                continue;
            }

            $rowData = [];
            for ($col = 0; $col < $highestColumn; $col++) {
                $value = $objPHPExcel->getActiveSheet()->getCellByColumnAndRow($col, $row)->getValue();
                $rowData[] = is_null($value) ? '' : (string) $value;
            }

            if ($skipBlankRow) {
                if (!array_filter($rowData)) {
                    continue;
                }
            }

            yield $rowData;
        }
    }

    /**
     * 把数据写入一个文件
     *
     * @param string          $file          文件名
     * @param array|generator $data          数据，可以被 foreach 遍历的数据，数组或者生成器，如果多 sheet ，传入二维数组
     * @param string          $tplFile       模板文件，以哪个模板填写数据，如果不提供则生成空白 xlsx 文件
     * @param int             $skipRow       跳过表头的行数，默认为 1
     * @param string          $type          Excel 文件类型，如 Excel2007 或 Excel5
     * @param array|boolean   $multiSheet    是否要写入多 sheet，可以传入数组定义 sheet 名称，[0 => 'sheet1', 1 => 'sheet2']
     */
    public static function put($file, $data, $tplFile = null, $skipRow = 1, $type = 'Excel2007', $multiSheet = false)
    {
        if ($tplFile) {
            if (file_exists($tplFile)) {
                $objReader = IOFactory::createReader($type);
                $objPHPExcel = $objReader->load($tplFile);
            } else {
                throw new \Exception("File `{$tplFile}` not exists");
            }
        } else {
            $objPHPExcel = new PHPExcel();
        }

        $write = function ($data) use ($objPHPExcel, $multiSheet, $skipRow) {
            if (! is_array($multiSheet)) {
                $multiSheet = [];
            }

            foreach ($data as $sheetNum => $sheetData) {
                if ($sheetNum > $objPHPExcel->getSheetCount() - 1) {
                    $objPHPExcel->createSheet($sheetNum);
                }

                $objPHPExcel->setActiveSheetIndex($sheetNum);
                $objSheet = $objPHPExcel->getActiveSheet();
                $objSheet->setTitle(isset($multiSheet[$sheetNum]) ? $multiSheet[$sheetNum] : 'sheet' . ($sheetNum + 1));

                $rowNum = 1;
                foreach ($sheetData as $row) {
                    $colNum = 0;
                    foreach ($row as $val) {
                        if ($val instanceof Closure) {
                            $val($objSheet, $colNum, $rowNum + $skipRow);
                        } else {
                            $objSheet->setCellValueByColumnAndRow(
                                $colNum,
                                $rowNum + $skipRow,
                                $val
                            );
                        }
                        ++$colNum;
                    }
                    ++$rowNum;
                }
            }
        };

        $write($multiSheet === false ? [$data] : $data);
        $objPHPExcel->setActiveSheetIndex(0);

        IOFactory::createWriter($objPHPExcel, $type)->save($file);
    }
}
