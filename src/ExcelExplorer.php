<?php

namespace SunDrop;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\RowIterator;
use PhpOffice\PhpSpreadsheet\Writer\Exception;

/**
 * *) Create csv file
 * *) Create .csv.zip
 * *) Create .xls
 * *) Create .xls.zip
 * *) Create .xlsx
 * *) Create .xlsx.zip
 * *) Calculate all files size
 * *) Drop all files
 */
class ExcelExplorer
{
    const LIMITS = [1, 100, 1000, 5000, 10000, 30000, 50000, 100000, 200000, 500000, 1000000,];

    // Just Tmp Files
    private $filenameCsv;
    private $filenameXls;
    private $filenameXlsx;

    private $fpCsv;
    /** @var Spreadsheet */
    private $spreadsheet;
    /** @var Worksheet */
    private $worksheet;
    /** @var RowIterator */
    private $xlsRowIterator;

    /**
     * @return \Generator
     */
    public function getFilesSize()
    {
        foreach (self::LIMITS as $limit) {
            $count = 0;
            $executionStartTime = microtime(true);
            $this->initFiles();

            foreach ($this->getNextRow() as $row) {
                if ($count++ > $limit) {
                    break;
                }
                // add .csv .xls
                $this->putRow($row);
            }
            // Save files and calculate size
            $filesSize = $this->calculateFilesSize();
            $this->deleteFiles();

            $executionEndTime = microtime(true);
            $seconds = $executionEndTime - $executionStartTime;
            yield ['count' => $count - 2, 'seconds' => $seconds] + $filesSize;
        }
    }

    private function putRow($row)
    {
        \fputcsv($this->fpCsv, $row);

        $xlsRow = $this->xlsRowIterator->current();
        $cellIterator = $xlsRow->getCellIterator();
        foreach ($row as $item) {
            $cellIterator->current()->setValue($item);
            $cellIterator->next();
        }
        $this->xlsRowIterator->next();
    }

    private function calculateFilesSize()
    {
        $filesSize = [];
        \fclose($this->fpCsv);
        $filesSize['.csv'] = $this->formatSizeUnits(\filesize($this->filenameCsv));
        $filesSize['.csv.zip'] = $this->formatSizeUnits($this->getZippedSize($this->filenameCsv));

        try {
            // If we try to write more than 65k rows
            $xlsWriter = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($this->spreadsheet, 'Xls');
            $xlsWriter->save($this->filenameXls);
            $filesSize['.xls'] = $this->formatSizeUnits(\filesize($this->filenameXls));
            $filesSize['.xls.zip'] = $this->formatSizeUnits($this->getZippedSize($this->filenameXls));
        } catch (Exception $e) {
            $filesSize['.xls'] = 0;
            $filesSize['.xls.zip'] = 0;
        }

        $xlsxWriter = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($this->spreadsheet, 'Xlsx');
        $xlsxWriter->save($this->filenameXlsx);
        $filesSize['.xlsx'] = $this->formatSizeUnits(\filesize($this->filenameXlsx));
        $filesSize['.xlsx.zip'] = $this->formatSizeUnits($this->getZippedSize($this->filenameXlsx));

        return $filesSize;
    }

    private function getZippedSize($filename)
    {
        $zipFilename = \tempnam(\sys_get_temp_dir(), 'zip');
        $z = new \ZipArchive();
        $z->open($zipFilename, \ZipArchive::CREATE);
        $z->addFile($filename);
        $z->setCompressionIndex(0, \ZipArchive::CM_DEFLATE);
        $z->close();
        $filesize = \filesize($zipFilename);
        \unlink($zipFilename);

        return $filesize;
    }

    private function initFiles()
    {
        $this->filenameCsv = \tempnam(\sys_get_temp_dir(), 'exp');
        $this->fpCsv = \fopen($this->filenameCsv, 'w');
        $this->filenameXls = \tempnam(\sys_get_temp_dir(), 'exp');
        $this->filenameXlsx = \tempnam(\sys_get_temp_dir(), 'exp');

        $this->spreadsheet = new Spreadsheet();
        $this->worksheet = $this->spreadsheet->getActiveSheet();
        $this->xlsRowIterator = new RowIterator($this->worksheet);
    }

    private function deleteFiles()
    {
        \unlink($this->filenameCsv);
        \unlink($this->filenameXls);
        \unlink($this->filenameXlsx);
    }

    private function getNextRow()
    {
        yield ['Email', 'Status', 'Reads', 'Clicks', 'Name', 'Reason']; // Header
        while (1) {
            yield [
                $this->generateRandomString() . '@email.com',
                'StatusOK',
                \rand(0, PHP_INT_MAX),
                \rand(0, PHP_INT_MAX),
                $this->generateRandomString(),
                $this->generateRandomString(200),
            ];
        }
    }

    private function generateRandomString($length = 10)
    {
        $characters = '0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ';
        $charactersLength = \strlen($characters);
        $randomString = '';
        for ($i = 0; $i < $length; $i++) {
            $randomString .= $characters[\rand(0, $charactersLength - 1)];
        }
        return $randomString;
    }

    private function formatSizeUnits($bytes)
    {
        if ($bytes >= 1073741824) {
            $bytes = number_format($bytes / 1073741824, 2) . ' GB';
        } elseif ($bytes >= 1048576) {
            $bytes = number_format($bytes / 1048576, 2) . ' MB';
        } elseif ($bytes >= 1024) {
            $bytes = number_format($bytes / 1024, 2) . ' KB';
        } elseif ($bytes > 1) {
            $bytes = $bytes . ' bytes';
        } elseif ($bytes == 1) {
            $bytes = $bytes . ' byte';
        } else {
            $bytes = '0 bytes';
        }

        return $bytes;
    }

}