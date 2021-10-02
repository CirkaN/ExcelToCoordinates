<?php

namespace Vendor\Package\Tests;

use ExcelToCoordinates\ExcelToCoordinates;

use PHPUnit\Framework\TestCase;

class ExcelToCoordinatesTest extends TestCase
{
    public function testGoogleApiCodeNotSet()
    {
        $this->expectException(\Exception::class);
        $t = new ExcelToCoordinates();
        $t->setGoogleApiCode('')->getGoogleApiCode();
    }

    public function testCountryIteratorNotSet()
    {
        $this->expectException(\Exception::class);
        $t = new ExcelToCoordinates();
        $t->setCountryIterator('')->getCountryIterator();
    }

    public function testAddressIteratorNotSet()
    {
        $this->expectException(\Exception::class);
        $t = new ExcelToCoordinates();
        $t->setAddressIterator('')->getAddressIterator();
    }

    public function testPostCodeIteratorNotSet()
    {
        $this->expectException(\Exception::class);
        $t = new ExcelToCoordinates();
        $t->setPostCodeIterator('')->getPostCodeIterator();
    }

    public function testExcelTypeNotSet()
    {
        $this->expectException(\Exception::class);
        $t = new ExcelToCoordinates();
        $t->setExcelType('')->getExcelType();
    }

    public function testExcelPathNotSet()
    {
        $this->expectException(\Exception::class);
        $t = new ExcelToCoordinates();
        $t->setExcelPath('')->getExcelPath();
    }

    public function testExcelPathIsWrong()
    {
        $this->expectException(\InvalidArgumentException::class);
        $t = new ExcelToCoordinates();
        $t->setExcelType('Xlsx')
            ->setExcelPath("/../example/addresses.xlsx")
            ->setAddressIterator('A')
            ->setPostCodeIterator('B')
            ->setCountryIterator('C')
            ->setGoogleApiCode('2344534253425')
            ->setStartRow(1)
            ->setEndRow(2)
            ->loadData();
    }

    public function testInvalidExcelType()
    {
        $this->expectException(\PhpOffice\PhpSpreadsheet\Reader\Exception::class);
        $t = new ExcelToCoordinates();
        $t->setExcelType('my_wrong_excel_type')
            ->setExcelPath(__DIR__ . "/../example/addresses.xlsx")
            ->setAddressIterator('A')
            ->setPostCodeIterator('B')
            ->setCountryIterator('C')
            ->setGoogleApiCode('2344534253425')
            ->setStartRow(1)
            ->setEndRow(2)
            ->loadData();
    }

    public function testLoadFunction()
    {
        if (! getenv('GOOGLE_API_KEY')) {
            $this->markTestSkipped('Missing Google Api Key Credentials');
        }

        $t = new ExcelToCoordinates();
        $output = $t->setExcelType('Xlsx')
            ->setExcelPath(__DIR__ . "/../example/addresses.xlsx")
            ->setAddressIterator('A')
            ->setPostCodeIterator('B')
            ->setCountryIterator('C')
            ->setGoogleApiCode(getenv('GOOGLE_API_KEY'))
            ->setStartRow(2)
            ->setEndRow(2)
            ->loadData();


        $this->assertSame(['lat' => 44.7859479, 'long' => 20.4814831, 'address' => 'Ustanicka 11320 Serbia'], $output[0]);
    }
}
