# Package

## Usage
    Start by creating a new instance of ExcelToCoorindates class:

    $excel_to_cordinates = new ExcelToCoordinates();
## Methods
            ->setExcelPath(__DIR__ . "/../example/addresses.xlsx") // path to the excel file
            ->setAddressIterator('A') // (required) The address iterator, for better understanding please check the example excel file provided in /example
            ->setPostCodeIterator('B') // (required) The Post Code Iterator 
            ->setCountryIterator('C') // (required) The country Iterator
            ->setGoogleApiCode('') //Google API Key
            ->setStartRow(1) //row from where you want to start 
            ->setEndRow(2) // row where converting will stop
## Installation
    `composer require cirkovic/excel-to-coordinates`
## Requirements
    * php 7.4
    * ext-json
## Testing
Test can be found in /tests.

## Contributing
Everyone is welcome for contribution.
## Credits
- <a href="https://github.com/CirkaN">Cirkovic Nikola<a>
- <a href="https://github.com/CirkaN/ExcelToCoordinates/graphs/contributors">All Contributors </a>
## Licence
