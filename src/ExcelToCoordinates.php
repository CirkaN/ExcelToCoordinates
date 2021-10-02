<?php

namespace ExcelToCoordinates;

use ExcelToCoordinates\Excel\ReadFilter;
use GuzzleHttp\Client;
use PhpOffice\PhpSpreadsheet\IOFactory;

class ExcelToCoordinates
{
    private int $startRow = 0;

    private int $endRow = 0;

    private string $path;

    private string $type;

    private array $range;

    private string $addressIterator;

    private string $postCodeIterator;

    private string $countryIterator;

    private string $googleApiCode;

    private Client $client;

    public function __construct()
    {
        $this->client = new Client();
    }

    public function setExcelPath(string $path): self
    {
        $this->path = $path;

        return $this;
    }

    public function setExcelType(string $type): self
    {
        $this->type = $type;

        return $this;
    }

    /**
     * @throws \Exception
     */
    public function setRange(): array
    {
        $range = [$this->getAddressIterator(), $this->getPostCodeIterator(), $this->getCountryIterator()];

        return $this->range = $range;
    }

    public function setStartRow(int $startRow): self
    {
        $this->startRow = $startRow;

        return $this;
    }

    public function setEndRow(int $endRow): self
    {
        $this->endRow = $endRow;

        return $this;
    }

    /**
     * @return string
     * @throws \Exception
     */
    public function getExcelPath(): string
    {
        if ($this->path) {
            return $this->path;
        }

        throw new \Exception('Please set excel file path');
    }

    /**
     * @return string
     * @throws \Exception
     */
    public function getExcelType(): string
    {
        if ($this->type) {
            return $this->type;
        }

        throw new \Exception('Please set excel type');
    }

    /**
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     * @throws \Exception
     * @throws \GuzzleHttp\Exception\GuzzleException
     */
    public function loadData(): array
    {
        $this->setRange();
        $readFilter = new ReadFilter($this->startRow, $this->endRow, $this->range);
        $reader = IOFactory::createReader($this->getExcelType());
        $reader->setReadFilter($readFilter);
        $spreadsheet = $reader->load($this->getExcelPath());

        $addresses = $spreadsheet->getActiveSheet()->toArray(null, true, true, true);

        return $this->formatAddresses($addresses);
    }

    /**
     * @throws \Exception
     * @throws \GuzzleHttp\Exception\GuzzleException
     */
    public function formatAddresses(array $addresses): array
    {
        $formattedAddresses = [];
        foreach ($addresses as $address) {
            $formattedAddresses[] = $address[$this->getAddressIterator()] . ' ' . $address[$this->getPostCodeIterator()] . ' ' . $address[$this->getCountryIterator()];
        }

        return $this->fetchGeoLocation($formattedAddresses);
    }

    /**
     * @throws \GuzzleHttp\Exception\GuzzleException
     * @throws \Exception
     */
    public function fetchGeoLocation(array $addresses): array
    {

        $coordinates = [];
        foreach ($addresses as $address) {
            $response = $this->client->request('POST', 'https://maps.googleapis.com/maps/api/geocode/json?key=' . $this->getGoogleApiCode() . '&address=' . $address);
            $json = json_decode($response->getBody()->getContents());

            if ($json->status !== 'ZERO_RESULTS') {
                $latitude = $json->results[0]->geometry->location->lat;
                $longitude = $json->results[0]->geometry->location->lng;
                $coordinates[] = ['lat' => $latitude, 'long' => $longitude, 'address' => $address];
            }
        }

        return $coordinates;
    }

    /**
     * @return string
     * @throws \Exception
     */
    public function getAddressIterator(): string
    {
        if ($this->addressIterator) {
            return $this->addressIterator;
        }

        throw new \Exception('Please set address iterator');
    }

    /**
     * @param string $addressIterator
     */
    public function setAddressIterator(string $addressIterator): self
    {
        $this->addressIterator = $addressIterator;

        return $this;
    }

    /**
     * @return string
     * @throws \Exception
     */
    public function getPostCodeIterator(): string
    {
        if ($this->postCodeIterator) {
            return $this->postCodeIterator;
        }

        throw new \Exception('Please set post code iterator');
    }

    /**
     * @param string $postCodeIterator
     */
    public function setPostCodeIterator(string $postCodeIterator): self
    {
        $this->postCodeIterator = $postCodeIterator;

        return $this;
    }

    /**
     * @return string
     * @throws \Exception
     */
    public function getCountryIterator(): string
    {
        if ($this->countryIterator) {
            return $this->countryIterator;
        }

        throw new \Exception('Please set country iterator');
    }

    /**
     * @param string $countryIterator
     */
    public function setCountryIterator(string $countryIterator): self
    {
        $this->countryIterator = $countryIterator;

        return $this;
    }

    /**
     * @return string
     * @throws \Exception
     */
    public function getGoogleApiCode(): string
    {
        if ($this->googleApiCode) {
            return $this->googleApiCode;
        }

        throw new \Exception('Please set your google api code');
    }

    /**
     * @param string $googleApiCode
     */
    public function setGoogleApiCode(string $googleApiCode): self
    {
        $this->googleApiCode = $googleApiCode;

        return $this;
    }
}
