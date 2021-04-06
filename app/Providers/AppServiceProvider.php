<?php

namespace App\Providers;

use Illuminate\Database\Eloquent\Collection;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Symfony\Component\HttpFoundation\StreamedResponse;
use Illuminate\Support\ServiceProvider;
use Illuminate\Support\Facades\Schema;

class AppServiceProvider extends ServiceProvider
{
    /**
     * Register any application services.
     *
     * @return void
     */
    public function register()
    {
        //
    }

    /**
     * Bootstrap any application services.
     *
     * @return void
     */
    public function boot()
    {
        Schema::defaultStringLength(191);

        Collection::macro('exportExcel', function ($filename = '', $filters = []) {

            if (empty($filename)) {

                $filename = $this->first()->getTable() . '_' . date('Ymd_His') . '.xlsx';
            }

            $spreadsheet = new Spreadsheet();
            $sheet = $spreadsheet->getActiveSheet();
            $excel_data = [];


            $this->each(function ($item) use (&$excel_data, $filters) {

                $row_data = [];

                if (empty($filters)) {

                    $row_data = array_values($item->toArray());
                } else {

                    foreach ($filters as $filter) {

                        if (is_callable($filter)) {

                            $row_data[] = $filter($item);
                        }
                    }
                }

                $excel_data[] = $row_data;
            });
            $sheet->fromArray($excel_data, null, 'A1');

            $callback = function () use ($spreadsheet) {

                $writer = new Xlsx($spreadsheet);
                $writer->save('php://output');
            };
            $status = 200;
            $headers = [
                'Content-Type' => 'application/vnd.ms-excel',
                'Content-Disposition' => 'attachment;filename="' . $filename . '"',
                'Cache-Control' => 'max-age=0',
            ];
            return new StreamedResponse($callback, $status, $headers);
        });
    }
}
