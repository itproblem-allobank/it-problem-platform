<?php

namespace App\Imports;

use App\Models\Service;
use Maatwebsite\Excel\Concerns\ToModel;
use Maatwebsite\Excel\Concerns\WithStartRow;
use Carbon\Carbon;
use Maatwebsite\Excel\Concerns\WithMultipleSheets;

class ServiceImports implements ToModel, WithStartRow, WithMultipleSheets
{
    /**
     * @param array $row
     *
     * @return \Illuminate\Database\Eloquent\Model|null
     */

     public function startRow(): int
     {
        // return 8; kalo dari Apps
         return 2; //kalo dari Export Excel
     }
 
     public function sheets(): array
     {
         return
             [
                 1 => $this,
             ];
     }
    public function model(array $row)
    {
        // dd($row[6]);
        $row6 = ($row[6] - 25569) * 86400;
        $created = gmdate("Y-m-d H:i:s", $row6);
        $row7 = ($row[7] - 25569) * 86400;
        $updated = gmdate("Y-m-d H:i:s", $row7);

        $str = $row[2];
        $ctr = explode(" - ", $str);

        $data = [
            'issue_type'    => $row[0],
            'code_jira'     => $row[1],
            'summary'       => $row[2],
            'assignee'      => $row[3],
            'reporter'      => $row[4],
            'status'        => $row[5],
            'created'       => $created,
            'updated'       => $updated,
            'priority'      => $row[8],
            'sub_category'   => $row[9],
            'ticket_number' => $row[10],
            'customer_care_category' => $row[11],
        ];

        return new Service($data);
    }
}
