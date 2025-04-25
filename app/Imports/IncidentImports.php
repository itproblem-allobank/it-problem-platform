<?php

namespace App\Imports;

use App\Models\Incident;
use Maatwebsite\Excel\Concerns\ToModel;
use Maatwebsite\Excel\Concerns\WithStartRow;
use Carbon\Carbon;
use Maatwebsite\Excel\Concerns\WithMultipleSheets;

class IncidentImports implements ToModel, WithStartRow, WithMultipleSheets
{
    /**
     * @param array $row
     *
     * @return \Illuminate\Database\Eloquent\Model|null
     */

    public function startRow(): int
    {
        return 2;
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

        // pisahkan incident & category 
        $incident_category = $row[2];
        $array_incident_cat = explode(" - ", $incident_category);
        $incident = $array_incident_cat[0];
        $category = $array_incident_cat[1] ?? "-";


        // // convert hyperlink to code jira
        // $hyperlink = $row[0]; // Gunakan regular expression untuk mendapatkan kode Jira
        // preg_match('/browse\/([A-Z]+-\d+)/', $hyperlink, $matches);

        // // $matches[1] akan berisi 'DIPM-3681'
        // $no_jira = $matches[1];

        // dd($row[7], $row[8]);
        $created = ($row[7] - 25569) * 86400;
        $created_time = gmdate("Y-m-d H:i:s", $created);

        // resolved time
        if ($row[8] == null) {
            $resolved_time = null;
        } else {
            $resolved = ($row[7] - 25569) * 86400;
            $resolved_time = gmdate("Y-m-d H:i:s", $resolved);
        }


        $data = [
            'no_jira'         => $row[0],
            'summary'         => $row[1],
            'incident'         => $incident,
            'category'        => $category,
            'status_ticket'   => $row[3],
            'priority'        => $row[4],
            'rootcause'       => $row[5],
            'mitigation'      => $row[6],
            'created_time'    => $created_time,
            'resolved_time'   => $resolved_time

        ];

        return new Incident($data);
    }
}
