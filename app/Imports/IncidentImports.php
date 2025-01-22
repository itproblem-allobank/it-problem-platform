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
        return 8;
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
        // convert hyperlink to code jira
        $hyperlink = $row[0]; // Gunakan regular expression untuk mendapatkan kode Jira
        preg_match('/browse\/([A-Z]+-\d+)/', $hyperlink, $matches);

        // $matches[1] akan berisi 'DIPM-3681'
        $no_jira = $matches[1];

        $created = ($row[6] - 25569) * 86400;
        $created_time = gmdate("Y-m-d H:i:s", $created);
        $resolved = ($row[7] - 25569) * 86400;
        $resolved_time = gmdate("Y-m-d H:i:s", $resolved);

        $data = [
            'no_jira'         => $no_jira,
            'summary'         => $row[1],
            'status_ticket'   => $row[2],
            'disruption'      => $row[3],
            'rootcause'       => $row[4],
            'mitigation'      => $row[5],
            'created_time'    => $created_time,
            'resolved_time'   => $resolved_time
            
        ];

        return new Incident($data);
    }
}
