<?php

namespace App\Imports;

use App\Models\Data;
use Maatwebsite\Excel\Concerns\ToModel;
use Maatwebsite\Excel\Concerns\WithStartRow;

class DataImports implements ToModel, WithStartRow
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
    public function model(array $row)
    {
        //format
        $row14 = ($row[14] - 25569) * 86400;
        $created = gmdate("Y-m-d H:i:s", $row14);
        $row15 = ($row[15] - 25569) * 86400;
        $updated = gmdate("Y-m-d H:i:s", $row15);

        //insert
        return new Data([
            'code_jira'         => $row[0],
            'environment'       => $row[1],
            'problem_category'  => $row[2],
            'summary'           => $row[3],
            'zentao_link'       => $row[4],
            'priority'          => $row[5],
            'status'            => $row[6],
            'pending_reason'    => $row[7],
            'target_version'    => $row[8],
            'impact_analyst'    => $row[9],
            'root_cause'        => $row[10],
            'work_around'       => $row[11],
            'reporter'          => $row[12],
            'assignee_to'       => $row[13],
            'created'           => $created,
            'updated'           => $updated
        ]);
    }
}