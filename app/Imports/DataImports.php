<?php

namespace App\Imports;

use App\Models\Data;
use Maatwebsite\Excel\Concerns\ToModel;
use Maatwebsite\Excel\Concerns\WithStartRow;
use Carbon\Carbon;
use Maatwebsite\Excel\Concerns\WithMultipleSheets;

class DataImports implements ToModel, WithStartRow, WithMultipleSheets
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
        $hyperlink = $row[0];// Gunakan regular expression untuk mendapatkan kode Jira
        preg_match('/browse\/([A-Z]+-\d+)/', $hyperlink, $matches);
        
        // $matches[1] akan berisi 'DIPM-3681'
        $code_jira = $matches[1];
        
        $row17 = ($row[17] - 25569) * 86400;
        $created = gmdate("Y-m-d H:i:s", $row17);
        $row18 = ($row[18] - 25569) * 86400;
        $updated = gmdate("Y-m-d H:i:s", $row18);
        $row19 = ($row[19] - 25569) * 86400;
        $changed = gmdate("Y-m-d H:i:s", $row19);

        // set RCA Time
        if ($row[20] == null) {
            $rca_time = null;
        } else {
            $row20 = ($row[20] - 25569) * 86400;
            $rca_time = gmdate("Y-m-d H:i:s", $row20);
        }

        // set Closed Time
        if ($row[21] == null) {
            $closed_time = null;
        } else {
            $row21 = ($row[21] - 25569) * 86400;
            $closed_time = gmdate("Y-m-d H:i:s", $row21);
        }

        // pisahkan problem & category 
        $problem_category = $row[2];
        $array_problem_cat = explode(" - ", $problem_category);
        $problem = $array_problem_cat[0];
        $category = $array_problem_cat[1] ?? "-";

        $data = [
            'code_jira'         => $code_jira,
            'environment'       => $row[1],
            'problem'           => $problem,
            'category'          => $category,
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
            'description'       => $row[14],
            'frequent'          => $row[15],
            'complain_info'     => $row[16],
            'created'           => $created,
            'updated'           => $updated,
            'changed_at'        => $changed,
            'rca_time'          => $rca_time,
            'closed_time'       => $closed_time
        ];

        $assignee_too = $row[13];
        if ($assignee_too == 'Ahmad Syauqi') {
            $data = array_merge($data, [
                'nickname' => 'Syauqi',
            ]);
        } else if ($assignee_too == 'Daffa Ramadhan') {
            $data = array_merge($data, [
                'nickname' => 'Daffa',
            ]);
        } else if ($assignee_too == 'Nanda Mahdiaritama Basuki') {
            $data = array_merge($data, [
                'nickname' => 'Nanda',
            ]);
        } else if ($assignee_too == 'Stefano Adrian Sambora') {
            $data = array_merge($data, [
                'nickname' => 'Stefano',
            ]);
        } else if ($assignee_too == 'Ian Daniel Adinata') {
            $data = array_merge($data, [
                'nickname' => 'Daniel',
            ]);
        } else if ($assignee_too == 'Rizki Febrian Aziz') {
            $data = array_merge($data, [
                'nickname' => 'Rizki',
            ]);
        } else if ($assignee_too == 'Tri Intan Siska Permatasari') {
            $data = array_merge($data, [
                'nickname' => 'Intan',
            ]);
        } else if ($assignee_too == 'Alexander Lucas') {
            $data = array_merge($data, [
                'nickname' => 'Lucas',
            ]);
        }

        return new Data($data);
    }
}
