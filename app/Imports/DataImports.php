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
        // return 8; kalo dari Apps
        return 2;  //kalo dari Export Excel
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

        $row16 = ($row[16] - 25569) * 86400;
        $created = gmdate("Y-m-d H:i:s", $row16);
        $row17 = ($row[17] - 25569) * 86400;
        $updated = gmdate("Y-m-d H:i:s", $row17);
        $row18 = ($row[18] - 25569) * 86400;
        $changed = gmdate("Y-m-d H:i:s", $row18);

        // set RCA Time
        $rca_fix = null;
        if ($row[19] == null) {
            $rca_time = null;
            $rca_days = null;
        } else {
            $row19 = ($row[19] - 25569) * 86400;
            $rca_time = gmdate("Y-m-d H:i:s", $row19);
            $rca_days = Carbon::parse($created)->diffInDays(Carbon::parse($rca_time));
            // dd($rca_days);
            if ($rca_days <= 1) {
                $rca_fix = 1;
            } else {
                $rca_fix = $rca_days;
            }
        }

        // set Closed Time
        if ($row[20] == null) {
            $closed_time = null;
        } else {
            $row20 = ($row[20] - 25569) * 86400;
            $closed_time = gmdate("Y-m-d H:i:s", $row20);
        }

        // set Target Date
        if ($row[24] == null) {
            $target_date = null;
        } else {
            $row24 = ($row[24] - 25569) * 86400;
            $target_date = gmdate("Y-m-d H:i:s", $row24);
        }

        // pisahkan problem & category 
        $problem_category = $row[3];
        $array_problem_cat = explode(" - ", $problem_category);
        $problem = $array_problem_cat[0];
        $category = $array_problem_cat[1] ?? "-";

        // set resolved days
        if ($row[6] == 'Closed') {
            $resolved_days = Carbon::parse($created)->diffInDays(Carbon::parse($changed));
        } else {
            $resolved_days = null;
        }

        $data = [
            'code_jira'         => $row[0],
            'summary'           => $row[1],
            'environment'       => $row[2],
            'problem'           => $problem,
            'category'          => $category,
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
            'frequent'          => $row[14],
            'complain_info'     => $row[15],
            'created'           => $created,
            'updated'           => $updated,
            'changed_at'        => $changed,
            'rca_time'          => $rca_time,
            'closed_time'       => $closed_time,
            'resolved_days'     => $resolved_days,
            'rca_days'          => $rca_fix,
            'team'              => $row[21],
            'aspect'            => $row[22],
            'zoho_ticket'       => $row[23],
            'target_date'       => $target_date
        ];

        $assignee_too = $row[13];
        if ($assignee_too == 'Ahmad Syauqi') {
            $data = array_merge($data, [
                'nickname' => 'Syauqi',
            ]);
        } else if ($assignee_too == 'Fadel Ferniawan') {
            $data = array_merge($data, [
                'nickname' => 'Fadel',
            ]);
        } else if ($assignee_too == 'Nanda Mahdiaritama Basuki') {
            $data = array_merge($data, [
                'nickname' => 'Nanda',
            ]);
        } else if ($assignee_too == 'Febri Syahri Ramadhan') {
            $data = array_merge($data, [
                'nickname' => 'Febri',
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
        } else if ($assignee_too == 'Fachri Fachri') {
            $data = array_merge($data, [
                'nickname' => 'Fachri',
            ]);
        }

        return new Data($data);
    }
}
