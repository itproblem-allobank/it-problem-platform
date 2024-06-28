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

        $row14 = ($row[14] - 25569) * 86400;
        $created = gmdate("Y-m-d H:i:s", $row14);
        $row15 = ($row[15] - 25569) * 86400;
        $updated = gmdate("Y-m-d H:i:s", $row15);

        $str = $row[2];
        $ctr = explode(" - ", $str);

        $data = [
            'code_jira'         => $row[0],
            'environment'       => $row[1],
            'category'          => $row[2],
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
        ];

        if ($ctr[0] == 'QRIS' || $ctr[0] == 'Transfer' || $ctr[0] == 'Topup Incoming' || $ctr[0] == 'Tabungan' || $ctr[0] == 'Cashout' || $ctr[0] == 'Balance' || $ctr[0] == 'Virtual Debit Card') {
            $data = array_merge($data, [
                'problem' => 'Transaction',
            ]);
        } else if ($ctr[0] == 'Bill Payment' || $ctr[0] == 'E-Wallet' || $ctr[0] == 'Secure Parking' || $ctr[0] == 'SNAP') {
            $data = array_merge($data, [
                'problem' => 'Online Payment',
            ]);
        } else if ($ctr[0] == 'MPC' || $ctr[0] == 'Payment Gateway' || $ctr[0] == 'Topup') {
            $data = array_merge($data, [
                'problem' => 'Ecosistem & MPC',
            ]);
        } else if ($ctr[0] == 'Onboarding') {
            $data = array_merge($data, [
                'problem' => 'Onboarding',
            ]);
        } else if ($ctr[0] == 'Paylater' || $ctr[0] == 'Instant Cash' || $ctr[0] == 'Telemarketing') {
            $data = array_merge($data, [
                'problem' => 'Loan',
            ]);
        } else if ($ctr[0] == 'Surrounding Apps' || $ctr[0] == 'E-Statement' || $ctr[0] == 'Message' || $ctr[0] == 'Server' || $ctr[0] == 'Database' || $ctr[0] == 'Requirement') {
            $data = array_merge($data, [
                'problem' => 'Core System & Surrounding Apps',
            ]);
        } else if ($ctr[0] == 'Temenos' || $ctr[0] == 'IBB' || $ctr[0] == 'BI Applications' || $ctr[0] == 'Bank Devisa' || $ctr[0] == 'Payroll' ) {
            $data = array_merge($data, [
                'problem' => 'Wholesale Banking',
            ]);
        } else {
            $data = array_merge($data, [
                'problem' => '-',
            ]);
        }

        return new Data($data);
        //format
        // $row14 = ($row[14] - 25569) * 86400;
        // $created = gmdate("Y-m-d H:i:s", $row14);
        // $row15 = ($row[15] - 25569) * 86400;
        // $updated = gmdate("Y-m-d H:i:s", $row15);

        // $str = $ctr;
        // $category_explode = explode(" - ", $str);


        // $problem = 'test';
        // if ($category_explode[0] == 'QRIS') {
        //     $problem = 'Transaction';
        // }
        // //insert
        // return new Data([
        //     'code_jira'         => $row[0],
        //     'environment'       => $row[1],
        //     'problem'           => $problem,
        //     'category'          => $row[2],
        //     'summary'           => $row[3],
        //     'zentao_link'       => $row[4],
        //     'priority'          => $row[5],
        //     'status'            => $row[6],
        //     'pending_reason'    => $row[7],
        //     'target_version'    => $row[8],
        //     'impact_analyst'    => $row[9],
        //     'root_cause'        => $row[10],
        //     'work_around'       => $row[11],
        //     'reporter'          => $row[12],
        //     'assignee_to'       => $row[13],
        //     'created'           => $created,
        //     'updated'           => $updated
        // ]);
        // }
    }
}
