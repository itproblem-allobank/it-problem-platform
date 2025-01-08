<?php

namespace App\Exports;

use App\Models\Data;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithHeadings;
use Illuminate\Support\Facades\DB;

class DataExport implements FromCollection, WithHeadings
{
    /**
     * @return \Illuminate\Support\Collection
     */
    protected $start_date;
    protected $end_date;

    public function __construct($start_date, $end_date)
    {
        $this->start_date = $start_date;
        $this->end_date = $end_date;
    }
    public function collection()
    {
        // dd($this->start_date, $this->end_date);
        $created = Data::select(
            'code_jira',
            'environment',
            'problem',
            'summary',
            'zentao_link',
            'priority',
            'status',
            'pending_reason',
            'target_version',
            'team',
            'impact_analyst',
            'root_cause',
            'work_around',
            'reporter',
            'assignee_to',
            'description',
            'frequent',
            'complain_info',
            'created',
            'updated',
            'changed_at',
            'nickname',
        )->whereBetween(DB::raw('DATE(created)'), [$this->start_date, $this->end_date])->whereIn('status', ['Pending', 'Root Cause Identified'])->get();
        $closed = Data::select(
            'code_jira',
            'environment',
            'problem',
            'summary',
            'zentao_link',
            'priority',
            'status',
            'pending_reason',
            'target_version',
            'team',
            'impact_analyst',
            'root_cause',
            'work_around',
            'reporter',
            'assignee_to',
            'description',
            'frequent',
            'complain_info',
            'created',
            'updated',
            'changed_at',
            'nickname',
        )->whereBetween(DB::raw('DATE(changed_at)'), [$this->start_date, $this->end_date])->where('status', '=', 'closed')->get();

        $createdArray = $created->toArray();
        $closedArray = $closed->toArray();

        $mergedArray = array_merge($createdArray, $closedArray);
        $mergedData = collect($mergedArray);
        // dd($mergedData);
        return $mergedData;
    }

    public function headings(): array
    {
        return [
            'Code Jira',
            'Environment',
            'Problem Category',
            'Summary',
            'Zentao Link',
            'Priority',
            'Status',
            'Pending Reason',
            'Target Version',
            'Escalation Team',
            'Impact Analyst',
            'Root Cause',
            'Workaround',
            'Reporter',
            'Assignee To',
            'Description',
            'Frequent',
            'Complain Info',
            'Created Date',
            'Update Date',
            'Changed Date',
            'Nickname',
        ];
    }
}
