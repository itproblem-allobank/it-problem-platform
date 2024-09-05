<?php

namespace App\Exports;

use App\Models\Data;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithHeadings;

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
            'impact_analyst',
            'root_cause',
            'work_around',
            'reporter',
            'assignee_to',
            'created',
            'updated',
            'changed_at',
            'nickname',
        )->whereBetween('created', [$this->start_date, $this->end_date])->where('status', '=', 'pending')->get();
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
            'impact_analyst',
            'root_cause',
            'work_around',
            'reporter',
            'assignee_to',
            'created',
            'updated',
            'changed_at',
            'nickname',
        )->whereBetween('changed_at', [$this->start_date, $this->end_date])->where('status', '=', 'closed')->get();

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
            'code_jira',
            'environment',
            'problem_category',
            'summary',
            'zentao_link',
            'priority',
            'status',
            'pending_reason',
            'target_version',
            'impact_analyst',
            'root_cause',
            'work_around',
            'reporter',
            'assignee_to',
            'created',
            'updated',
            'changed_at',
            'nickname',
        ];
    }
}
