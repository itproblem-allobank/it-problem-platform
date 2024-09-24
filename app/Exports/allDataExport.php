<?php

namespace App\Exports;

use App\Models\Data;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithHeadings;
use Illuminate\Support\Facades\DB;
use Illuminate\Support\Carbon;

class allDataExport implements FromCollection, WithHeadings
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

        
        $lastweek = [Carbon::parse($this->start_date)->subDays(7), Carbon::parse($this->start_date)->subDays(1)];
        $lastweek_created = Data::select(
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
            'description',
            'frequent',
            'complain_info',
            'created',
            'updated',
            'changed_at',
            'nickname',
        )->whereBetween(DB::raw('DATE(created)'), $lastweek)->whereIn('status', ['Pending', 'Root Cause Identified'])->get();

        $lastweek_closed = Data::select(
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
            'description',
            'frequent',
            'complain_info',
            'created',
            'updated',
            'changed_at',
            'nickname',
        )->whereBetween(DB::raw('DATE(changed_at)'), $lastweek)->where('status', '=', 'closed')->get();


        $createdArray = $created->toArray();
        $closedArray = $closed->toArray();

        $lastweek_createdArray = $lastweek_created->toArray();
        $lastweek_closedArray = $lastweek_closed->toArray();

        $mergedArray = array_merge($createdArray, $closedArray, $lastweek_createdArray, $lastweek_closedArray);
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
            'description',
            'frequent',
            'complain_info',
            'created',
            'updated',
            'changed_at',
            'nickname',
        ];
    }
}
