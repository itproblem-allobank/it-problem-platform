<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class Data extends Model
{
    protected $table = 'excel_data';

    protected $fillable = ['code_jira', 'environment', 'problem', 'category', 'summary', 'zentao_link', 'priority', 'status', 'pending_reason', 'target_version', 'impact_analyst', 'root_cause', 'work_around', 'reporter', 'assignee_to', 'description', 'frequent', 'complain_info', 'created', 'updated', 'changed_at', 'nickname', 'rca_time', 'closed_time', 'resolved_days', 'rca_days', 'team'];
}
