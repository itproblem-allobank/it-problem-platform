<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class Service extends Model
{
    protected $table = 'service_data';

    protected $fillable = ['issue_type', 'code_jira', 'summary', 'assignee', 'reporter', 'status', 'created', 'updated', 'priority', 'sub_category', 'ticket_number', 'customer_care_category'];
}
