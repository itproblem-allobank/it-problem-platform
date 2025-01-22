<?php

namespace App\Models;


use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class Incident extends Model
{
    use HasFactory;
    protected $table = 'incident_data';

    protected $fillable = [
        'no_jira',
        'summary',
        'status_ticket',
        'disruption',
        'rootcause',
        'mitigation',
        'created_time',
        'resolved_time'
    ];
}