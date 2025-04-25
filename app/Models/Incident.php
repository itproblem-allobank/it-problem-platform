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
        'incident',
        'category',
        'status_ticket',
        'priority',
        'rootcause',
        'mitigation',
        'created_time',
        'resolved_time'
    ];
}