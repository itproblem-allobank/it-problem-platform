<?php

namespace App\Http\Controllers;

use App\Models\Data;
use DataTables;
use Illuminate\Http\Request;
use Maatwebsite\Excel\Facades\Excel;
use Illuminate\Support\Facades\Storage;
use App\Imports\DataImports;
use Barryvdh\DomPDF\Facade\PDF;

class MonthlyController extends Controller
{
    public function __construct()
    {
        $this->middleware('auth');
    }

    public function index()
    {
        if (request()->ajax()) {
            return DataTables::make(Data::all())->make(true);
        }
        $data = Data::whereDate('created', '>', now()->subDays(30))->get();
        return view('monthly', compact('data'));
    }
}