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
        
        $data = Data::all();
        if (request()->ajax()) {
            return DataTables::make(Data::all())->make(true);
        }
        return view('monthly', compact('data'));
    }
}