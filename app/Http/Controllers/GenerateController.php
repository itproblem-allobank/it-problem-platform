<?php

namespace App\Http\Controllers;

use App\Models\Data;
use Yajra\DataTables\DataTables;
use Illuminate\Http\Request;
use Maatwebsite\Excel\Facades\Excel;
use Illuminate\Support\Facades\Storage;
use App\Imports\DataImports;
use Illuminate\Support\Facades\DB;

use Exception;

class GenerateController extends Controller
{
    public function __construct()
    {
        $this->middleware('auth');
    }

    public function index() {

        return view('generate_ppt');
    }
}