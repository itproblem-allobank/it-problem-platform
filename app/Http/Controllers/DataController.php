<?php

namespace App\Http\Controllers;

use App\Models\Data;
use Yajra\DataTables\Facades\DataTables;

class DataController extends Controller
{
    public function __construct()
    {
        $this->middleware('auth');
    }

    public function index()
    {
        return view('data');
    }

    public function getData()
    {
        try {
            $data = Data::all();

            return response()->json([
              'status' => 'success',
              'message' => 'Get all data success',
              'data' => $data,
            ]);
          } catch (\Exception $e) {
            return response()->json([
              'status' => 'error',
              'message' => 'Get all data failed',
              'error' => $e->getMessage(),
            ]);
          }
    }

}
