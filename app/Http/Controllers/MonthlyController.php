<?php

namespace App\Http\Controllers;

use App\Models\Data;
use DataTables;
use Illuminate\Http\Request;
use Maatwebsite\Excel\Facades\Excel;
use Illuminate\Support\Facades\Storage;
use App\Imports\DataImports;
use Illuminate\Support\Facades\DB;
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

    public function import(Request $request)
    {
        Data::truncate();
        $this->validate($request, [
            'file' => 'required|mimes:csv,xls,xlsx'
        ]);
        $file = $request->file('file');
        // membuat nama file unik
        $nama_file = $file->hashName();
        //temporary file
        $path = $file->storeAs('public/excel/', $nama_file);
        // import data
        $import = Excel::import(new DataImports(), storage_path('app/public/excel/' . $nama_file));
        //remove from server
        Storage::delete($path);
        if ($import) {
            return redirect()->route('monthly')->with(['success' => 'Data Berhasil Diimport!']);
        } else {
            return redirect()->route('monthly')->with(['error' => 'Data Gagal Diimport!']);
        }
    }

    public function delete()
    {
        Data::truncate();
        return redirect()->route('monthly')->with(['success' => 'Data Berhasil Dihapus!']);
    }

    public function chart()
    {
        try {
            $data = Data::all();
            $total = Data::select('problem', DB::raw('count(*) as total'))->groupBy('problem')->get();
            return response()->json([
                'status' => 'success',
                'message' => 'Get all data success',
                'data' => $data,
                'total' => $total
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