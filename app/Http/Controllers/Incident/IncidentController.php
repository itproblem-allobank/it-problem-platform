<?php

namespace App\Http\Controllers\Incident;

use App\Http\Controllers\Controller;
use Yajra\DataTables\DataTables;
use Illuminate\Http\Request;
use Maatwebsite\Excel\Facades\Excel;
use Illuminate\Support\Facades\Storage;
use App\Imports\IncidentImports;
use App\Models\Incident;
use Illuminate\Support\Facades\DB;

use Exception;

class IncidentController extends Controller
{
    public function __construct()
    {
        $this->middleware('auth');
    }

    public function index()
    {
        $data = Incident::all();
        if (request()->ajax()) {
            return DataTables::make(Incident::all())->make(true);
        }
        return view('incident/i-index', compact('data'));
    }

    public function import(Request $request)
    {
        Incident::truncate();
        $this->validate($request, [
            'file' => 'required|mimes:csv,xls,xlsx'
        ]);
        $file = $request->file('file');
        $nama_file = $file->hashName();
        $path = $file->storeAs('public/excel/', $nama_file);
        $import = Excel::import(new IncidentImports(), storage_path('app/public/excel/' . $nama_file));
        Storage::delete($path);
        if ($import) {
            return redirect()->route('incident.index')->with(['success' => 'Data Berhasil Diimport!']);
        } else {
            return redirect()->route('incident.index')->with(['error' => 'Data Gagal Diimport!']);
        }
    }

    public function delete()
    {
        Incident::truncate();
        return redirect()->route('incident.index')->with(['success' => 'Data Berhasil Dihapus!']);
    }
}
