<?php

namespace App\Http\Controllers;

use App\Models\Service;
use Illuminate\Http\Request;
use Maatwebsite\Excel\Facades\Excel;
use Illuminate\Support\Facades\Storage;
use App\Imports\ServiceImports;
use Yajra\DataTables\DataTables;
use Illuminate\Support\Facades\DB;

use Exception;

class ServiceController extends Controller
{
    public function __construct()
    {
        $this->middleware('auth');
    }

    public function index()
    {
        $data = Service::all();
        if (request()->ajax()) {
            return DataTables::make(Service::all())->make(true);
        }

        return view('service_request', compact('data'));
    }

    public function import(Request $request)
    {
        Service::truncate();
        $this->validate($request, [
            'file' => 'required|mimes:csv,xls,xlsx'
        ]);
        $file = $request->file('file');
        $nama_file = $file->hashName();
        $path = $file->storeAs('public/excel/', $nama_file);
        $import = Excel::import(new ServiceImports(), storage_path('app/public/excel/' . $nama_file));
        Storage::delete($path);
        if ($import) {
            return redirect()->route('service.index')->with(['success' => 'Data Berhasil Diimport!']);
        } else {
            return redirect()->route('service.index')->with(['error' => 'Data Gagal Diimport!']);
        }
    }
}
