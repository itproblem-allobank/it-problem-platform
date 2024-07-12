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

class JiraController extends Controller
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
        return view('ticket_jira', compact('data'));
    }

    public function import(Request $request)
    {
        Data::truncate();
        $this->validate($request, [
            'file' => 'required|mimes:csv,xls,xlsx'
        ]);
        $file = $request->file('file');
        $nama_file = $file->hashName();
        $path = $file->storeAs('public/excel/', $nama_file);
        $import = Excel::import(new DataImports(), storage_path('app/public/excel/' . $nama_file));
        Storage::delete($path);
        if ($import) {
            return redirect()->route('jira.index')->with(['success' => 'Data Berhasil Diimport!']);
        } else {
            return redirect()->route('jira.index')->with(['error' => 'Data Gagal Diimport!']);
        }
    }

    public function delete()
    {
        Data::truncate();
        return redirect()->route('jira.index')->with(['success' => 'Data Berhasil Dihapus!']);
    }
}
