<?php

namespace App\Http\Controllers;

use App\Models\Data;
use Illuminate\Http\Request;
use Maatwebsite\Excel\Facades\Excel;
use Illuminate\Support\Facades\Storage;
use App\Imports\DataImports;

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

    public function import(Request $request)
    {
        $this->validate($request, [
            'file' => 'required|mimes:csv,xls,xlsx'
        ]);

        $file = $request->file('file');

        // membuat nama file unik
        $nama_file = $file->hashName();

        //temporary file
        $path = $file->storeAs('public/excel/',$nama_file);

        // import data
        $import = Excel::import(new DataImports(), storage_path('app/public/excel/'.$nama_file));

        //remove from server
        Storage::delete($path);

        if($import) {
            //redirect
            return redirect()->route('data')->with(['success' => 'Data Berhasil Diimport!']);
        } else {
            //redirect
            return redirect()->route('data')->with(['error' => 'Data Gagal Diimport!']);
        }
    }

}
