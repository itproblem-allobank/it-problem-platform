<?php

namespace App\Http\Controllers;

use App\Models\Data;
use Illuminate\Http\Request;
use Maatwebsite\Excel\Facades\Excel;
use Illuminate\Support\Facades\Storage;
use App\Imports\DataImports;
use Barryvdh\DomPDF\Facade\PDF;

class DataController extends Controller
{
  public function __construct()
  {
    $this->middleware('auth');
  }

  public function index()
  {
    $highest = Data::where('Priority', 'Highest')->get()->count();
    $high = Data::where('Priority', 'High')->get()->count();
    $medium = Data::where('Priority', 'Medium')->get()->count();
    $low = Data::where('Priority', 'Low')->get()->count();
    $lowest = Data::where('Priority', 'Lowest')->get()->count();

    // ticket weekly
    $ticket_weekly = Data::whereDate('created', '>', now()->subDays(7))->get();


    return view('data', compact( 'highest', 'high', 'medium', 'low', 'lowest', 'ticket_weekly')); 
  }

  public function getData()
  {
    try {
      // $data = Data::all();
      
    $data = Data::whereDate('created', '>', now()->subDays(7))->get();
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
      return redirect()->route('data')->with(['success' => 'Data Berhasil Diimport!']);
    } else {
      return redirect()->route('data')->with(['error' => 'Data Gagal Diimport!']);
    }
  }

  public  function delete() {
    Data::truncate();
    return redirect()->route('data')->with(['success' => 'Data Berhasil Dihapus!']);
  }

  public function cetak_pdf()
  {
    // $data = Data::all(); 
    $data = Data::whereDate('created', '>', now()->subDays(7))->get();

    $pdf = PDF::loadview('data_pdf', ['data' => $data]);
    // return $pdf->download('laporan-data.pdf');
    return $pdf->stream();
  }
}
