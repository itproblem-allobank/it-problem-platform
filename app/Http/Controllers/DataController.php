<?php

namespace App\Http\Controllers;

use App\Models\Data;
use Illuminate\Http\Request;
use Maatwebsite\Excel\Facades\Excel;
use Illuminate\Support\Facades\Storage;
use App\Imports\DataImports;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Style\Alignment;
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

  public function cetak_pdf()
  {
    // $data = Data::all(); 
    $data = Data::whereDate('created', '>', now()->subDays(7))->get();

    $pdf = PDF::loadview('data_pdf', ['data' => $data]);
    // return $pdf->download('laporan-data.pdf');
    return $pdf->stream();
  }


  public function export()
  {
    $presentation = new PhpPresentation();
    // Retrieve data from database, example:
    $data = Data::all();
    // Create a slide
    $slide = $presentation->getActiveSlide();
    // Add data from database to slide
    foreach ($data as $item) {
      $shape = $slide->createRichTextShape();
      $shape->setHeight(300);
      $shape->setWidth(600);
      $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
      $shape->getActiveParagraph()->createTextRun($item->field1 . ' - ' . $item->field2);
    }
    // Save PowerPoint file
    $writer = new \PhpOffice\PhpPresentation\Writer\PowerPoint2007($presentation);
    // $writer->save(storage_path('app/public/powerpoint.pptx'));
    $writer->save(storage_path('app/public/powerpoint.pptx'));
    return 'PowerPoint file has been generated!';
  }
}
