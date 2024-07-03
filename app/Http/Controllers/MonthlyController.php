<?php

namespace App\Http\Controllers;

use App\Models\Data;
use Yajra\DataTables\DataTables;
use Illuminate\Http\Request;
use Maatwebsite\Excel\Facades\Excel;
use Illuminate\Support\Facades\Storage;
use App\Imports\DataImports;
use Illuminate\Support\Facades\DB;
use Barryvdh\DomPDF\Facade\PDF;
use Exception;

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

    public function export()
    {
    $phpWord = new \PhpOffice\PhpWord\PhpWord();
        $section = $phpWord->addSection();
        $description = "Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodoconsequat. Duis aute irure dolor in reprehenderit in voluptate velit essecillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat nonproident, sunt in culpa qui officia deserunt mollit anim id est laborum.";
        $section->addText('Report IT Problem');
        $section->addText($description);
        // $section->addImage($imagedecode);


        $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
        try {

            $objWriter->save(storage_path('helloWorld.docx'));
        } catch (Exception $e) {
        }
        return response()->download(storage_path('helloWorld.docx'));
    }

    public function delete()
    {
        Data::truncate();
        return redirect()->route('monthly')->with(['success' => 'Data Berhasil Dihapus!']);
    }

    public function chart()
    {
        try {
            $problem = Data::select('problem', DB::raw('count(*) as count'))->groupBy('problem')->get();
            $total = [];
            foreach ($problem as $key => $value) {
                $highest = Data::where('problem', '=', $value->problem)->where('priority', '=', 'Highest')->get()->count();
                $high = Data::where('problem', '=', $value->problem)->where('priority', '=', 'High')->get()->count();
                $medium = Data::where('problem', '=', $value->problem)->where('priority', '=', 'Medium')->get()->count();
                $low = Data::where('problem', '=', $value->problem)->where('priority', '=', 'Low')->get()->count();
                $lowest = Data::where('problem', '=', $value->problem)->where('priority', '=', 'Lowest')->get()->count();
                $total[] = [
                    'problem' => $value->problem,
                    'total' => $value->count,
                    'high' => $highest + $high,
                    'medium' => $medium,
                    'low' => $low + $lowest,
                ];
            }
            $problem_pending = Data::where('status', '=', 'Pending')->select('problem', DB::raw('count(*) as count'))->groupBy('problem')->get();
            $pending = [];
            foreach ($problem_pending as $key => $value) {
                $highest = Data::where('problem', '=', $value->problem)->where('status', '=', 'Pending')->where('priority', '=', 'Highest')->get()->count();
                $high = Data::where('problem', '=', $value->problem)->where('status', '=', 'Pending')->where('priority', '=', 'High')->get()->count();
                $medium = Data::where('problem', '=', $value->problem)->where('status', '=', 'Pending')->where('priority', '=', 'Medium')->get()->count();
                $low = Data::where('problem', '=', $value->problem)->where('status', '=', 'Pending')->where('priority', '=', 'Low')->get()->count();
                $lowest = Data::where('problem', '=', $value->problem)->where('status', '=', 'Pending')->where('priority', '=', 'Lowest')->get()->count();
                $pending[] = [
                    'problem' => $value->problem,
                    'total' => $value->count,
                    'high' => $highest + $high,
                    'medium' => $medium,
                    'low' => $low + $lowest,
                ];
            }
            $problem_closed = Data::where('status', '=', 'Closed')->select('problem', DB::raw('count(*) as count'))->groupBy('problem')->get();
            $closed = [];
            foreach ($problem_closed as $key => $value) {
                $highest = Data::where('problem', '=', $value->problem)->where('status', '=', 'Closed')->where('priority', '=', 'Highest')->get()->count();
                $high = Data::where('problem', '=', $value->problem)->where('status', '=', 'Closed')->where('priority', '=', 'High')->get()->count();
                $medium = Data::where('problem', '=', $value->problem)->where('status', '=', 'Closed')->where('priority', '=', 'Medium')->get()->count();
                $low = Data::where('problem', '=', $value->problem)->where('status', '=', 'Closed')->where('priority', '=', 'Low')->get()->count();
                $lowest = Data::where('problem', '=', $value->problem)->where('status', '=', 'Closed')->where('priority', '=', 'Lowest')->get()->count();
                $closed[] = [
                    'problem' => $value->problem,
                    'total' => $value->count,
                    'high' => $highest + $high,
                    'medium' => $medium,
                    'low' => $low + $lowest,
                ];
            }

            return response()->json([
                'status' => 'success',
                'message' => 'Get all data success',
                'total' => $total,
                'pending' => $pending,
                'closed' => $closed,
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
