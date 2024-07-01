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
                    'high' => $highest+$high,
                    'medium' => $medium,
                    'low' => $low+$lowest,
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
                    'high' => $highest+$high,
                    'medium' => $medium,
                    'low' => $low+$lowest,
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
                    'high' => $highest+$high,
                    'medium' => $medium,
                    'low' => $low+$lowest,
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