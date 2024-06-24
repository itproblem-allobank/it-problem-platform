<?php

namespace App\Http\Controllers;

use App\Models\Data;
use Illuminate\Support\Facades\DB;
use Illuminate\Http\Request;
use Barryvdh\DomPDF\Facade\PDF;
use Carbon\Carbon;

class ChartsController extends Controller
{

    public function index() {
        $table = Data::whereDate('created', '>', now()->subDays(7))->get();
        $highest = Data::where('Priority', 'Highest')->get()->count();
        $high = Data::where('Priority', 'High')->get()->count();
        $medium = Data::where('Priority', 'Medium')->get()->count();
        $low = Data::where('Priority', 'Low')->get()->count();
        $lowest = Data::where('Priority', 'Lowest')->get()->count();
    
        // ticket weekly
        $ticket_weekly = Data::whereDate('created', '>', now()->subDays(7))->get();
    
    
        return view('charts', compact( 'highest', 'high', 'medium', 'low', 'lowest', 'ticket_weekly', 'table')); 
    }

    public function print(Request $request)
    {
        // dd($request->all());
        $today = Carbon::today();
        // dd($today);
        $table = Data::whereDate('created', '>', now()->subDays(7))->get();
        $weekly = $request->weekly;
        $total = $request->total;
        $priority = $request->priority;
        $pdf = PDF::loadView('temp', compact('weekly', 'total', 'priority', 'table', 'today'));
        return $pdf->download('charts.pdf');
    }

    public function weekly()
    {
        try {
            $data = Data::whereDate('created', '>', now()->subDays(7))
                ->select('problem_category', DB::raw('count(*) as count'))
                ->groupBy('problem_category')
                ->get();
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

    public function total()
    {
        try {
            $data = Data::all();
            $total = Data::select('problem_category', DB::raw('count(*) as count'))
                ->groupBy('problem_category')
                ->get();

            $closed = Data::select('problem_category', DB::raw('count(*) as count'))
                ->groupBy('problem_category')
                ->get();

            $pending = Data::select('problem_category', DB::raw('count(*) as count'))
                ->groupBy('problem_category')
                ->where('status', 'Pending')
                ->get();

            return response()->json([
                'status' => 'success',
                'message' => 'Get all data success',
                'data' => $data,
                'total' => $total,
                'closed' => $closed,
                'pending' => $pending,
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
