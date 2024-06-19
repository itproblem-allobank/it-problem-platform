<?php

namespace App\Http\Controllers;

use App\Models\Data;
use Illuminate\Support\Facades\DB;
use Carbon\Carbon;

class ChartsController extends Controller
{
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
                ->where('status', 'Closed')
                ->groupBy('problem_category')
                ->get();

            return response()->json([
                'status' => 'success',
                'message' => 'Get all data success',
                'data' => $data,
                'total' => $total,
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
