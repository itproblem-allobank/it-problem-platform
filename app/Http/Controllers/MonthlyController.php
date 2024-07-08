<?php

namespace App\Http\Controllers;

use App\Models\Data;
use Yajra\DataTables\DataTables;
use Illuminate\Http\Request;
use Maatwebsite\Excel\Facades\Excel;
use Illuminate\Support\Facades\Storage;
use App\Imports\DataImports;
use Illuminate\Support\Facades\DB;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Font;
use PhpOffice\PhpPresentation\Shape\Chart\Chart;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar3D;
use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Shape\Chart\Legend;

use Exception;

class MonthlyController extends Controller
{
    public function __construct()
    {
        $this->middleware('auth');
    }

    public function index()
    {
        $problem = Data::select('problem', DB::raw('count(*) as count'))->groupBy('problem')->get();
        $total = [];
        foreach ($problem as $key => $value) {
            $highest = Data::where('problem', '=', $value->problem)->where('priority', '=', 'Highest')->get()->count();
            $high = Data::where('problem', '=', $value->problem)->where('priority', '=', 'High')->get()->count();
            $medium = Data::where('problem', '=', $value->problem)->where('priority', '=', 'Medium')->get()->count();
            $low = Data::where('problem', '=', $value->problem)->where('priority', '=', 'Low')->get()->count();
            $lowest = Data::where('problem', '=', $value->problem)->where('priority', '=', 'Lowest')->get()->count();
            $highestmonthly = Data::where('problem', '=', $value->problem)->where('priority', '=', 'Highest')->where('created' ,'>', now()->subDays(30))->get()->count();
            $highmonthly = Data::where('problem', '=', $value->problem)->where('priority', '=', 'High')->where('created' ,'>', now()->subDays(30))->get()->count();
            $mediummonthly = Data::where('problem', '=', $value->problem)->where('priority', '=', 'Medium')->where('created' ,'>', now()->subDays(30))->get()->count();
            $lowmonthly = Data::where('problem', '=', $value->problem)->where('priority', '=', 'Low')->where('created' ,'>', now()->subDays(30))->get()->count();
            $lowestmonthly = Data::where('problem', '=', $value->problem)->where('priority', '=', 'Lowest')->where('created' ,'>', now()->subDays(30))->get()->count();
            $total[] = [
                'problem' => $value->problem,
                'total' => $value->count,
                'high' => $highest + $high,
                'medium' => $medium,
                'low' => $low + $lowest,
                'highmonthly' => $highestmonthly + $highmonthly,
                'mediummonthly' => $mediummonthly,
                'lowmonthly' => $lowmonthly + $lowestmonthly
            ];
        }
        $data = Data::all();
        if (request()->ajax()) {
            return DataTables::make(Data::all())->make(true);
        }
        return view('monthly', compact('total','data'));
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

    public function export(Request $request)
    {
        // Pisahkan data base64 dan tipe gambar
        list($type, $chart_total) = explode(';', $request->total);
        list($type, $chart_pending) = explode(';', $request->pending);
        list($type, $chart_closed) = explode(';', $request->closed);

        list(, $chart_total) = explode(',', $chart_total);
        list(, $chart_pending) = explode(',', $chart_pending);
        list(, $chart_closed) = explode(',', $chart_closed);

        $chart_total = base64_decode($chart_total);
        $chart_pending = base64_decode($chart_pending);
        $chart_closed = base64_decode($chart_closed);

        // Tentukan nama file gambar sementara
        $image1 = 'image1.png';
        $image2 = 'image2.png';
        $image3 = 'image3.png';

        // Simpan gambar ke file sementara
        file_put_contents($image1, $chart_total);
        file_put_contents($image2, $chart_pending);
        file_put_contents($image3, $chart_closed);

        //Susunan Tampilan Word
        $phpWord = new \PhpOffice\PhpWord\PhpWord();
        $section = $phpWord->addSection();
        $section->addText('Report IT Problem');
        $section->addImage($image1, array('width' => 200));
        $section->addImage($image2, array('width' => 200));
        $section->addImage($image3, array('width' => 200));

        $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
        try {

            $objWriter->save(storage_path('charts.docx'));
        } catch (Exception $e) {
        }
        return response()->download(storage_path('charts.docx'));
        // Hapus file gambar sementara
        unlink($image1, $image2, $image3);
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


    public function generateppt()
    {
        // Buat instance baru dari PhpPresentation
        $objPHPPresentation = new PhpPresentation();

        // Tambahkan slide baru
        $currentSlide = $objPHPPresentation->getActiveSlide();

        // Tambahkan teks judul slide
        $shape = $currentSlide->createRichTextShape()
            ->setHeight(100)
            ->setWidth(600)
            ->setOffsetX(170)
            ->setOffsetY(50);

        $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        $textRun = $shape->createTextRun('Bar Chart Example');
        $textRun->getFont()->setBold(true)
            ->setSize(24)
            ->setColor(new Color('FFE06B20'));

        // Data untuk chart
        $categories = ['Category 1', 'Category 2', 'Category 3'];
        $values1 = [10, 20, 30];
        $values2 = [15, 25, 35];

        // Tambahkan chart bar ke slide
        $chartShape = $currentSlide->createChartShape();
        $chartShape->setHeight(400)
            ->setWidth(600)
            ->setOffsetX(170)
            ->setOffsetY(150);

        // Set judul chart
        $chartShape->getTitle()->setText('Sales Data');

        // Define tipe chart
        $chartType = new Bar3D();
        $chartShape->getPlotArea()->setType($chartType);

        // Tambahkan seri data ke chart
        $series1 = new Series('Series 1', $categories, $values1);
        $series2 = new Series('Series 2', $categories, $values2);
        $chartType->addSeries($series1);
        $chartType->addSeries($series2);

        // Simpan presentasi ke dalam file
        $filename = 'presentation_' . time() . '.pptx';
        $savePath = storage_path($filename);

        $writer = IOFactory::createWriter($objPHPPresentation, 'PowerPoint2007');
        $writer->save($savePath);

        // Return file sebagai response download
        return response()->download($savePath)->deleteFileAfterSend(true);
    }
}
