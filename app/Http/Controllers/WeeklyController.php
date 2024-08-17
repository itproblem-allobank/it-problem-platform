<?php

namespace App\Http\Controllers;

use App\Models\Data;
use App\Models\Service;
use App\Exports\DataExport;
use Maatwebsite\Excel\Facades\Excel;
use Illuminate\Support\Carbon;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Storage;
use ZipArchive;
use Illuminate\Support\Facades\DB;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\DocumentLayout;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar;
use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Shape\Drawing\File;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Fill;
use Exception;
use Termwind\Components\Raw;

class WeeklyController extends Controller
{
    public function __construct()
    {
        $this->middleware('auth');
    }

    public function index()
    {

        return view('weekly');
    }

    public function exceldownload(Request $request)
    {
        // dd($request);

        $start_date = $request->start_date;
        $end_date = $request->end_date;

        // dd($start_date, $end_date);
        return Excel::download(new DataExport($start_date, $end_date), 'list_problem_weekly.xlsx');
    }

    public function download(Request $request)
    {
        $start_date = $request->start_date;
        $end_date = $request->end_date;

        $objPHPPresentation = new PhpPresentation();
        //Slide 1
        $slide1 = $objPHPPresentation->getActiveSlide();

        $imagePath = storage_path('image/allobank.png');
        $pictureShape = new File();
        $pictureShape->setPath($imagePath);
        $pictureShape->setWidth(350);  // Ubah ukuran gambar sesuai kebutuhan
        $pictureShape->setOffsetX(50); // Posisi horizontal gambar
        $pictureShape->setOffsetY(50); // Posisi vertikal gambar
        $slide1->addShape($pictureShape);

        //Text
        $shape = $slide1->createRichTextShape()
            ->setHeight(50)
            ->setWidth(700)
            ->setOffsetX(50)
            ->setOffsetY(300);
        $textRun = $shape->createTextRun('Report Weekly IT Problem');
        $textRun->getFont()->setBold(true)
            ->setSize(32);

        //Divider
        $lineShape1 = $slide1->createLineShape(50, 370, 1150, 370);
        $lineShape1->getBorder()->setColor(new Color('FF000000'));
        $lineShape1->getBorder()->setLineWidth(2);


        //Text
        $shape = $slide1->createRichTextShape()
            ->setHeight(50)
            ->setWidth(1150)
            ->setOffsetX(50)
            ->setOffsetY(380);
        $textRun1 = $shape->createTextRun('Information Technology Infrastructure & Operations No ');
        $textRun1->getFont()->setBold(true)
            ->setSize(24);
        $textRun2 = $shape->createTextRun('002/ITIO-DOC/XI/2023');
        $textRun2->getFont()->setBold(true)
            ->setSize(24)->setColor(new Color('FFFF0000'));

        //Text
        $shape = $slide1->createRichTextShape()
            ->setHeight(50)
            ->setWidth(280)
            ->setOffsetX(980)
            ->setOffsetY(640);
        $textRun = $shape->createTextRun('PT Allo Bank Indonesia');
        $textRun->getFont()->setSize(20);

        //Slide 2
        $slide2 = $objPHPPresentation->createSlide();


        // Tambahkan teks judul slide
        $shape = $slide2->createRichTextShape()
            ->setHeight(50)
            ->setWidth(400)
            ->setOffsetX(50)
            ->setOffsetY(25);
        $textRun = $shape->createTextRun('Document Control');
        $textRun->getFont()->setBold(true)
            ->setSize(30)->setColor(new Color('FFFFA500'));

        // Add a table for document control details
        $tableShape = $slide2->createTableShape(2);
        $tableShape->setWidth(600);

        // Position the table on the slide
        $tableShape->setOffsetX(50);
        $tableShape->setOffsetY(120);

        // Function to set cell text with font size
        function setCellText($row, $cell, $text, $fontSize = 12)
        {
            $row->setHeight(60);  // Set row height
            $cell->getActiveParagraph()->getAlignment()->setMarginLeft(10);
            $cell->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
            $textRun = $cell->createTextRun($text);
            $textRun->getFont()->setSize($fontSize);
            $textRun->getFont()->setColor(new Color('FF000000')); // Black color
        }

        // Add rows and cells to the table
        $row = $tableShape->createRow();
        $cell = $row->nextCell();
        setCellText($row, $cell, 'Division', 15);
        $cell = $row->nextCell();
        setCellText($row, $cell, 'Information Technology Infrastructure & Operations', 15);

        $row = $tableShape->createRow();
        $cell = $row->nextCell();
        setCellText($row, $cell, 'Title', 15);
        $cell = $row->nextCell();
        setCellText($row, $cell, 'Report Weekly IT Problem', 15);

        $row = $tableShape->createRow();
        $cell = $row->nextCell();
        setCellText($row, $cell, 'Version', 15);
        $cell = $row->nextCell();
        setCellText($row, $cell, Carbon::parse($end_date)->format('F Y'), 15);

        $row = $tableShape->createRow();
        $cell = $row->nextCell();
        setCellText($row, $cell, 'Review date', 15);
        $cell = $row->nextCell();
        setCellText($row, $cell, Carbon::parse($end_date)->format('d F Y'), 15);

        //Text Shape 1
        $textShape1 = $slide2->createRichTextShape();
        $textShape1->setHeight(250);
        $textShape1->setWidth(300);
        $textShape1->setOffsetX(50);
        $textShape1->setOffsetY(420);
        $textShape1->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);

        // Create the text run for the left-aligned text
        $date = Carbon::parse($end_date)->format('d F Y');
        $textRun2 = $textShape1->createTextRun("Jakarta, " . $date . "\n\nDisetujui oleh,\n\n\n\n\n");
        $textRun2->getFont()->setSize(15);
        $textRun2->getFont()->setColor(new Color('FF000000')); // Black color

        // Create the bold text run for "Iswibowo Isakar"
        $boldTextRun = $textShape1->createTextRun("Iswibowo Isakar\n");
        $boldTextRun->getFont()->setSize(15);
        $boldTextRun->getFont()->setColor(new Color('FF000000')); // Black color
        $boldTextRun->getFont()->setBold(true); // Set the text to bold

        // Create the text run for "IT infra Operation"
        $textRun3 = $textShape1->createTextRun("Information Technology\nInfrastructure & Operations");
        $textRun3->getFont()->setSize(15);
        $textRun3->getFont()->setColor(new Color('FF000000')); // Black color

        //Text Shape 2
        $textShape2 = $slide2->createRichTextShape();
        $textShape2->setHeight(250);
        $textShape2->setWidth(300);
        $textShape2->setOffsetX(800);
        $textShape2->setOffsetY(420);
        $textShape2->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);

        // Create the text run for the left-aligned text
        $textRun2 = $textShape2->createTextRun("\n\nDibuat oleh,\n\n\n\n\n");
        $textRun2->getFont()->setSize(15);
        $textRun2->getFont()->setColor(new Color('FF000000')); // Black color

        // Create the bold text run for "Tri Intan Siska Permatasari"
        $boldTextRun = $textShape2->createTextRun("Tri Intan Siska Permatasari\n");
        $boldTextRun->getFont()->setSize(15);
        $boldTextRun->getFont()->setColor(new Color('FF000000')); // Black color
        $boldTextRun->getFont()->setBold(true); // Set the text to bold

        // Create the text run for "IT Problem Lead"
        $textRun3 = $textShape2->createTextRun("IT Problem Lead");
        $textRun3->getFont()->setSize(15);
        $textRun3->getFont()->setColor(new Color('FF000000')); // Black color



        //Slide 3
        $slide3 = $objPHPPresentation->createSlide();
        $backgroundImagePath = storage_path('image/background.png');
        $backgroundImage = new File();
        $backgroundImage->setPath($backgroundImagePath);
        $backgroundImage->setWidth(1280);
        $backgroundImage->setOffsetX(0);
        $backgroundImage->setOffsetY(0);
        $slide3->addShape($backgroundImage);


        $imagePath = storage_path('image/allobank.png');
        $pictureShape = new File();
        $pictureShape->setPath($imagePath);
        $pictureShape->setWidth(200);  // Ubah ukuran gambar sesuai kebutuhan
        $pictureShape->setOffsetX(1050); // Posisi horizontal gambar
        $pictureShape->setOffsetY(20); // Posisi vertikal gambar
        $slide3->addShape($pictureShape);

        $objPHPPresentation->getLayout()->setDocumentLayout(['cx' => 1280, 'cy' => 700], true)
            ->setCX(1280, DocumentLayout::UNIT_PIXEL)
            ->setCY(700, DocumentLayout::UNIT_PIXEL);

        // Tambahkan teks judul slide
        $shape = $slide3->createRichTextShape()
            ->setHeight(50)
            ->setWidth(400)
            ->setOffsetX(25)
            ->setOffsetY(15);
        $textRun = $shape->createTextRun('Report IT Problem');
        $textRun->getFont()->setBold(true)
            ->setSize(30);

        $shape = $slide3->createRichTextShape()
            ->setHeight(25)
            ->setWidth(400)
            ->setOffsetX(25)
            ->setOffsetY(60);
        $startdate = Carbon::parse($start_date)->format('d F Y');
        $enddate = Carbon::parse($end_date)->format('d F Y');
        $textRun = $shape->createTextRun('As of ' . $startdate . ' - ' . $enddate);
        $textRun->getFont()->setSize(14);

        //data container category
        $problem = Data::select('problem', DB::raw('count(*) as count'))->groupBy('problem')->get();
        // dd($problem);
        $total = [];
        foreach ($problem as $key => $value) {
            //declaredata priority
            $high_existing = Data::where(DB::raw('DATE(created)'), '<=', $start_date)->where('problem', '=', $value->problem)->where('status', '=', 'Pending')->where('priority', '=', 'High')->get()->count();
            $medium_existing = Data::where(DB::raw('DATE(created)'), '<=', $start_date)->where('problem', '=', $value->problem)->where('status', '=', 'Pending')->where('priority', '=', 'Medium')->get()->count();
            $low_existing = Data::where(DB::raw('DATE(created)'), '<=', $start_date)->where('problem', '=', $value->problem)->where('status', '=', 'Pending')->where('priority', '=', 'Low')->get()->count();
            $high_now = Data::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('problem', '=', $value->problem)->where('priority', '=', 'High')->get()->count();
            $medium_now = Data::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('problem', '=', $value->problem)->where('priority', '=', 'Medium')->get()->count();
            $low_now = Data::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('problem', '=', $value->problem)->where('priority', '=', 'Low')->get()->count();
            $highclosed = Data::whereBetween('changed_at', [$start_date, $end_date])->where('problem', '=', $value->problem)->where('status', '=', 'Closed')->where('priority', '=', 'High')->get()->count();
            $mediumclosed = Data::whereBetween('changed_at', [$start_date, $end_date])->where('problem', '=', $value->problem)->where('status', '=', 'Closed')->where('priority', '=', 'Medium')->get()->count();
            $lowclosed = Data::whereBetween('changed_at', [$start_date, $end_date])->where('problem', '=', $value->problem)->where('status', '=', 'Closed')->where('priority', '=', 'Low')->get()->count();
            //set total data by priority
            $high_total = $high_existing + $high_now;
            $medium_total = $medium_existing + $medium_now;
            $low_total = $low_existing + $low_now;
            //count data priority
            $countdata = $high_total + $medium_total + $low_total;
            //set color by problem
            $color = '';
            if ($value->problem == 'Core System & Surrounding Apps') {
                $color = 'ff89a64e';
            } else if ($value->problem == 'Ekosistem MPC') {
                $color = 'ff93aacf';
            } else if ($value->problem == 'Loan') {
                $color = 'ffa6a6a6';
            } else if ($value->problem == 'Onboarding') {
                $color = 'fff79646';
            } else if ($value->problem == 'Online Payment') {
                $color = 'ff4f81bd';
            } else if ($value->problem == 'Third Party') {
                $color = 'ffee52e1';
            } else if ($value->problem == 'Transaction') {
                $color = 'ffffc000';
            } else if ($value->problem == 'Wholesale Banking') {
                $color = 'ff8064a2';
            } else {
                $color = 'ffffffff';
            }
            //inject data to array
            $total[] = [
                'problem' => $value->problem,
                'total' => $countdata,
                'high_existing' => $high_existing + $highclosed,
                'medium_existing' => $medium_existing + $mediumclosed,
                'low_existing' => $low_existing + $lowclosed,
                'high' =>  $high_now,
                'medium' => $medium_now,
                'low' => $low_now,
                'highclosed' => $highclosed,
                'mediumclosed' => $mediumclosed,
                'lowclosed' => $lowclosed,
                'color' => $color
            ];
        }

        // dd($total);
        function truncateString($string, $limit = 18)
        {
            if (strlen($string) > $limit) {
                return substr($string, 0, $limit) . '...';
            } else {
                return $string;
            }
        }

        //set tempat
        $offsetx = 25;
        $offsety = 100;
        //loop category data
        foreach ($total as $key => $data) {
            // Tambahkan tabel dengan 4 baris dan 3 kolom
            $tableShape = $slide3->createTableShape(3);
            $tableShape->setHeight(100);
            $tableShape->setWidth(144);
            $tableShape->setOffsetX($offsetx);
            $tableShape->setOffsetY($offsety);

            //row judul
            $rowShape = $tableShape->createRow();
            $rowShape->setHeight(40);
            $cell = $rowShape->nextCell();
            $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color($data["color"]));
            $cell->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            $cell->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
            $cell->setColSpan(3);
            $textRun = $cell->createTextRun($data["total"] . "\n" . truncateString($data["problem"]));
            $textRun->getFont()->setBold(true);
            $textRun->getFont()->setSize(12);

            //row title
            $rowShape = $tableShape->createRow();
            $rowShape->setHeight(20);
            $val = [['status' => 'High', 'color' => 'FFFF0000'], ['status' => 'Med', 'color' => 'FFDCFF00'], ['status' => 'Low', 'color' => 'FF00B050']];
            foreach ($val as $key => $v) {
                $cell = $rowShape->nextCell();
                $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color($v['color']));
                $cell->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                $cell->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
                $textRun = $cell->createTextRun($v['status']);
                $textRun->getFont()->setBold(true);
            }

            $rowShape = $tableShape->createRow();
            $rowShape->setHeight(20);
            $value = [$data['high_existing'], $data['medium_existing'], $data['low_existing']];
            foreach ($value as $key => $v) {
                $cell = $rowShape->nextCell();
                $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color($data["color"]));
                $cell->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                $cell->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
                $cell->createTextRun($v);
            }

            $rowShape = $tableShape->createRow();
            $rowShape->setHeight(20);
            $value = [$data['high'], $data['medium'], $data['low']];
            foreach ($value as $key => $v) {
                $cell = $rowShape->nextCell();
                $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color($data["color"]));
                $cell->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                $cell->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
                $cell->createTextRun($v);
            }

            $rowShape = $tableShape->createRow();
            $rowShape->setHeight(20);
            $value = [$data['highclosed'], $data['mediumclosed'], $data['lowclosed']];
            foreach ($value as $key => $v) {
                $cell = $rowShape->nextCell();
                $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color($data["color"]));
                $cell->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                $cell->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
                $cell->createTextRun($v);
            }

            //set tempat box selanjutnya
            $offsetx = $offsetx + 155;
        }

        //set data chart 1
        $data_chart1 = Data::where(DB::raw('DATE(created)'), '<=', $end_date)->select('problem', DB::raw('count(*) as count'))->groupBy('problem')->get();
        $resultdata_chart1 = [];
        foreach ($data_chart1 as $key => $value) {
            $status_closed = Data::where(DB::raw('DATE(created)'), '<=', $end_date)->where('problem', '=', $value->problem)->where('status', '=', 'Closed')->get()->count();
            $status_pending = Data::where(DB::raw('DATE(created)'), '<=', $end_date)->where('problem', '=', $value->problem)->where('status', '=', 'Pending')->get()->count();
            $closed_thisweek = Data::whereBetween(DB::raw('DATE(changed_at)'), [$start_date, $end_date])->where('problem', '=', $value->problem)->where('status', '=', 'Closed')->get()->count();
            // dd($status_pending, $closed_thisweek, $count_pending);
            $color = '';
            if ($value->problem == 'Core System & Surrounding Apps') {
                $color = 'ff89a64e';
            } else if ($value->problem == 'Ekosistem MPC') {
                $color = 'ff93aacf';
            } else if ($value->problem == 'Loan') {
                $color = 'ffa6a6a6';
            } else if ($value->problem == 'Onboarding') {
                $color = 'fff79646';
            } else if ($value->problem == 'Online Payment') {
                $color = 'ff4f81bd';
            } else if ($value->problem == 'Third Party') {
                $color = 'ffee52e1';
            } else if ($value->problem == 'Transaction') {
                $color = 'ffffc000';
            } else if ($value->problem == 'Wholesale Banking') {
                $color = 'ff8064a2';
            } else {
                $color = 'ffffffff';
            }
            $resultdata_chart1[] =
                [
                    'problem' => $value->problem,
                    'total' => $value->count,
                    'count_closed' => $status_closed,
                    'count_pending' => $status_pending,
                    'color' => $color
                ];
        }

        // Chart 1 Ticket by Category
        $chartShape = $slide3->createChartShape();
        $chartShape->setHeight(250)
            ->setWidth(410)
            ->setOffsetX(25)
            ->setOffsetY(225);
        // Define tipe chart
        $chartType = new Bar();
        $chartShape->getPlotArea()->setType($chartType);
        // Set judul chart
        $chartShape->getTitle()->setText('Ticket by Category');
        // Mendapatkan objek sumbu
        $xAxis = $chartShape->getPlotArea()->getAxisX();
        $yAxis = $chartShape->getPlotArea()->getAxisY();
        // Mengatur judul sumbu menjadi kosong
        $xAxis->setTitle('');
        $yAxis->setTitle('');
        // Chart Bordered
        $chartShape->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $chartShape->getBorder()->setColor(new Color('FF000000')); // Black border
        $chartShape->getBorder()->setLineWidth(1);
        $chartShape->getPlotArea()->getAxisY()->setIsVisible(false);
        $chartShape->getLegend()->getBorder()->setLineStyle(Border::LINE_NONE); // Menghilangkan kotak pada legenda

        // Tambahkan seri data ke chart
        foreach ($resultdata_chart1 as $key => $value) {
            $series = new Series($value['problem'], ['Closed' => $value['count_closed'], 'Pending' => $value['count_pending']]);
            $series->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color($value['color'])); // Blue
            $chartType->addSeries($series);
        }


        //Declare DAY
        $lastweek = [Carbon::parse($start_date)->subDays(7), Carbon::parse($start_date)->subDays(1)];
        $twoweeksago = [Carbon::parse($start_date)->subDays(14), Carbon::parse($start_date)->subDays(8)];
        $threeweeksago = [Carbon::parse($start_date)->subDays(21), Carbon::parse($start_date)->subDays(15)];
        //
        $changed_closed_lweek = Data::whereBetween('changed_at', $lastweek)->where('status', '=', 'Closed')->get();
        $changed_closed_2week = Data::whereBetween('changed_at', $twoweeksago)->where('status', '=', 'Closed')->get();
        $changed_closed_3week = Data::whereBetween('changed_at', $threeweeksago)->where('status', '=', 'Closed')->get();

        $created_closed_lweek = Data::whereBetween(DB::raw('DATE(created)'), $lastweek)->where('status', '=', 'Closed')->get();
        $created_closed_2week = Data::whereBetween(DB::raw('DATE(created)'), $twoweeksago)->where('status', '=', 'Closed')->get();
        $created_closed_3week = Data::whereBetween(DB::raw('DATE(created)'), $threeweeksago)->where('status', '=', 'Closed')->get();

        $created_pending_lweek = Data::whereBetween(DB::raw('DATE(created)'), $lastweek)->where('status', '=', 'Pending')->get()->count();
        $created_pending_2week = Data::whereBetween(DB::raw('DATE(created)'), $twoweeksago)->where('status', '=', 'Pending')->get()->count();
        $created_pending_3week = Data::whereBetween(DB::raw('DATE(created)'), $threeweeksago)->where('status', '=', 'Pending')->get()->count();

        $closedlweek = [];
        $closed2week = [];
        $closed3week = [];

        $temp1week = []; // Array sementara untuk menyimpan elemen unik
        $temp2week = [];
        $temp3week = [];

        function createUniqueKey($value)
        {
            return serialize([
                'problem' => $value->summary,
                'created' => $value->created,
                'status' => $value->status
            ]);
        }
        //gabung lastweek
        foreach ($changed_closed_lweek as $key => $value) {
            $uniqueKey = createUniqueKey($value);
            if (!in_array($uniqueKey, $temp1week)) {
                $temp1week[] = $uniqueKey;
                $closedlweek[] = [
                    'problem' => $value->summary,
                    'created' => $value->created,
                    'status' => $value->status
                ];
            }
        }
        foreach ($created_closed_lweek as $key => $value) {
            $uniqueKey = createUniqueKey($value);
            if (!in_array($uniqueKey, $temp1week)) {
                $temp1week[] = $uniqueKey;
                $closedlweek[] = [
                    'problem' => $value->summary,
                    'created' => $value->created,
                    'status' => $value->status
                ];
            }
        }
        //gabung 2 week
        foreach ($changed_closed_2week as $key => $value) {
            $uniqueKey = createUniqueKey($value);
            if (!in_array($uniqueKey, $temp2week)) {
                $temp2week[] = $uniqueKey;
                $closed2week[] = [
                    'problem' => $value->summary,
                    'created' => $value->created,
                    'status' => $value->status
                ];
            }
        }
        foreach ($created_closed_2week as $key => $value) {
            $uniqueKey = createUniqueKey($value);
            if (!in_array($uniqueKey, $temp2week)) {
                $temp2week[] = $uniqueKey;
                $closed2week[] = [
                    'problem' => $value->summary,
                    'created' => $value->created,
                    'status' => $value->status
                ];
            }
        }
        //gabung 3 week
        foreach ($changed_closed_3week as $key => $value) {
            $uniqueKey = createUniqueKey($value);
            if (!in_array($uniqueKey, $temp3week)) {
                $temp3week[] = $uniqueKey;
                $closed3week[] = [
                    'problem' => $value->summary,
                    'created' => $value->created,
                    'status' => $value->status
                ];
            }
        }
        foreach ($created_closed_3week as $key => $value) {
            $uniqueKey = createUniqueKey($value);
            if (!in_array($uniqueKey, $temp3week)) {
                $temp3week[] = $uniqueKey;
                $closed3week[] = [
                    'problem' => $value->summary,
                    'created' => $value->created,
                    'status' => $value->status
                ];
            }
        }

        // Chart 2, 
        $chartShape = $slide3->createChartShape();
        $chartShape->setHeight(250)
            ->setWidth(410)
            ->setOffsetX(435)
            ->setOffsetY(225);
        // Define tipe chart
        $chartType = new Bar();
        $chartShape->getPlotArea()->setType($chartType);
        // Set judul chart
        $chartShape->getTitle()->setText('Ticket by Last 3 Weeks');
        // Mendapatkan objek sumbu
        $xAxis = $chartShape->getPlotArea()->getAxisX();
        $yAxis = $chartShape->getPlotArea()->getAxisY();
        // Mengatur judul sumbu menjadi kosong
        $xAxis->setTitle('');
        $yAxis->setTitle('');

        // Chart Bordered
        $chartShape->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $chartShape->getBorder()->setColor(new Color('FF000000')); // Black border
        $chartShape->getBorder()->setLineWidth(1);
        $chartShape->getPlotArea()->getAxisY()->setIsVisible(false);
        $chartShape->getLegend()->getBorder()->setLineStyle(Border::LINE_NONE); // Menghilangkan kotak pada legenda

        $series = new Series('Closed', ['3 Weeks Ago' => count($closed3week), '2 Weeks Ago' => count($closed2week), 'Last Weeks' => count($closedlweek)]);
        $series2 = new Series('Pending', ['3 Weeks Ago' => $created_pending_3week, '2 Weeks Ago' => $created_pending_2week, 'Last Weeks' => $created_pending_lweek]);
        $chartType->addSeries($series);
        $chartType->addSeries($series2);

        // Chart 3 Ticket Service Request Nasabah
        // $data_chart3 = Service::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->select('issue_type', DB::raw('count(*) as count'))->groupBy('issue_type')->get();
        $data_chart3 = Service::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('issue_type', '=', '[JSM] Allo Care Service Request')->select('sub_category', DB::raw('count(*) as count'))->groupBy('sub_category')->get();
        $resultdata_chart3 = [];
        foreach ($data_chart3 as $key => $value) {
            $total = Service::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('sub_category', '=', $value->sub_category)->get()->count();
            $status_closed = Service::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('sub_category', '=', $value->sub_category)->where('status', '=', 'Closed')->get()->count();
            $status_declined = Service::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('sub_category', '=', $value->sub_category)->where('status', '=', 'Declined')->get()->count();
            $resultdata_chart3[] =
                [
                    'sub_category' => $value->sub_category,
                    'total' => $total,
                    'count_closed' => $status_closed,
                    'count_declined' => $status_declined,
                ];
        }

        // Set Size Chart
        $chartShape = $slide3->createChartShape();
        $chartShape->setHeight(250)
            ->setWidth(410)
            ->setOffsetX(845)
            ->setOffsetY(225);
        // Define tipe chart
        $chartType = new Bar();
        $chartShape->getPlotArea()->setType($chartType);
        // Set judul chart
        $chartShape->getTitle()->setText('Ticket Service Request Nasabah');
        // Mendapatkan objek sumbu
        $xAxis = $chartShape->getPlotArea()->getAxisX();
        $yAxis = $chartShape->getPlotArea()->getAxisY();
        // Mengatur judul sumbu menjadi kosong
        $xAxis->setTitle('');
        $yAxis->setTitle('');

        // Chart Bordered
        $chartShape->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $chartShape->getBorder()->setColor(new Color('FF000000')); // Black border
        $chartShape->getBorder()->setLineWidth(1);
        $chartShape->getPlotArea()->getAxisY()->setIsVisible(false);
        $chartShape->getLegend()->getBorder()->setLineStyle(Border::LINE_NONE); // Menghilangkan kotak pada legenda

        // Tambahkan seri data ke chart
        foreach ($resultdata_chart3 as $key => $value) {
            $series = new Series($value['sub_category'], ['Total' => $value['total'], 'Closed' => $value['count_closed'],]);
            $chartType->addSeries($series);
        }

        //Chart Customer Care
        $data_chart4 = Service::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('issue_type', '=', '[JSM] Contact Center Request')->select('sub_category', DB::raw('count(*) as count'))->groupBy('sub_category')->get();
        $resultdata_chart4 = [];
        foreach ($data_chart4 as $key => $value) {
            $total = Service::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('sub_category', '=', $value->sub_category)->get()->count();
            $status_closed = Service::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('sub_category', '=', $value->sub_category)->where('status', '=', 'Closed')->get()->count();
            $status_declined = Service::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('sub_category', '=', $value->sub_category)->where('status', '=', 'Declined')->get()->count();
            // $status_userconfirm = Service::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('sub_category', '=', $value->sub_category)->where('status', '=', 'User Confirmation')->get()->count();
            $status_approval = Service::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('sub_category', '=', $value->sub_category)->where('status', 'like', '%' . 'Approval' . '%')->get()->count();
            $status_inprogress = Service::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('sub_category', '=', $value->sub_category)->where('status', '=', 'In Progress')->get()->count();
            $resultdata_chart4[] =
                [
                    'sub_category' => $value->sub_category,
                    'total' => $total,
                    'count_closed' => $status_closed,
                    'count_declined' => $status_declined,
                    // 'count_userconfirm' => $status_userconfirm,
                    'count_approval' => $status_approval,
                    'count_inprogress' => $status_inprogress,
                ];
        }
        // Set Size Chart
        $chartShape = $slide3->createChartShape();
        $chartShape->setHeight(210)
            ->setWidth(410)
            ->setOffsetX(845)
            ->setOffsetY(475);
        // Define tipe chart
        $chartType = new Bar();
        $chartShape->getPlotArea()->setType($chartType);
        // Set judul chart
        $chartShape->getTitle()->setText('Ticket Service Customer Care');

        // Mendapatkan objek sumbu
        $xAxis = $chartShape->getPlotArea()->getAxisX();
        $yAxis = $chartShape->getPlotArea()->getAxisY();

        // Mengatur judul sumbu menjadi kosong
        $xAxis->setTitle('');
        $yAxis->setTitle('');

        // Chart Bordered
        $chartShape->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $chartShape->getBorder()->setColor(new Color('FF000000')); // Black border
        $chartShape->getBorder()->setLineWidth(1);
        $chartShape->getPlotArea()->getAxisY()->setIsVisible(false);
        $chartShape->getLegend()->getBorder()->setLineStyle(Border::LINE_NONE); // Menghilangkan kotak pada legenda

        // Tambahkan seri data ke chart
        foreach ($resultdata_chart4 as $key => $value) {
            $series = new Series($value['sub_category'], ['Total' => $value['total'], 'Closed' => $value['count_closed'], 'Declined' => $value['count_declined'], 'Approval' => $value['count_approval'], 'In Progress' => $value['count_inprogress']]);
            $chartType->addSeries($series);
        }

        // TABLE PROBLEM STATUS
        // Define table properties
        $columns = 3; // Number of columns
        $tableShape = $slide3->createTableShape($columns);
        $tableShape->getBorder()->setLineStyle(Border::LINE_SINGLE);

        // Set the table's position and size
        $tableShape->setHeight(210);
        $tableShape->setWidth(820);
        $tableShape->setOffsetX(25);
        $tableShape->setOffsetY(475);

        // Define the data for the table
        $datacreated = Data::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->select('problem', 'summary', 'status', 'created', 'changed_at')->get();
        $dataclosed = Data::whereBetween('changed_at', [$start_date, $end_date])->where('status', '=', 'Closed')->select('problem', 'summary', 'status', 'created', 'changed_at')->get();
        $tempdata = [
            ['', 'Summary', 'Status', 'Completion Time'],
        ];
        foreach ($datacreated as $key => $value) {
            $status = $value->status . ' - ' . Carbon::parse($value->changed_at)->format('d F Y');
            $tempdata[] = [$value->problem, $value->summary,  $status, '-'];
        }
        foreach ($dataclosed as $key => $value) {
            $status = $value->status . ' - ' . Carbon::parse($value->changed_at)->format('d F Y');
            $created = Carbon::parse($value->created);
            $changed_at = Carbon::parse($value->changed_at);

            // $daysDifference = ($updated_at - $created_at) / (60 * 60 * 24);
            $daysDifference = intval($created->diffInDays($changed_at));
            $daysString = strval($daysDifference) . ' Days';

            $tempdata[] = [$value->problem, $value->summary,  $status, $daysString];
        }
        $tempdata[] = ['', '',  '', ''];
        // dd($tempdata);

        foreach ($tempdata as $rowIndex => $row) {
            $tableRow = $tableShape->createRow();
            $tableRow->setHeight(25); // Set the height of the row
            foreach ($row as $cellIndex => $cellText) {
                if ($cellIndex == 0) {
                    continue; // Lewati kolom yang disembunyikan
                }
                //set status
                $problem = $row[0];
                $status = explode(' - ', $row[2]);
                $firstStatus = $status[0];
                $cell = $tableRow->nextCell();
                $textRun = $cell->createTextRun($cellText);
                $textRun->getFont()->setBold($rowIndex == 0);
                $cell->getFill()->setFillType(Fill::FILL_SOLID);
                $cell->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                $cell->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
                //
                if ($rowIndex == 0) {
                    $cell->getFill()->setStartColor(new Color(Color::COLOR_BLACK));
                    $textRun->getFont()->setColor(new Color(Color::COLOR_WHITE));
                } else {
                    if ($cellIndex == 1) {
                        //coloring by problem
                        if ($problem == 'Core System & Surrounding Apps') {
                            $cell->getFill()->setStartColor(new Color('ff89a64e'));
                        } else if ($problem == 'Ekosistem MPC') {
                            $cell->getFill()->setStartColor(new Color('ff93aacf'));
                        } else if ($problem == 'Loan') {
                            $cell->getFill()->setStartColor(new Color('ffa6a6a6'));
                        } else if ($problem == 'Onboarding') {
                            $cell->getFill()->setStartColor(new Color('fff79646'));
                        } else if ($problem == 'Online Payment') {
                            $cell->getFill()->setStartColor(new Color('ff4f81bd'));
                        } else if ($problem == 'Third Party') {
                            $cell->getFill()->setStartColor(new Color('ffee52e1'));
                        } else if ($problem == 'Transaction') {
                            $cell->getFill()->setStartColor(new Color('ffffc000'));
                        } else if ($problem == 'Wholesale Banking') {
                            $cell->getFill()->setStartColor(new Color('ff8064a2'));
                        } else {
                            $cell->getFill()->setStartColor(new Color('ffffffff'));
                        }
                    } else if ($cellIndex == 2) {
                        //coloring by status
                        if ($firstStatus == 'Pending') {
                            $cell->getFill()->setStartColor(new Color('fff6f610'));
                        } elseif ($firstStatus == 'Closed') {
                            $cell->getFill()->setStartColor(new Color('ff14ca66'));
                        } else {
                            $cell->getFill()->setFillType(Fill::FILL_NONE);
                        }
                    } else {
                        $cell->getFill()->setFillType(Fill::FILL_NONE);
                    }
                }
            }
        }


        //Slide 4
        $slide4 = $objPHPPresentation->createSlide();
        $backgroundImagePath = storage_path('image/background.png');
        $backgroundImage = new File();
        $backgroundImage->setPath($backgroundImagePath);
        $backgroundImage->setWidth(1280);
        $backgroundImage->setOffsetX(0);
        $backgroundImage->setOffsetY(0);
        $slide4->addShape($backgroundImage);


        $imagePath = storage_path('image/allobank.png');
        $pictureShape = new File();
        $pictureShape->setPath($imagePath);
        $pictureShape->setWidth(200);  // Ubah ukuran gambar sesuai kebutuhan
        $pictureShape->setOffsetX(1050); // Posisi horizontal gambar
        $pictureShape->setOffsetY(20); // Posisi vertikal gambar
        $slide4->addShape($pictureShape);

        $objPHPPresentation->getLayout()->setDocumentLayout(['cx' => 1280, 'cy' => 700], true)
            ->setCX(1280, DocumentLayout::UNIT_PIXEL)
            ->setCY(700, DocumentLayout::UNIT_PIXEL);

        // Tambahkan teks judul slide
        $shape = $slide4->createRichTextShape()
            ->setHeight(50)
            ->setWidth(800)
            ->setOffsetX(25)
            ->setOffsetY(15);
        $textRun = $shape->createTextRun('All Ticket Pending - Priority High');
        $textRun->getFont()->setBold(true)
            ->setSize(30);

        $shape = $slide4->createRichTextShape()
            ->setHeight(25)
            ->setWidth(400)
            ->setOffsetX(25)
            ->setOffsetY(60);
        $startdate = Carbon::parse($start_date)->format('d F Y');
        $enddate = Carbon::parse($end_date)->format('d F Y');
        $textRun = $shape->createTextRun('As of ' . $startdate . ' - ' . $enddate);
        $textRun->getFont()->setSize(14);

        //Table Ticket Pending - Priority HIGH

        //Data
        $data_hpriority = Data::where('priority', '=', 'High')->where('status', '=', 'Pending')->orderBy('problem', 'asc')->get();
        $table = [];

        $id = 1;
        foreach ($data_hpriority as $key => $value) {
            $status = $value->status . "\n" . Carbon::parse($value->changed_at)->format('d/m/Y');
            if ($value->pending_reason == null) {
                $pending_reason = 'No Schedule Yet';
            } else {
                $pending_reason = $value->pending_reason;
            }
            if ($value->target_version == null) {
                $target_version = 'No Schedule Yet';
            } else {
                $target_version = $value->target_version;
            }
            $table[] = [$id, $value->problem, $value->summary, $pending_reason, $target_version, $status];
            $id++;
        }

        // dd($table);

        $table1 = array_slice($table, 0, 17);
        $table2 = array_slice($table, 17, 35);

        //Table 1
        $columns = 6;
        $tableShape = $slide4->createTableShape($columns);
        $tableShape->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $tableShape->setHeight(300);
        $tableShape->setWidth(600);
        $tableShape->setOffsetX(25);
        $tableShape->setOffsetY(100);
        $rowHeader = $tableShape->createRow();
        $rowHeader->setHeight(25);
        //header 
        $header = ['No', 'Problem', 'Summary', 'Pending Reason', 'Target Version', 'Status'];
        foreach ($header as $cellIndex => $cellText) {
            $cell = $rowHeader->nextCell();
            if ($cellIndex == 0) {
                $cell->setWidth(20);
            } else if ($cellIndex == 1) {
                $cell->setWidth(100);
            } else if ($cellIndex == 2) {
                $cell->setWidth(200);
            } else if ($cellIndex == 3) {
                $cell->setWidth(95);
            } else if ($cellIndex == 4) {
                $cell->setWidth(95);
            } else if ($cellIndex == 5) {
                $cell->setWidth(90);
            }
            $textRun = $cell->createTextRun($cellText);
            $textRun->getFont()->setBold(true);
            $cell->getFill()->setStartColor(new Color(Color::COLOR_BLACK));
            $textRun->getFont()->setColor(new Color(Color::COLOR_WHITE));
            $cell->getFill()->setFillType(Fill::FILL_SOLID);
            $cell->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            $cell->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
        }
        //data
        foreach ($table1 as $rowIndex => $row) {
            $tableRow = $tableShape->createRow();
            $tableRow->setHeight(25);
            foreach ($row as $cellIndex => $cellText) {
                $cell = $tableRow->nextCell();
                if ($cellIndex == 0) {
                    $cell->setWidth(20);
                } else if ($cellIndex == 1) {
                    $cell->setWidth(100);
                } else if ($cellIndex == 2) {
                    $cell->setWidth(200);
                } else if ($cellIndex == 3) {
                    $cell->setWidth(95);
                } else if ($cellIndex == 4) {
                    $cell->setWidth(95);
                } else if ($cellIndex == 5) {
                    $cell->setWidth(90);
                }
                $textRun = $cell->createTextRun($cellText);
                $cell->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                $cell->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
                //coloring by problem
                if ($row[1] == 'Core System & Surrounding Apps') {
                    $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ff89a64e'));
                } else if ($row[1] == 'Ekosistem MPC') {
                    $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ff93aacf'));
                } else if ($row[1] == 'Loan') {
                    $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ffa6a6a6'));
                } else if ($row[1] == 'Onboarding') {
                    $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('fff79646'));
                } else if ($row[1] == 'Online Payment') {
                    $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ff4f81bd'));
                } else if ($row[1] == 'Third Party') {
                    $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ffee52e1'));
                } else if ($row[1] == 'Transaction') {
                    $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ffffc000'));
                } else if ($row[1] == 'Wholesale Banking') {
                    $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ff8064a2'));
                } else {
                    $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ffffffff'));
                }
            }
        }

        //Table 2
        $columns = 6;
        $tableShape = $slide4->createTableShape($columns);
        $tableShape->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $tableShape->setHeight(300);
        $tableShape->setWidth(600);
        $tableShape->setOffsetX(635);
        $tableShape->setOffsetY(100);
        $rowHeader = $tableShape->createRow();
        $rowHeader->setHeight(25);
        //header 
        $header = ['No', 'Problem', 'Summary', 'Pending Reason', 'Target Version', 'Status'];
        foreach ($header as $cellIndex => $cellText) {
            $cell = $rowHeader->nextCell();
            if ($cellIndex == 0) {
                $cell->setWidth(20);
            } else if ($cellIndex == 1) {
                $cell->setWidth(100);
            } else if ($cellIndex == 2) {
                $cell->setWidth(200);
            } else if ($cellIndex == 3) {
                $cell->setWidth(95);
            } else if ($cellIndex == 4) {
                $cell->setWidth(95);
            } else if ($cellIndex == 5) {
                $cell->setWidth(90);
            }
            $textRun = $cell->createTextRun($cellText);
            $textRun->getFont()->setBold(true);
            $cell->getFill()->setStartColor(new Color(Color::COLOR_BLACK));
            $textRun->getFont()->setColor(new Color(Color::COLOR_WHITE));
            $cell->getFill()->setFillType(Fill::FILL_SOLID);
            $cell->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            $cell->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
        }
        //data
        foreach ($table2 as $rowIndex => $row) {
            $tableRow = $tableShape->createRow();
            $tableRow->setHeight(25);
            foreach ($row as $cellIndex => $cellText) {
                $cell = $tableRow->nextCell();
                if ($cellIndex == 0) {
                    $cell->setWidth(20);
                } else if ($cellIndex == 1) {
                    $cell->setWidth(100);
                } else if ($cellIndex == 2) {
                    $cell->setWidth(200);
                } else if ($cellIndex == 3) {
                    $cell->setWidth(95);
                } else if ($cellIndex == 4) {
                    $cell->setWidth(95);
                } else if ($cellIndex == 5) {
                    $cell->setWidth(90);
                }
                $textRun = $cell->createTextRun($cellText);
                $cell->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                $cell->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
                //coloring by problem
                if ($row[1] == 'Core System & Surrounding Apps') {
                    $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ff89a64e'));
                } else if ($row[1] == 'Ekosistem MPC') {
                    $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ff93aacf'));
                } else if ($row[1] == 'Loan') {
                    $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ffa6a6a6'));
                } else if ($row[1] == 'Onboarding') {
                    $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('fff79646'));
                } else if ($row[1] == 'Online Payment') {
                    $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ff4f81bd'));
                } else if ($row[1] == 'Third Party') {
                    $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ffee52e1'));
                } else if ($row[1] == 'Transaction') {
                    $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ffffc000'));
                } else if ($row[1] == 'Wholesale Banking') {
                    $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ff8064a2'));
                } else {
                    $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ffffffff'));
                }
            }
        }



        //Slide 5
        $slide5 = $objPHPPresentation->createSlide();
        $backgroundImagePath = storage_path('image/background_end.png');
        $backgroundImage = new File();
        $backgroundImage->setPath($backgroundImagePath);
        $backgroundImage->setWidth(1280);
        $backgroundImage->setOffsetX(0);
        $backgroundImage->setOffsetY(0);
        $slide5->addShape($backgroundImage);

        $shape = $slide5->createRichTextShape()
            ->setHeight(100)
            ->setWidth(400)
            ->setOffsetX(120)
            ->setOffsetY(300);
        $textRun = $shape->createTextRun('Thankyou');
        $textRun->getFont()->setBold(true)
            ->setSize(60)->setColor(new Color('FFFFFF'));


        // Simpan presentasi ke dalam file
        $filename = 'Report Weekly IT Problem' . ' - ' . Carbon::parse($start_date)->format('d F Y') . ' s.d ' . Carbon::parse($end_date)->format('d F Y') . '.pptx';
        $savePath = storage_path($filename);
        $writer = IOFactory::createWriter($objPHPPresentation, 'PowerPoint2007');
        $writer->save($savePath);

        // Simpan file Excel sementara
        $excelPath = 'exports/list_problem_weekly.xlsx';
        Excel::store(new DataExport($start_date, $end_date), $excelPath, 'local');

        // 3. Buat file ZIP yang berisi kedua file tersebut
        $zipFilename = 'weekly_report.zip';
        $zipFilePath = storage_path('app/exports/' . $zipFilename);
        $zip = new ZipArchive;
        if ($zip->open($zipFilePath, ZipArchive::CREATE) === TRUE) {
            $zip->addFile(storage_path('app/' . $excelPath), 'list_problem_weekly.xlsx');
            $zip->addFile($savePath, $filename);
            $zip->close();
        }

        // 4. Hapus file sementara setelah digabungkan
        Storage::delete([$excelPath]);
        unlink($savePath); // Menghapus file PPT secara manual karena disimpan di luar storage facade

        // 5. Unduh file ZIP dan hapus setelah diunduh
        return response()->download($zipFilePath)->deleteFileAfterSend(true);
    }
    //
}
