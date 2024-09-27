<?php

namespace App\Http\Controllers;

use App\Models\Data;
use App\Models\Service;
use App\Exports\DataExport;
use App\Exports\allDataExport;
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
            ->setWidth(1000)
            ->setOffsetX(25)
            ->setOffsetY(15);
        $textRun = $shape->createTextRun('IT Problem - Status Rootcause Identified & Pending');
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
            //count ticket existing
            $high_existing = Data::where(DB::raw('DATE(created)'), '<', $start_date)
                ->where('problem', '=', $value->problem)
                ->where('priority', '=', 'High')
                ->whereIn('status', ['Root Cause Identified', 'Pending'])
                ->count();
            $medium_existing = Data::where(DB::raw('DATE(created)'), '<', $start_date)
                ->where('problem', '=', $value->problem)
                ->where('priority', '=', 'Medium')
                ->whereIn('status', ['Root Cause Identified', 'Pending'])
                ->count();
            $low_existing = Data::where(DB::raw('DATE(created)'), '<', $start_date)
                ->where('problem', '=', $value->problem)
                ->where('priority', '=', 'Low')
                ->whereIn('status', ['Root Cause Identified', 'Pending'])
                ->count();
            //count ticket created
            $high_now = Data::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])
                ->where('problem', '=', $value->problem)
                ->where('priority', '=', 'High')
                ->count();
            $medium_now = Data::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])
                ->where('problem', '=', $value->problem)
                ->where('priority', '=', 'Medium')
                ->count();
            $low_now = Data::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])
                ->where('problem', '=', $value->problem)
                ->where('priority', '=', 'Low')
                ->count();
            //count ticket closed
            $highclosed = Data::whereBetween(DB::raw('DATE(changed_at)'), [$start_date, $end_date])
                ->where('problem', '=', $value->problem)
                ->where('status', '=', 'Closed')
                ->where('priority', '=', 'High')
                ->count();
            $mediumclosed = Data::whereBetween(DB::raw('DATE(changed_at)'), [$start_date, $end_date])
                ->where('problem', '=', $value->problem)
                ->where('status', '=', 'Closed')
                ->where('priority', '=', 'Medium')
                ->count();
            $lowclosed = Data::whereBetween(DB::raw('DATE(changed_at)'), [$start_date, $end_date])
                ->where('problem', '=', $value->problem)
                ->where('status', '=', 'Closed')
                ->where('priority', '=', 'Low')
                ->count();

            //count total created
            $totalcreated = $high_now + $medium_now + $low_now;

            $highexist = $high_existing + $highclosed;
            $mediumexist = $medium_existing + $mediumclosed;
            $lowexist = $low_existing + $lowclosed;

            //count data priority
            $countdata = $highexist + $mediumexist + $lowexist + $totalcreated - $highclosed - $mediumclosed - $lowclosed;

            //set color by problem
            $color = '';
            if ($value->problem == 'Core & Surrounding') {
                $color = 'ff89a64e';
            } else if ($value->problem == 'Ekosistem MPC') {
                $color = 'ff00b0f0';
            } else if ($value->problem == 'Loan') {
                $color = 'ffa6a6a6';
            } else if ($value->problem == 'Onboarding') {
                $color = 'ff81ff63';
            } else if ($value->problem == 'Online Payment') {
                $color = 'ff09b1a7';
            } else if ($value->problem == 'Switching & 3rdparty') {
                $color = 'ffee52e1';
            } else if ($value->problem == 'Transaction') {
                $color = 'ff8380ee';
            } else if ($value->problem == 'Wholesale Banking') {
                $color = 'ff8064a2';
            } else {
                $color = 'ffffffff';
            }

            //inject data to array
            $total[] = [
                'problem' => $value->problem,
                'total' => $countdata,
                'high_existing' => $highexist,
                'medium_existing' => $mediumexist,
                'low_existing' => $lowexist,
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

        function truncateString($string, $limit = 20)
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
            $val = [['status' => 'High', 'color' => 'FFFF0000'], ['status' => 'Med', 'color' => 'fffeb909'], ['status' => 'Low', 'color' => 'fffffe00']];
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

        $totalexisting = 0;
        $totalcreated = 0;
        $totalclosed = 0;
        $totalhigh = 0;
        $totalmed = 0;
        $totallow = 0;
        foreach ($total as $key => $value) {
            $totalexisting += $value["high_existing"] + $value["medium_existing"] + $value["low_existing"];
            $totalcreated += $value["high"] + $value["medium"] + $value["low"];
            $totalclosed += $value["highclosed"] + $value["mediumclosed"] + $value["lowclosed"];
            $totalhigh += $value["high_existing"] + $value["high"]  - $value["highclosed"];
            $totalmed += $value["medium_existing"] + $value["medium"] - $value["mediumclosed"];
            $totallow += $value["low_existing"] + $value["low"] - $value["lowclosed"];
        }
        // dd($totalcreated, $totalclosed, $totalexisting, $totalhigh);

        // Total High, Med & Low
        // $shape = $slide3->createRichTextShape();
        // $shape->setHeight(25)
        //     ->setWidth(500)
        //     ->setOffsetX(850)
        //     ->setOffsetY(65);
        // $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
        // $textRun = $shape->createTextRun('Total High: ' . $totalhigh . ' | Total Medium: ' . $totalmed . ' | Total Low: ' . $totallow);
        // $textRun->getFont()->setBold(true)
        //     ->setSize(14)
        //     ->setColor(new Color(Color::COLOR_BLACK));

        // Total HIGH, MED, LOW
        $tableShape = $slide3->createTableShape(3);
        $tableShape->setHeight(100);
        $tableShape->setWidth(400);
        $tableShape->setOffsetX(855);
        $tableShape->setOffsetY(75);

        //row title
        $rowShape = $tableShape->createRow();
        $rowShape->setHeight(20);
        $val = [['status' => 'Total High', 'color' => 'FFFF0000', 'value' => $totalhigh], ['status' => 'Total Medium', 'color' => 'fffeb909', 'value' => $totalmed], ['status' => 'Total Low', 'color' => 'fffffe00', 'value' => $totallow]];
        foreach ($val as $key => $v) {
            $cell = $rowShape->nextCell();
            $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color($v['color']));
            $cell->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            $cell->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
            $textRun = $cell->createTextRun($v['status'] . ' : ' . $v['value']);
            $textRun->getFont()->setBold(true)
                ->setSize(12);
        }


        // Icon +
        $shape = $slide3->createRichTextShape();
        $shape->setHeight(25)
            ->setWidth(40)
            ->setOffsetX(-5)
            ->setOffsetY(175);
        $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $textRun = $shape->createTextRun('+');
        $textRun->getFont()->setBold(true)
            ->setSize(16)
            ->setColor(new Color(Color::COLOR_BLACK));

        // Icon -
        $shape = $slide3->createRichTextShape();
        $shape->setHeight(25)
            ->setWidth(40)
            ->setOffsetX(-5)
            ->setOffsetY(195);
        $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $textRun = $shape->createTextRun('-');
        $textRun->getFont()->setBold(true)
            ->setSize(16)
            ->setColor(new Color(Color::COLOR_BLACK));

        // Total Existing
        $shape = $slide3->createRichTextShape();
        $shape->setHeight(25)
            ->setWidth(40)
            ->setOffsetX(1247)
            ->setOffsetY(155);
        $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $textRun = $shape->createTextRun($totalexisting);
        $textRun->getFont()->setBold(true)
            ->setSize(12)
            ->setColor(new Color(Color::COLOR_BLACK));

        //Total Created
        $shape = $slide3->createRichTextShape();
        $shape->setHeight(25)
            ->setWidth(40)
            ->setOffsetX(1247)
            ->setOffsetY(175);
        $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $textRun = $shape->createTextRun($totalcreated);
        $textRun->getFont()->setBold(true)
            ->setSize(12)
            ->setColor(new Color(Color::COLOR_BLACK));

        //Total Closed
        $shape = $slide3->createRichTextShape();
        $shape->setHeight(25)
            ->setWidth(40)
            ->setOffsetX(1247)
            ->setOffsetY(195);
        $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $textRun = $shape->createTextRun($totalclosed);
        $textRun->getFont()->setBold(true)
            ->setSize(12)
            ->setColor(new Color(Color::COLOR_BLACK));




        // -------------------- CHART 1 ---------------------
        $data_chart1 = Data::where(DB::raw('DATE(created)'), '<=', $end_date)->select('problem', DB::raw('count(*) as count'))->groupBy('problem')->get();
        $resultdata_chart1 = [];
        foreach ($data_chart1 as $key => $value) {
            $status_RCI = Data::where(DB::raw('DATE(created)'), '<=', $end_date)
                ->where('problem', '=', $value->problem)
                ->where('status', '=', 'Root Cause Identified')
                ->count();
            $status_pending = Data::where(DB::raw('DATE(created)'), '<=', $end_date)
                ->where('problem', '=', $value->problem)
                ->where('status', '=', 'Pending')
                ->count();
            // $closed_thisweek = Data::whereBetween(DB::raw('DATE(changed_at)'), [$start_date, $end_date])
            //     ->where('problem', '=', $value->problem)
            //     ->where('status', '=', 'Closed')
            //     ->count();

            //set color to chart
            $color = '';
            if ($value->problem == 'Core & Surrounding') {
                $color = 'ff89a64e';
            } else if ($value->problem == 'Ekosistem MPC') {
                $color = 'ff00b0f0';
            } else if ($value->problem == 'Loan') {
                $color = 'ffa6a6a6';
            } else if ($value->problem == 'Onboarding') {
                $color = 'ff81ff63';
            } else if ($value->problem == 'Online Payment') {
                $color = 'ff09b1a7';
            } else if ($value->problem == 'Switching & 3rdparty') {
                $color = 'ffee52e1';
            } else if ($value->problem == 'Transaction') {
                $color = 'ff8380ee';
            } else if ($value->problem == 'Wholesale Banking') {
                $color = 'ff8064a2';
            } else {
                $color = 'ffffffff';
            }

            // insert data to array
            $resultdata_chart1[] =
                [
                    'problem' => $value->problem,
                    'total' => $value->count,
                    'count_RCI' => $status_RCI,
                    'count_pending' => $status_pending,
                    // 'closed_thisweek' => $closed_thisweek,
                    'color' => $color
                ];
        }

        // set chart shape
        $chartShape = $slide3->createChartShape();
        $chartShape->setHeight(200)
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
            $series = new Series($value['problem'], ['RC Identified' => $value['count_RCI'], 'Pending' => $value['count_pending']]);
            $series->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color($value['color'])); // Blue
            $chartType->addSeries($series);
        }


        // -------------------- CHART 2 ---------------------
        //Declare DAY
        $lastweek = [Carbon::parse($start_date)->subDays(7), Carbon::parse($start_date)->subDays(1)];
        $twoweeksago = [Carbon::parse($start_date)->subDays(14), Carbon::parse($start_date)->subDays(8)];
        $threeweeksago = [Carbon::parse($start_date)->subDays(21), Carbon::parse($start_date)->subDays(15)];
        //
        $changed_closed_lweek = Data::whereBetween('changed_at', $lastweek)->where('status', '=', 'Closed')->get();
        $changed_closed_2week = Data::whereBetween('changed_at', $twoweeksago)->where('status', '=', 'Closed')->get();
        $changed_closed_3week = Data::whereBetween('changed_at', $threeweeksago)->where('status', '=', 'Closed')->get();

        // $created_closed_lweek = Data::whereBetween(DB::raw('DATE(created)'), $lastweek)->where('status', '=', 'Closed')->get();
        $created_closed_2week = Data::whereBetween(DB::raw('DATE(created)'), $twoweeksago)->where('status', '=', 'Closed')->get();
        $created_closed_3week = Data::whereBetween(DB::raw('DATE(created)'), $threeweeksago)->where('status', '=', 'Closed')->get();

        $created_pending_lweek = Data::whereBetween(DB::raw('DATE(created)'), $lastweek)->where('status', '=', 'Pending')->get()->count();
        $created_pending_2week = Data::whereBetween(DB::raw('DATE(created)'), $twoweeksago)->where('status', '=', 'Pending')->get()->count();
        $created_pending_3week = Data::whereBetween(DB::raw('DATE(created)'), $threeweeksago)->where('status', '=', 'Pending')->get()->count();

        $created_rci_1week = Data::whereBetween(DB::raw('DATE(created)'), $lastweek)->where('status', '=', 'Root Cause Identified')->get()->count();
        $created_rci_2week = Data::whereBetween(DB::raw('DATE(created)'), $twoweeksago)->where('status', '=', 'Root Cause Identified')->get()->count();
        $created_rci_3week = Data::whereBetween(DB::raw('DATE(created)'), $threeweeksago)->where('status', '=', 'Root Cause Identified')->get()->count();

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
        // foreach ($created_closed_lweek as $key => $value) {
        //     $uniqueKey = createUniqueKey($value);
        //     if (!in_array($uniqueKey, $temp1week)) {
        //         $temp1week[] = $uniqueKey;
        //         $closedlweek[] = [
        //             'problem' => $value->summary,
        //             'created' => $value->created,
        //             'status' => $value->status
        //         ];
        //     }
        // }
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

        // set chart shape
        $chartShape = $slide3->createChartShape();
        $chartShape->setHeight(200)
            ->setWidth(820)
            ->setOffsetX(440)
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


        //set data
        $series = new Series('Closed', ['3 Weeks Ago' => count($closed3week), '2 Weeks Ago' => count($closed2week), 'Last Weeks' => count($closedlweek)]);
        $series2 = new Series('RC Identified', ['3 Weeks Ago' => $created_rci_3week, '2 Weeks Ago' => $created_rci_2week, 'Last Weeks' => $created_rci_1week]);
        $series3 = new Series('Pending', ['3 Weeks Ago' => $created_pending_3week, '2 Weeks Ago' => $created_pending_2week, 'Last Weeks' => $created_pending_lweek]);

        //coloring category
        $series->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ff00b050'));
        $series2->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('fff85208'));
        $series3->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('fff6f610'));

        //set series
        $chartType->addSeries($series);
        $chartType->addSeries($series2);
        $chartType->addSeries($series3);


        // -------------------- TABLE DETAIL PENDING/RCI TICKET THIS WEEK ---------------------

        // TITLE TABLE
        $titleTable = $slide3->createRichTextShape();
        $titleTable->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $titleTable->setHeight(50);
        $titleTable->setWidth(410);
        $titleTable->setOffsetX(25);
        $titleTable->setOffsetY(425);
        //coloring
        $titleTable->getFill()->setFillType(Fill::FILL_SOLID);
        $titleTable->getFill()->setStartColor(new Color('ffddd9c3'));
        //set margin
        $titleTable->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $titleTable->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
        // Create a TextRun for "Ticket Detail This Week" with bold formatting
        $textRun1 = $titleTable->createTextRun('Ticket Detail This Week');
        $textRun1->getFont()->setBold(true);
        $textRun1->getFont()->setSize(10); // Set the desired font size here
        // Create another TextRun for the second line with custom font size
        $textRun2 = $titleTable->createTextRun("\nPending problems and RC identified problems created this week + Newly closed problems this week");
        $textRun2->getFont()->setSize(9); // Set the desired font size here

        // Define table properties
        $columns = 5; // Number of columns
        $tableShape = $slide3->createTableShape($columns);
        $tableShape->getBorder()->setLineStyle(Border::LINE_SINGLE);

        // Set the table's position and size
        $tableShape->setHeight(210);
        $tableShape->setWidth(410);
        $tableShape->setOffsetX(25);
        $tableShape->setOffsetY(475);

        // GET DATA FROM DATABASE
        $data_table = Data::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])
            ->whereIn('status', ['Pending', 'Root Cause Identified'])
            ->select('code_jira', 'problem', 'category', 'summary', 'status', 'created', 'changed_at', 'rca_time', 'closed_time')
            ->union(
                Data::whereBetween(DB::raw('DATE(changed_at)'), [$start_date, $end_date])
                    ->where('status', '=', 'Closed')
                    ->select('code_jira', 'problem', 'category', 'summary', 'status', 'created', 'changed_at', 'rca_time', 'closed_time')
            )
            ->get();

        // DEFINE ARRAY
        $tempdata = [
            ['', 'Category', 'Summary', 'Status', 'RCA Time', 'Complete Time'],
        ];

        // ADD ARRAY DATA
        foreach ($data_table as $key => $value) {
            $tempstatus = $value->status;
            if ($value->status == 'Root Cause Identified') {
                $tempstatus = 'RC Identified';
            }
            $status = $tempstatus . "\n" . Carbon::parse($value->created)->format('d/m/y');
            $summary = "[" . $value->code_jira . "]" . " " . $value->summary;

            //convert date to carbon parse
            $created = Carbon::parse($value->created);
            $rcatime = Carbon::parse($value->rca_time);
            $closed_time = Carbon::parse($value->closed_time);

            //declare rca time
            if ($value->rca_time == null) {
                $rca_time = '-';
            } else {
                $rca_days = intval($created->diffInDays($rcatime));
                $rca_days_string = strval($rca_days) . ' days';
                $rca_time = $rca_days_string . "\n" . Carbon::parse($value->rca_time)->format('d/m/y');
            }

            //declare completion time
            if ($value->closed_time == null) {
                $completion_time = '-';
            } else {
                $completion_days = intval($created->diffInDays($closed_time));
                $completion_days_string = strval($completion_days) . ' Days';
                $completion_time = $completion_days_string . "\n" . Carbon::parse($value->closed_time)->format('d/m/y');
            }

            $tempdata[] = [$value->problem, $value->category, $summary,  $status, $rca_time,  $completion_time];
        }


        // INSERT ARRAY TO TABLE
        foreach ($tempdata as $rowIndex => $row) {
            $tableRow = $tableShape->createRow();
            $tableRow->setHeight(25); // Set the height of the row
            foreach ($row as $cellIndex => $cellText) {
                if ($cellIndex == 0) {
                    continue; // Lewati kolom yang disembunyikan
                }

                //set width
                $cell = $tableRow->nextCell();
                if ($cellIndex == 1) {
                    $cell->setWidth(50);
                } else if ($cellIndex == 2) {
                    $cell->setWidth(195);
                } else if ($cellIndex == 3) {
                    $cell->setWidth(55);
                } else if ($cellIndex == 4) {
                    $cell->setWidth(55);
                } else if ($cellIndex == 5) {
                    $cell->setWidth(55);
                }

                //set status
                $problem = $row[0];
                $status = explode("\n", $row[3]);
                $firstStatus = $status[0];
                // $cell = $tableRow->nextCell();
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
                    if ($cellIndex != 3) {
                        //coloring by problem
                        if ($problem == 'Core & Surrounding') {
                            $cell->getFill()->setStartColor(new Color('ff89a64e'));
                        } else if ($problem == 'Ekosistem MPC') {
                            $cell->getFill()->setStartColor(new Color('ff00b0f0'));
                        } else if ($problem == 'Loan') {
                            $cell->getFill()->setStartColor(new Color('ffa6a6a6'));
                        } else if ($problem == 'Onboarding') {
                            $cell->getFill()->setStartColor(new Color('ff81ff63'));
                        } else if ($problem == 'Online Payment') {
                            $cell->getFill()->setStartColor(new Color('ff09b1a7'));
                        } else if ($problem == 'Switching & 3rdparty') {
                            $cell->getFill()->setStartColor(new Color('ffee52e1'));
                        } else if ($problem == 'Transaction') {
                            $cell->getFill()->setStartColor(new Color('ff8380ee'));
                        } else if ($problem == 'Wholesale Banking') {
                            $cell->getFill()->setStartColor(new Color('ff8064a2'));
                        } else {
                            $cell->getFill()->setStartColor(new Color('ffffffff'));
                        }
                    } else if ($cellIndex == 3) {
                        //coloring by status
                        if ($firstStatus == 'Pending') {
                            $cell->getFill()->setStartColor(new Color('fff6f610'));
                        } elseif ($firstStatus == 'Closed') {
                            $cell->getFill()->setStartColor(new Color('ff14ca66'));
                        } elseif ($firstStatus == 'RC Identified') {
                            $cell->getFill()->setStartColor(new Color('fff85208'));
                        } else {
                            $cell->getFill()->setFillType(Fill::FILL_NONE);
                        }
                    } else {
                        $cell->getFill()->setFillType(Fill::FILL_NONE);
                    }
                }
            }
        }



        // -------------------- TABLE DETAIL PENDING/RCI TICKET LAST WEEK ---------------------
         // TITLE TABLE
         $titleTable = $slide3->createRichTextShape();
         $titleTable->getBorder()->setLineStyle(Border::LINE_SINGLE);
         $titleTable->setHeight(50);
         $titleTable->setWidth(820);
         $titleTable->setOffsetX(440);
         $titleTable->setOffsetY(425);
         //coloring
         $titleTable->getFill()->setFillType(Fill::FILL_SOLID);
         $titleTable->getFill()->setStartColor(new Color('ffddd9c3'));
         //set margin
         $titleTable->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
         $titleTable->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
         // Create a TextRun for "Ticket Detail This Week" with bold formatting
         $textRun1 = $titleTable->createTextRun('Ticket Detail This Week');
         $textRun1->getFont()->setBold(true);
         $textRun1->getFont()->setSize(10); // Set the desired font size here
         // Create another TextRun for the second line with custom font size
         $textRun2 = $titleTable->createTextRun("\nPending problems and RC identified problems created this week + Newly closed problems this week");
         $textRun2->getFont()->setSize(9); // Set the desired font size here

        // Define table properties
        $columns = 5; // Number of columns
        $tableShape = $slide3->createTableShape($columns);
        $tableShape->getBorder()->setLineStyle(Border::LINE_SINGLE);

        // Set the table's position and size
        $tableShape->setHeight(210);
        $tableShape->setWidth(410);
        $tableShape->setOffsetX(440);
        $tableShape->setOffsetY(475);

        // Define the data for the table
        $lastweek = [Carbon::parse($start_date)->subDays(7), Carbon::parse($start_date)->subDays(1)];
        $data_table_lastweek = Data::whereBetween(DB::raw('DATE(created)'), $lastweek)
            ->whereIn('status', ['Pending', 'Root Cause Identified'])
            ->select('code_jira', 'problem', 'category', 'summary', 'status', 'created', 'changed_at', 'rca_time', 'closed_time')
            ->union(
                Data::whereBetween(DB::raw('DATE(changed_at)'), $lastweek)
                    ->where('status', '=', 'Closed')
                    ->select('code_jira', 'problem', 'category', 'summary', 'status', 'created', 'changed_at', 'rca_time', 'closed_time')
            )
            ->get();

        //SET TABLE HEADER
        $tempdata = [
            ['', 'Category', 'Summary', 'Status', 'RCA Time', 'Complete Time'],
        ];

        //SET TABLE DATA
        foreach ($data_table_lastweek as $key => $value) {
            $tempstatus = $value->status;
            if ($value->status == 'Root Cause Identified') {
                $tempstatus = 'RC Identified';
            }
            $status = $tempstatus . "\n" . Carbon::parse($value->created)->format('d/m/y');
            $summary = "[" . $value->code_jira . "]" . " " . $value->summary;

            //convert date to carbon parse
            $created = Carbon::parse($value->created);
            $changed_at = Carbon::parse($value->changed_at);
            $rcatime = Carbon::parse($value->rca_time);

            //declare rca time
            if ($value->rca_time == null) {
                $rca_time = '-';
            } else {
                $rca_days = intval($created->diffInDays($rcatime));
                $rca_days_string = strval($rca_days) . ' days';
                $rca_time = $rca_days_string . "\n" . Carbon::parse($value->rca_time)->format('d/m/y');
            }

            if ($value->closed_time == null) {
                $completion_time = '-';
            } else {
                $completion_days = intval($created->diffInDays($closed_time));
                $completion_days_string = strval($completion_days) . ' Days';
                $completion_time = $completion_days_string . "\n" . Carbon::parse($value->closed_time)->format('d/m/y');
            }

            $tempdata[] = [$value->problem, $value->category, $summary,  $status, $rca_time,  $completion_time];
        }

        // dd($tempdata);


        // SET ARRAY TO TABLE
        foreach ($tempdata as $rowIndex => $row) {
            $tableRow = $tableShape->createRow();
            $tableRow->setHeight(25); // Set the height of the row
            foreach ($row as $cellIndex => $cellText) {
                if ($cellIndex == 0) {
                    continue; // Lewati kolom yang disembunyikan
                }

                //set width
                $cell = $tableRow->nextCell();
                if ($cellIndex == 1) {
                    $cell->setWidth(50);
                } else if ($cellIndex == 2) {
                    $cell->setWidth(195);
                } else if ($cellIndex == 3) {
                    $cell->setWidth(55);
                } else if ($cellIndex == 4) {
                    $cell->setWidth(55);
                } else if ($cellIndex == 5) {
                    $cell->setWidth(55);
                }

                //set status
                $problem = $row[0];
                $status = explode("\n", $row[3]);
                $firstStatus = $status[0];
                // $cell = $tableRow->nextCell();
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
                    if ($cellIndex != 3) {
                        //coloring by problem
                        if ($problem == 'Core & Surrounding') {
                            $cell->getFill()->setStartColor(new Color('ff89a64e'));
                        } else if ($problem == 'Ekosistem MPC') {
                            $cell->getFill()->setStartColor(new Color('ff00b0f0'));
                        } else if ($problem == 'Loan') {
                            $cell->getFill()->setStartColor(new Color('ffa6a6a6'));
                        } else if ($problem == 'Onboarding') {
                            $cell->getFill()->setStartColor(new Color('ff81ff63'));
                        } else if ($problem == 'Online Payment') {
                            $cell->getFill()->setStartColor(new Color('ff09b1a7'));
                        } else if ($problem == 'Switching & 3rdparty') {
                            $cell->getFill()->setStartColor(new Color('ffee52e1'));
                        } else if ($problem == 'Transaction') {
                            $cell->getFill()->setStartColor(new Color('ff8380ee'));
                        } else if ($problem == 'Wholesale Banking') {
                            $cell->getFill()->setStartColor(new Color('ff8064a2'));
                        } else {
                            $cell->getFill()->setStartColor(new Color('ffffffff'));
                        }
                    } else if ($cellIndex == 3) {
                        //coloring by status
                        if ($firstStatus == 'Pending') {
                            $cell->getFill()->setStartColor(new Color('fff6f610'));
                        } elseif ($firstStatus == 'Closed') {
                            $cell->getFill()->setStartColor(new Color('ff14ca66'));
                        } elseif ($firstStatus == 'RC Identified') {
                            $cell->getFill()->setStartColor(new Color('fff85208'));
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
            ->setWidth(1000)
            ->setOffsetX(25)
            ->setOffsetY(15);
        $textRun = $shape->createTextRun('IT Problem - Ticket Closed and Service Request');
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


        // ------------ CHART 1 / Problem Category Closed ----------------
        $data_chart1 = Data::where(DB::raw('DATE(created)'), '<=', $end_date)->select('problem', DB::raw('count(*) as count'))->groupBy('problem')->get();
        $resultdata_chart1 = [];
        foreach ($data_chart1 as $key => $value) {
            $status_closed = Data::where(DB::raw('DATE(created)'), '<=', $end_date)
                ->where('problem', '=', $value->problem)
                ->where('status', '=', 'Closed')
                ->count();

            //set color to chart
            $color = '';
            if ($value->problem == 'Core & Surrounding') {
                $color = 'ff89a64e';
            } else if ($value->problem == 'Ekosistem MPC') {
                $color = 'ff00b0f0';
            } else if ($value->problem == 'Loan') {
                $color = 'ffa6a6a6';
            } else if ($value->problem == 'Onboarding') {
                $color = 'ff81ff63';
            } else if ($value->problem == 'Online Payment') {
                $color = 'ff09b1a7';
            } else if ($value->problem == 'Switching & 3rdparty') {
                $color = 'ffee52e1';
            } else if ($value->problem == 'Transaction') {
                $color = 'ff8380ee';
            } else if ($value->problem == 'Wholesale Banking') {
                $color = 'ff8064a2';
            } else {
                $color = 'ffffffff';
            }

            // insert data to array
            $resultdata_chart1[] =
                [
                    'problem' => $value->problem,
                    'total' => $value->count,
                    'count_closed' => $status_closed,
                    'color' => $color
                ];
        }

        // set chart shape
        $chartShape = $slide4->createChartShape();
        $chartShape->setHeight(250)
            ->setWidth(410)
            ->setOffsetX(25)
            ->setOffsetY(115);

        // Define tipe chart
        $chartType = new Bar();
        $chartShape->getPlotArea()->setType($chartType);

        // Set judul chart
        $chartShape->getTitle()->setText('Problem Category Status Closed');

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
            $series = new Series($value['problem'], ['Closed' => $value['count_closed']]);
            $series->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color($value['color'])); // Blue
            $chartType->addSeries($series);
        }



        // ------------ CHART 2 / Ticket Service Request Nasabah ----------------
        $data_chart3 = Service::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('issue_type', '=', '[JSM] Allo Care Service Request')->select('sub_category', DB::raw('count(*) as count'))->groupBy('sub_category')->get();
        $resultdata_chart3 = [];
        foreach ($data_chart3 as $key => $value) {
            $total = Service::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('sub_category', '=', $value->sub_category)->get()->count();
            $status_closed = Service::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('sub_category', '=', $value->sub_category)->where('status', '=', 'Closed')->get()->count();
            $status_declined = Service::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('sub_category', '=', $value->sub_category)->where('status', '=', 'Declined')->get()->count();
            $status_review = Service::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('sub_category', '=', $value->sub_category)->where('status', '=', 'Review')->get()->count();
            $status_userconfirmation = Service::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('sub_category', '=', $value->sub_category)->where('status', '=', 'User Confirmation')->get()->count();
            $resultdata_chart3[] =
                [
                    'sub_category' => $value->sub_category,
                    'total' => $total,
                    'count_closed' => $status_closed,
                    'count_declined' => $status_declined,
                    'count_review' => $status_review,
                    'count_userconfirmation' => $status_userconfirmation
                ];
        }

        // Set Size Chart
        $chartShape = $slide4->createChartShape();
        $chartShape->setHeight(250)
            ->setWidth(410)
            ->setOffsetX(435)
            ->setOffsetY(115);
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
            $series = new Series($value['sub_category'], ['Total' => $value['total'], 'Closed' => $value['count_closed'], 'Declined' => $value['count_declined'], 'Review' => $value['count_review'], 'User Confirmation' => $value['count_userconfirmation']]);
            $chartType->addSeries($series);
        }


        // ------------ CHART 3 / Ticket Service Customer Care ----------------
        $data_chart4 = Service::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('issue_type', '=', '[JSM] Contact Center Request')->select('sub_category', DB::raw('count(*) as count'))->groupBy('sub_category')->get();
        $resultdata_chart4 = [];
        foreach ($data_chart4 as $key => $value) {
            $total = Service::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('sub_category', '=', $value->sub_category)->get()->count();
            $status_closed = Service::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('sub_category', '=', $value->sub_category)->where('status', '=', 'Closed')->get()->count();
            $status_declined = Service::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('sub_category', '=', $value->sub_category)->where('status', '=', 'Declined')->get()->count();
            $status_approval = Service::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('sub_category', '=', $value->sub_category)->where('status', 'like', '%' . 'Approval' . '%')->get()->count();
            $status_inprogress = Service::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('sub_category', '=', $value->sub_category)->where('status', '=', 'In Progress')->get()->count();
            $status_userconfirmation = Service::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('sub_category', '=', $value->sub_category)->where('status', '=', 'User Confirmation')->get()->count();
            $resultdata_chart4[] =
                [
                    'sub_category' => $value->sub_category,
                    'total' => $total,
                    'count_closed' => $status_closed,
                    'count_declined' => $status_declined,
                    'count_approval' => $status_approval,
                    'count_inprogress' => $status_inprogress,
                    'count_userconfirmation' => $status_userconfirmation
                ];
        }
        // Set Size Chart
        $chartShape = $slide4->createChartShape();
        $chartShape->setHeight(250)
            ->setWidth(410)
            ->setOffsetX(845)
            ->setOffsetY(115);
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
            $series = new Series($value['sub_category'], ['Total' => $value['total'], 'Closed' => $value['count_closed'], 'Declined' => $value['count_declined'], 'Approval' => $value['count_approval'], 'In Progress' => $value['count_inprogress'], 'User Confirmation' => $value['count_userconfirmation']]);
            $chartType->addSeries($series);
        }


        // -------------------- TABLE TICKET IT PROBLEM HIGH --------------------
        $data_combined = Data::whereBetween(DB::raw('DATE(changed_at)'), [$start_date, $end_date])
            ->where('priority', '=', 'High')
            ->where('status', '=', 'Closed')
            ->orderBy('problem', 'asc')
            ->union(
                Data::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])
                    ->where('priority', '=', 'High')
                    ->where('status', '!=', 'Closed')
            )
            ->orderBy('problem', 'asc')
            ->get();

        $table = [];

        foreach ($data_combined as $key => $value) {
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

            $rootcause = $value->root_cause ?? ' - ';

            //Status
            $status = $value->status . "\n" . Carbon::parse($value->created)->format('d/m/Y');

            //RCA Days
            $created_rca = Carbon::parse($value->created);
            $rca = Carbon::parse($value->rca_time);
            $rca_nod = intval($created_rca->diffInDays($rca));
            $rca_nod_string = strval($rca_nod) . ' days';
            $rca_time = $rca_nod_string . "\n" . Carbon::parse($value->rca_time)->format('d/m/Y');

            //Completion Days
            $created = Carbon::parse($value->created);
            if ($value->status == 'Closed') {
                $completion = Carbon::parse($value->changed_at);
                $completion_nod = intval($created->diffInDays($completion));
                $completion_nod_string = strval($completion_nod) . ' days';
                $completion_time = $completion_nod_string . "\n" . Carbon::parse($value->changed_at)->format('d/m/Y');
            } else {
                $completion_time = '-';
            }

            //insert to table
            $table[] = [$value->code_jira, $value->problem, $value->summary, $pending_reason, $target_version, $rootcause, $status, $rca_time, $completion_time];
        }

        // dd($table);

        //Table 1

        // CHECK DATA JIKA KOSONG MAKA TIDAK TAMPIL
        if ($table == []) {
        } else {
            // Tambahkan teks judul slide
            $shape = $slide4->createRichTextShape()
                ->setHeight(50)
                ->setWidth(1000)
                ->setOffsetX(25)
                ->setOffsetY(380);
            $textRun = $shape->createTextRun('Priority High this Week');
            $textRun->getFont()->setBold(true)
                ->setSize(30);
            $columns = 9;
            $tableShape = $slide4->createTableShape($columns);
            $tableShape->getBorder()->setLineStyle(Border::LINE_SINGLE);
            $tableShape->setHeight(300);
            $tableShape->setWidth(1200);
            $tableShape->setOffsetX(25);
            $tableShape->setOffsetY(440);
            $rowHeader = $tableShape->createRow();
            $rowHeader->setHeight(25);
            //header 
            $header = ['Code Jira', 'Problem', 'Summary', 'Pending Reason', 'Target Version', 'Root Cause', 'Status' . "\n" . 'Created time', 'RCA Time', 'Complete Time'];
            foreach ($header as $cellIndex => $cellText) {
                $cell = $rowHeader->nextCell();
                if ($cellIndex == 0) {
                    $cell->setWidth(50);
                } else if ($cellIndex == 1) {
                    $cell->setWidth(120);
                } else if ($cellIndex == 2) {
                    $cell->setWidth(350);
                } else if ($cellIndex == 3) {
                    $cell->setWidth(89);
                } else if ($cellIndex == 4) {
                    $cell->setWidth(89);
                } else if ($cellIndex == 5) {
                    $cell->setWidth(300);
                } else if ($cellIndex == 6) {
                    $cell->setWidth(74);
                } else if ($cellIndex == 7) {
                    $cell->setWidth(74);
                } else if ($cellIndex == 8) {
                    $cell->setWidth(74);
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
            foreach ($table as $rowIndex => $row) {
                $tableRow = $tableShape->createRow();
                $tableRow->setHeight(25);
                foreach ($row as $cellIndex => $cellText) {
                    $cell = $tableRow->nextCell();
                    if ($cellIndex == 0) {
                        $cell->setWidth(50);
                    } else if ($cellIndex == 1) {
                        $cell->setWidth(120);
                    } else if ($cellIndex == 2) {
                        $cell->setWidth(350);
                    } else if ($cellIndex == 3) {
                        $cell->setWidth(89);
                    } else if ($cellIndex == 4) {
                        $cell->setWidth(89);
                    } else if ($cellIndex == 5) {
                        $cell->setWidth(300);
                    } else if ($cellIndex == 6) {
                        $cell->setWidth(74);
                    } else if ($cellIndex == 7) {
                        $cell->setWidth(74);
                    } else if ($cellIndex == 8) {
                        $cell->setWidth(74);
                    }
                    $textRun = $cell->createTextRun($cellText);
                    $cell->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                    $cell->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
                    //coloring by problem
                    if ($row[1] == 'Core & Surrounding') {
                        $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ff89a64e'));
                    } else if ($row[1] == 'Ekosistem MPC') {
                        $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ff00b0f0'));
                    } else if ($row[1] == 'Loan') {
                        $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ffa6a6a6'));
                    } else if ($row[1] == 'Onboarding') {
                        $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ff81ff63'));
                    } else if ($row[1] == 'Online Payment') {
                        $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ff09b1a7'));
                    } else if ($row[1] == 'Switching & 3rdparty') {
                        $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ffee52e1'));
                    } else if ($row[1] == 'Transaction') {
                        $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ff8380ee'));
                    } else if ($row[1] == 'Wholesale Banking') {
                        $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ff8064a2'));
                    } else {
                        $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ffffffff'));
                    }
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
        $excelPathProduct = 'exports/List Problem for Product.xlsx';
        Excel::store(new DataExport($start_date, $end_date), $excelPathProduct, 'local');

        // Simpan file Excel sementara
        $excelPathWebank = 'exports/List Problem for Webank.xlsx';
        Excel::store(new allDataExport($start_date, $end_date), $excelPathWebank, 'local');

        // 3. Buat file ZIP yang berisi kedua file tersebut
        $zipFilename = 'Report Weekly IT Problem' . ' - ' . Carbon::parse($start_date)->format('d') . ' - ' . Carbon::parse($end_date)->format('d F Y') . '.zip';
        $zipFilePath = storage_path('app/exports/' . $zipFilename);
        $zip = new ZipArchive;
        if ($zip->open($zipFilePath, ZipArchive::CREATE) === TRUE) {
            $zip->addFile(storage_path('app/' . $excelPathProduct), 'List Problem for Product.xlsx');
            $zip->addFile(storage_path('app/' . $excelPathWebank), 'List Problem for Webank.xlsx');
            $zip->addFile($savePath, $filename);
            $zip->close();
        }

        // 4. Hapus file sementara setelah digabungkan
        Storage::delete([$excelPathProduct, $excelPathWebank]);
        unlink($savePath); // Menghapus file PPT secara manual karena disimpan di luar storage facade

        // 5. Unduh file ZIP dan hapus setelah diunduh
        return response()->download($zipFilePath)->deleteFileAfterSend(true);
    }
    //
}
