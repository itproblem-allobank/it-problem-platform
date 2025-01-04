<?php

namespace App\Http\Controllers;

use App\Models\Data;
use App\Models\Service;
use Illuminate\Support\Carbon;
use Illuminate\Http\Request;
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
use PhpOffice\PhpPresentation\Shape\Chart\Type\Line;
use Maatwebsite\Excel\Facades\Excel;
use Illuminate\Support\Facades\Storage;
use App\Exports\DataExport;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Pie3D;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Pie;
use ZipArchive;

use Exception;

class MonthlyController extends Controller
{
    public function __construct()
    {
        $this->middleware('auth');
    }

    public function index()
    {

        return view('monthly');
    }

    public function download(Request $request)
    {
        $start_date = $request->start_date;
        $end_date = $request->end_date;
        // dd($request->end_date);
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
        $textRun = $shape->createTextRun('Report Monthly IT Problem');
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
        setCellText($row, $cell, 'Report Monthly IT Problem', 15);

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
        $date = Carbon::parse($end_date)->format('F Y');
        $textRun = $shape->createTextRun('As of ' . $date);
        $textRun->getFont()->setSize(14);

        //data container category
        $problem = Data::select('problem', DB::raw('count(*) as count'))->groupBy('problem')->where('problem', '!=', 'Enhancement')->get();
        // dd($problem);
        $total = [];
        foreach ($problem as $key => $value) {
            $high_lastmonth = Data::where(DB::raw('DATE(created)'), '<', $start_date)
                ->where('problem', '=', $value->problem)
                ->where('priority', '=', 'High')
                ->whereIn('status', ['Root Cause Identified', 'Pending'])
                ->union(Data::where(DB::raw('DATE(created)'), '<', $start_date)
                    ->whereBetween(DB::raw('DATE(closed_time)'), [$start_date, $end_date])
                    ->where('problem', '=', $value->problem)
                    ->where('priority', '=', 'High')
                    ->where('status', '=', 'Closed'))
                ->count();

            $medium_lastmonth = Data::where(DB::raw('DATE(created)'), '<', $start_date)
                ->where('problem', '=', $value->problem)
                ->where('priority', '=', 'Medium')
                ->whereIn('status', ['Root Cause Identified', 'Pending'])
                ->union(Data::where(DB::raw('DATE(created)'), '<', $start_date)
                    ->whereBetween(DB::raw('DATE(closed_time)'), [$start_date, $end_date])
                    ->where('problem', '=', $value->problem)
                    ->where('priority', '=', 'Medium')
                    ->where('status', '=', 'Closed'))
                ->count();

            $low_lastmonth = Data::where(DB::raw('DATE(created)'), '<', $start_date)
                ->where('problem', '=', $value->problem)
                ->where('priority', '=', 'Low')
                ->whereIn('status', ['Root Cause Identified', 'Pending'])
                ->union(Data::where(DB::raw('DATE(created)'), '<', $start_date)
                    ->whereBetween(DB::raw('DATE(closed_time)'), [$start_date, $end_date])
                    ->where('problem', '=', $value->problem)
                    ->where('priority', '=', 'Low')
                    ->where('status', '=', 'Closed'))
                ->count();

            $high_thismonth = Data::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])
                ->where('problem', '=', $value->problem)
                ->where('priority', '=', 'High')
                ->count();

            $medium_thismonth = Data::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])
                ->where('problem', '=', $value->problem)
                ->where('priority', '=', 'Medium')
                ->count();

            $low_thismonth = Data::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])
                ->where('problem', '=', $value->problem)
                ->where('priority', '=', 'Low')
                ->count();

            $high_closed_thismonth = Data::whereBetween(DB::raw('DATE(changed_at)'), [$start_date, $end_date])
                ->where('problem', '=', $value->problem)
                ->where('priority', '=', 'High')
                ->where('status', '=', 'Closed')
                ->count();

            $medium_closed_thismonth = Data::whereBetween(DB::raw('DATE(changed_at)'), [$start_date, $end_date])
                ->where('problem', '=', $value->problem)
                ->where('priority', '=', 'Medium')
                ->where('status', '=', 'Closed')
                ->count();

            $low_closed_thismonth = Data::whereBetween(DB::raw('DATE(changed_at)'), [$start_date, $end_date])
                ->where('problem', '=', $value->problem)
                ->where('priority', '=', 'Low')
                ->where('status', '=', 'Closed')
                ->count();

            // COUNT DATA
            $total_high = $high_lastmonth + $high_thismonth - $high_closed_thismonth;
            $total_medium = $medium_lastmonth + $medium_thismonth - $medium_closed_thismonth;
            $total_low = $low_lastmonth + $low_thismonth - $low_closed_thismonth;

            $total_count = $total_high + $total_medium + $total_low;

            // SET COLOR
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


            $total[] = [
                'problem' => $value->problem,
                'total' => $total_count,
                'color' => $color,
                'high_lastmonth' => $high_lastmonth,
                'medium_lastmonth' => $medium_lastmonth,
                'low_lastmonth' => $low_lastmonth,
                'high_thismonth' => $high_thismonth,
                'medium_thismonth' => $medium_thismonth,
                'low_thismonth' => $low_thismonth,
                'high_closed_thismonth' => $high_closed_thismonth,
                'medium_closed_thismonth' => $medium_closed_thismonth,
                'low_closed_thismonth' => $low_closed_thismonth,
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
            $value = [
                $data['high_lastmonth'],
                $data['medium_lastmonth'],
                $data['low_lastmonth']
            ];
            foreach ($value as $key => $v) {
                $cell = $rowShape->nextCell();
                $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color($data["color"]));
                $cell->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                $cell->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
                $cell->createTextRun($v);
            }

            $rowShape = $tableShape->createRow();
            $rowShape->setHeight(20);
            $value = [
                $data['high_thismonth'],
                $data['medium_thismonth'],
                $data['low_thismonth']
            ];

            foreach ($value as $key => $v) {
                $cell = $rowShape->nextCell();
                $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color($data["color"]));
                $cell->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                $cell->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
                $cell->createTextRun($v);
            }

            $rowShape = $tableShape->createRow();
            $rowShape->setHeight(20);
            $value = [
                $data['high_closed_thismonth'],
                $data['medium_closed_thismonth'],
                $data['low_closed_thismonth']
            ];

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

        $total_last_month = 0;
        $total_this_month = 0;
        $total_closed_this_month = 0;
        $total_high = 0;
        $total_medium = 0;
        $total_low = 0;

        foreach ($total as $key => $value) {
            $total_last_month += $value["high_lastmonth"] + $value["medium_lastmonth"] + $value["low_lastmonth"];
            $total_this_month += $value["high_thismonth"] + $value["medium_thismonth"] + $value["low_thismonth"];
            $total_closed_this_month += $value["high_closed_thismonth"] + $value["medium_closed_thismonth"] + $value["low_closed_thismonth"];
            $total_high += $value["high_lastmonth"] + $value["high_thismonth"] - $value["high_closed_thismonth"];
            $total_medium += $value["medium_lastmonth"] + $value["medium_thismonth"] - $value["medium_closed_thismonth"];
            $total_low += $value["low_lastmonth"] + $value["low_thismonth"] - $value["low_closed_thismonth"];
        }
        // dd($totalcreated, $totalclosed, $totalexisting, $totalhigh);

        // Total HIGH, MED, LOW
        $tableShape = $slide3->createTableShape(3);
        $tableShape->setHeight(100);
        $tableShape->setWidth(400);
        $tableShape->setOffsetX(855);
        $tableShape->setOffsetY(75);

        //row title
        $rowShape = $tableShape->createRow();
        $rowShape->setHeight(20);
        $val = [['status' => 'Total High', 'color' => 'FFFF0000', 'value' => $total_high], ['status' => 'Total Medium', 'color' => 'fffeb909', 'value' => $total_medium], ['status' => 'Total Low', 'color' => 'fffffe00', 'value' => $total_low]];
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
        $textRun = $shape->createTextRun($total_last_month);
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
        $textRun = $shape->createTextRun($total_this_month);
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
        $textRun = $shape->createTextRun($total_closed_this_month);
        $textRun->getFont()->setBold(true)
            ->setSize(12)
            ->setColor(new Color(Color::COLOR_BLACK));




        //set data chart 1
        $data_chart1 = Data::where(DB::raw('DATE(created)'), '<=', $end_date)->select('problem', DB::raw('count(*) as count'))->where('problem', '!=', 'Enhancement')->groupBy('problem')->get();
        $resultdata_chart1 = [];
        foreach ($data_chart1 as $key => $value) {
            $status_closed = Data::where(DB::raw('DATE(created)'), '<=', $end_date)
                ->where('problem', '=', $value->problem)
                ->where('status', '=', 'Closed')
                ->count();
            $status_RCI = Data::where(DB::raw('DATE(created)'), '<=', $end_date)
                ->where('problem', '=', $value->problem)
                ->where('status', '=', 'Root Cause Identified')
                ->count();
            $status_pending = Data::where(DB::raw('DATE(created)'), '<=', $end_date)
                ->where('problem', '=', $value->problem)
                ->where('status', '=', 'Pending')
                ->count();

            // SET COLOR
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
            $resultdata_chart1[] =
                [
                    'problem' => $value->problem,
                    'total' => $value->count,
                    'count_closed' => $status_closed,
                    'count_RCI' => $status_RCI,
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
        $chartShape->getTitle()->setText('Ticket by Status');
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
            $series = new Series($value['problem'], ['Closed' => $value['count_closed'], 'RC Identified' => $value['count_RCI'], 'Pending' => $value['count_pending']]);
            $series->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color($value['color'])); // Blue
            $chartType->addSeries($series);
        }

        // -------------------- CHART 2 ---------------------
        $data_chart2 = Data::whereBetween(DB::raw('DATE(created)'), [Carbon::parse($start_date)->subMonths(3), Carbon::parse($end_date)->subMonths(1)])->select(DB::raw('MONTH(created) as month'), DB::raw('count(*) as count'))
            ->where('problem', '!=', 'Enhancement')
            ->groupBy(DB::raw('MONTH(created)'))
            ->get();
        $resultdata_chart2 = [];
        foreach ($data_chart2 as $key => $value) {
            $closed = Data::whereBetween(DB::raw('DATE(created)'), [Carbon::parse($start_date)->subMonths(3), Carbon::parse($end_date)->subMonths(1)])->where('status', '=', 'Closed')->where(DB::raw('MONTH(created)'), '=', $value->month)->get()->count();
            $rcidentified = Data::whereBetween(DB::raw('DATE(created)'), [Carbon::parse($start_date)->subMonths(3), Carbon::parse($end_date)->subMonths(1)])->where('status', '=', 'Root Cause Identified')->where(DB::raw('MONTH(created)'), '=', $value->month)->get()->count();
            $pending = Data::whereBetween(DB::raw('DATE(created)'), [Carbon::parse($start_date)->subMonths(3), Carbon::parse($end_date)->subMonths(1)])->where('status', '=', 'Pending')->where(DB::raw('MONTH(created)'), '=', $value->month)->get()->count();
            $totalCount = $data_chart2->sum('count');
            $totalValue = $closed + $pending;
            $number = ($totalValue / $totalCount) * 100;
            $percentage = round($number);
            $resultdata_chart2[] = [
                'month' => Carbon::create()->month(intval($value->month))->format('F'),
                'count' => $value->count,
                'closed' => $closed,
                'rcidentified' => $rcidentified,
                'pending' => $pending,
                'percentage' => $percentage
            ];
        }

        // Chart 2
        $chartShape = $slide3->createChartShape();
        $chartShape->setHeight(250)
            ->setWidth(410)
            ->setOffsetX(435)
            ->setOffsetY(225);
        // Define tipe chart
        $chartType = new Bar();
        $chartShape->getPlotArea()->setType($chartType);

        // Set judul chart
        $chartShape->getTitle()->setText('Ticket by Last 3 Months');
        $chartShape->getLegend()->getBorder()->setLineStyle(Border::LINE_NONE); // Menghilangkan kotak pada legenda
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

        $dataclosed = [];
        foreach ($resultdata_chart2 as $key => $value) {
            $dataclosed[$value['month'] . "\n" . ' (' . $value['percentage'] . '%)'] = $value['closed'];
        }
        $datarcidentified = [];
        foreach ($resultdata_chart2 as $key => $value) {
            $datarcidentified[$value['month'] . "\n" . ' (' . $value['percentage'] . '%)'] = $value['rcidentified'];
        }
        $datapending = [];
        foreach ($resultdata_chart2 as $key => $value) {
            $datapending[$value['month'] . "\n" . ' (' . $value['percentage'] . '%)'] = $value['pending'];
        }

        $series = new Series('Closed', $dataclosed);
        $series1 = new Series('Root Cause Identified', $datarcidentified);
        $series2 = new Series('Pending', $datapending);
        $chartType->addSeries($series);
        $chartType->addSeries($series1);
        $chartType->addSeries($series2);

        // Chart 3 Ticket Service Request Jira
        $data_chart3 = Service::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->select('issue_type', DB::raw('count(*) as count'))
            ->groupBy('issue_type')->get();
        $resultdata_chart3 = [];
        foreach ($data_chart3 as $key => $value) {
            $status_closed = Service::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('issue_type', '=', $value->issue_type)->where('status', '=', 'Closed')->get()->count();
            $status_pending = Service::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('issue_type', '=', $value->issue_type)->where('status', '=', 'Pending')->get()->count();
            $resultdata_chart3[] =
                [
                    'issue_type' => $value->issue_type,
                    'total' => $value->count,
                    'count_closed' => $status_closed,
                    'count_pending' => $status_pending,
                ];
        }

        // Set Size Chart
        $chartShape = $slide3->createChartShape();
        $chartShape->setHeight(215)
            ->setWidth(410)
            ->setOffsetX(845)
            ->setOffsetY(475);
        // Define tipe chart
        $chartType = new Bar();
        $chartShape->getPlotArea()->setType($chartType);
        // Set judul chart
        $chartShape->getTitle()->setText('Ticket Jira Service Request');
        $chartShape->getLegend()->getBorder()->setLineStyle(Border::LINE_NONE); // Menghilangkan kotak pada legenda
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

        // Tambahkan seri data ke chart
        foreach ($resultdata_chart3 as $key => $value) {
            $series = new Series($value['issue_type'], ['Closed' => $value['count_closed'], 'Pending' => $value['count_pending']]);
            $chartType->addSeries($series);
        }

        //Chart 5 Problem by Assignee & Status
        $data_chart5 = Data::select('nickname', DB::raw('count(*) as count'))
            ->where('problem', '!=', 'Enhancement')
            ->groupBy('nickname')
            ->get();
        $resultdata_chart5 = [];
        foreach ($data_chart5 as $key => $value) {
            $closed = Data::whereBetween(DB::raw('DATE(closed_time)'), [$start_date, $end_date])->where('nickname', '=', $value->nickname)->where('status', '=', 'Closed')->get()->count();
            $rcidentified = Data::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('nickname', '=', $value->nickname)->where('status', '=', 'Root Cause Identified')->get()->count();
            $pending = Data::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('nickname', '=', $value->nickname)->where('status', '=', 'Pending')->get()->count();
            $resultdata_chart5[] = [
                'nickname' => $value->nickname,
                'count' => $value->count,
                'closed' => $closed,
                'rcidentified' => $rcidentified,
                'pending' => $pending
            ];
        }
        // dd($data_chart5);
        $data_closed = [];
        foreach ($resultdata_chart5 as $key => $value) {
            $data_closed[$value['nickname']] = $value['closed'];
        }
        $data_rcidentified = [];
        foreach ($resultdata_chart5 as $key => $value) {
            $data_rcidentified[$value['nickname']] = $value['rcidentified'];
        }
        $data_pending = [];
        foreach ($resultdata_chart5 as $key => $value) {
            $data_pending[$value['nickname']] = $value['pending'];
        }
        $chartShape = $slide3->createChartShape();
        $chartShape->setHeight(215)
            ->setWidth(410)
            ->setOffsetX(435)
            ->setOffsetY(475);
        // Define tipe chartsss
        $chartType = new Bar();
        $chartShape->getPlotArea()->setType($chartType);
        // Set judul chart
        $chartShape->getTitle()->setText('Problem By Assignee & Status');
        $chartShape->getLegend()->getBorder()->setLineStyle(Border::LINE_NONE); // Menghilangkan kotak pada legenda
        // Mendapatkan objek sumbu
        $xAxis = $chartShape->getPlotArea()->getAxisX();
        $yAxis = $chartShape->getPlotArea()->getAxisY();
        // Mengatur judul sumbu menjadi kosong
        $xAxis->setTitle('');
        $yAxis->setTitle('');
        // Tambahkan seri data ke chart
        $series1 = new Series('Closed', $data_closed);
        $series2 = new Series('Root Cause Identified', $data_rcidentified);
        $series3 = new Series('Pending', $data_pending);
        $chartType->addSeries($series1);
        $chartType->addSeries($series2);
        $chartType->addSeries($series3);
        // Chart Bordered
        $chartShape->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $chartShape->getBorder()->setColor(new Color('FF000000')); // Black border
        $chartShape->getBorder()->setLineWidth(1);


        //Chart 6 Container
        $shape = $slide3->createRichTextShape()
            ->setHeight(250)
            ->setWidth(205)
            ->setOffsetX(845)
            ->setOffsetY(225);
        $shape->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFFFFF'));
        $shape->getBorder()->setLineStyle(Border::LINE_SINGLE)->setColor(new Color('FF000000'));

        // Set Data
        $curr_created = Data::where(DB::raw('DATE(created)'), '<=', $end_date)->where('problem', '!=', 'Enhancement')->get()->count();
        $prev_created = Data::where(DB::raw('DATE(created)'), '<', $start_date)->where('problem', '!=', 'Enhancement')->get()->count();
        $curr_closed = Data::where(DB::raw('DATE(created)'), '<=', $end_date)->where('problem', '!=', 'Enhancement')->where('status', '=', 'Closed')->get()->count();
        $prev_closed = Data::where(DB::raw('DATE(created)'), '<', $start_date)->where('problem', '!=', 'Enhancement')->where('status', '=', 'Closed')->get()->count();
        // dd($curr_created, $prev_created, $curr_closed, $prev_closed);
        $percen_created = ($curr_created - $prev_created) / $prev_created * 100;
        $percen_closed = ($curr_closed - $prev_closed) / $prev_closed * 100;

        // Menambahkan teks ke kotak pertama
        $percentage = $shape->createTextRun("▲ " . number_format($percen_created, 2) . "%");
        $percentage->getFont()->setBold(true)->setSize(28)->setColor(new Color('FFC00000'));
        $title = $shape->createTextRun("\nIssues Created");
        $title->getFont()->setBold(true)->setSize(20)->setColor(new Color('FFC00000'));
        $c_month = $shape->createTextRun("\n\n\nCurrent Month : ");
        $percentage->getFont()->setBold(true)->setSize(28)->setColor(new Color('FFC00000'));
        $c_month->getFont()->setBold(true)->setSize(12);
        $vc_month = $shape->createTextRun("\n" . $curr_created);
        $vc_month->getFont()->setBold(true)->setSize(18);
        $p_month = $shape->createTextRun("\nPrevious Month : ");
        $p_month->getFont()->setBold(true)->setSize(12);
        $vp_month = $shape->createTextRun("\n" . $prev_created);
        $vp_month->getFont()->setBold(true)->setSize(18);

        // Menambahkan kotak kedua untuk "Issues Closed"
        $shape2 = $slide3->createRichTextShape()
            ->setHeight(250)
            ->setWidth(205)
            ->setOffsetX(1050)
            ->setOffsetY(225);
        $shape2->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFFFFF'));
        $shape2->getBorder()->setLineStyle(Border::LINE_SINGLE)->setColor(new Color('FF000000'));

        // Menambahkan teks ke kotak kedua
        $percentage2 = $shape2->createTextRun("▲ " . number_format($percen_closed, 2) . "%");
        $percentage2->getFont()->setBold(true)->setSize(28)->setColor(new Color('FF00C000'));
        $title2 = $shape2->createTextRun("\nIssues Closed");
        $title2->getFont()->setBold(true)->setSize(20)->setColor(new Color('FF00C000'));
        $c_month2 = $shape2->createTextRun("\n\n\nCurrent Month : ");
        $c_month2->getFont()->setBold(true)->setSize(12);
        $vc_month2 = $shape2->createTextRun("\n" . $curr_closed);
        $vc_month2->getFont()->setBold(true)->setSize(18);
        $p_month2 = $shape2->createTextRun("\nPrevious Month : ");
        $p_month2->getFont()->setBold(true)->setSize(12);
        $vp_month2 = $shape2->createTextRun("\n" . $prev_closed);
        $vp_month2->getFont()->setBold(true)->setSize(18);

        // Chart 4 - Chart Line Created vs Closed
        //convert data per week
        $w1 = Carbon::parse($start_date)->addDays(7);
        $w2 = Carbon::parse($start_date)->addDays(14);
        $w3 = Carbon::parse($start_date)->addDays(21);

        //created data
        $totalcr = Data::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])->where('problem', '!=', 'Enhancement')->get()->count();
        $cr1 = Data::whereBetween(DB::raw('DATE(created)'), [$start_date, $w1])->where('problem', '!=', 'Enhancement')->get()->count();
        $cr2 = Data::whereBetween(DB::raw('DATE(created)'), [$w1, $w2])->where('problem', '!=', 'Enhancement')->get()->count();
        $cr3 = Data::whereBetween(DB::raw('DATE(created)'), [$w2, $w3])->where('problem', '!=', 'Enhancement')->get()->count();
        $cr4 = Data::whereBetween(DB::raw('DATE(created)'), [$w3, $end_date])->where('problem', '!=', 'Enhancement')->get()->count();

        //closed data
        $totalcl = Data::where('status', '=', 'Closed')->whereBetween('changed_at', [$start_date, $end_date])->where('problem', '!=', 'Enhancement')->get()->count();
        $cl1 = Data::where('status', '=', 'Closed')->whereBetween('changed_at', [$start_date, $w1])->where('problem', '!=', 'Enhancement')->get()->count();
        $cl2 = Data::where('status', '=', 'Closed')->whereBetween('changed_at', [$w1, $w2])->where('problem', '!=', 'Enhancement')->get()->count();
        $cl3 = Data::where('status', '=', 'Closed')->whereBetween('changed_at', [$w2, $w3])->where('problem', '!=', 'Enhancement')->get()->count();
        $cl4 = Data::where('status', '=', 'Closed')->whereBetween('changed_at', [$w3, $end_date])->where('problem', '!=', 'Enhancement')->get()->count();

        // Set Size Chart
        $chartShape = $slide3->createChartShape();
        $chartShape->setHeight(215)
            ->setWidth(410)
            ->setOffsetX(25)
            ->setOffsetY(475);
        // Define tipe chart
        $chartType = new Line();
        $chartShape->getPlotArea()->setType($chartType);
        // Set judul chart
        $chartShape->getTitle()->setText('Ticket Created vs Closed');
        $chartShape->getLegend()->getBorder()->setLineStyle(Border::LINE_NONE); // Menghilangkan kotak pada legenda
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

        // Tambahkan seri data ke chart
        $created = new Series('Created', ['Week1' => $cr1, 'Week2' => $cr2, 'Week3' => $cr3, 'Week4' => $cr4]);
        $chartType->addSeries($created);
        $closed = new Series('Closed', ['Week1' => $cl1, 'Week2' => $cl2, 'Week3' => $cl3, 'Week4' => $cl4]);
        $chartType->addSeries($closed);


        // ----------------- ADDITIONAL SLIDE ----------------
        $additionalslide = $objPHPPresentation->createSlide();
        $backgroundImagePath = storage_path('image/background.png');
        $backgroundImage = new File();
        $backgroundImage->setPath($backgroundImagePath);
        $backgroundImage->setWidth(1280);
        $backgroundImage->setOffsetX(0);
        $backgroundImage->setOffsetY(0);
        $additionalslide->addShape($backgroundImage);


        $imagePath = storage_path('image/allobank.png');
        $pictureShape = new File();
        $pictureShape->setPath($imagePath);
        $pictureShape->setWidth(200);  // Ubah ukuran gambar sesuai kebutuhan
        $pictureShape->setOffsetX(1050); // Posisi horizontal gambar
        $pictureShape->setOffsetY(20); // Posisi vertikal gambar
        $additionalslide->addShape($pictureShape);

        $objPHPPresentation->getLayout()->setDocumentLayout(['cx' => 1280, 'cy' => 700], true)
            ->setCX(1280, DocumentLayout::UNIT_PIXEL)
            ->setCY(700, DocumentLayout::UNIT_PIXEL);

        // Tambahkan teks judul slide
        $shape = $additionalslide->createRichTextShape()
            ->setHeight(50)
            ->setWidth(700)
            ->setOffsetX(50)
            ->setOffsetY(25);
        $textRun = $shape->createTextRun('Report IT Problem');
        $textRun->getFont()->setBold(true)
            ->setSize(30);


        // -------------- SET CHART --------------------------
        // set title chart
        $titleTable = $additionalslide->createRichTextShape();
        $titleTable->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $titleTable->setHeight(50);
        $titleTable->setWidth(410);
        $titleTable->setOffsetX(25);
        $titleTable->setOffsetY(100);
        $titleTable->getFill()->setFillType(Fill::FILL_SOLID);
        $titleTable->getFill()->setStartColor(new Color('ffddd9c3'));
        $titleTable->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $titleTable->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
        $textRun1 = $titleTable->createTextRun('IT Problem Ticket RCA Time');
        $textRun1->getFont()->setBold(true);
        $textRun1->getFont()->setSize(10);
        $textRun2 = $titleTable->createTextRun("\nCounting IT Problem Tickets by RCA Time Identified (8 Sept - Present)");
        $textRun2->getFont()->setSize(9);

        // Define data
        $days1 = Data::where('created', '>=', '2024-09-01')
            ->where('rca_time', '!=', null)->where('rca_days', '=', 1)
            ->count();
        $days2 = Data::where('created', '>=', '2024-09-01')
            ->where('rca_time', '!=', null)->where('rca_days', '=', 2)
            ->count();
        $days3 = Data::where('created', '>=', '2024-09-01')
            ->where('rca_time', '!=', null)->where('rca_days', '=', 3)
            ->count();
        $days4 = Data::where('created', '>=', '2024-09-01')
            ->where('rca_time', '!=', null)->where('rca_days', '=', 4)
            ->count();
        $days5 = Data::where('created', '>=', '2024-09-01')
            ->where('rca_time', '!=', null)->where('rca_days', '=', 5)
            ->count();
        // $daysover5 = Data::where('rca_time', '!=', null)->where('rca_days', '>', 5)->count();

        $pie_data = ['1 Day' => $days1, '2 Days' => $days2, '3 Days' => $days3, '4 Days' => $days4, '5 Days' => $days5];

        // Create pie chart & Insert to slide
        $pie3DChart = new Pie();
        $pie3DChart->setExplosion(0);
        $series = new Series('RCA Time', $pie_data);
        $series->setShowPercentage(true);
        $series->setShowValue(true);
        $series->setShowSeriesName(false);
        $series->getDataPointFill(0)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ffff0000'));
        $series->getDataPointFill(1)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ffFF4C4C'));
        $series->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('fffeb909'));
        $series->getDataPointFill(3)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFFFC634'));
        $series->getDataPointFill(4)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('fffffe00'));
        $series->getDataPointFill(5)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ffFCFB84'));
        $pie3DChart->addSeries($series);

        /* Create a shape (chart) */
        $shape = $additionalslide->createChartShape();
        $shape->setResizeProportional(false)
            ->setHeight(215)
            ->setWidth(410)
            ->setOffsetX(25)
            ->setOffsetY(150);
        $shape->getTitle()->setText('RCA Time');
        $shape->getTitle()->setVisible(false);
        $shape->getPlotArea()->setType($pie3DChart);
        $shape->getView3D()->setRotationX(40);
        $shape->getView3D()->setPerspective(10);
        //set borders
        $shape->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $shape->getBorder()->setColor(new Color('FF000000')); // Black border
        $shape->getBorder()->setLineWidth(1);
        $shape->getPlotArea()->getAxisY()->setIsVisible(false);
        $shape->getLegend()->getBorder()->setLineStyle(Border::LINE_NONE); // Menghilangkan kotak pada legenda
        //



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
            ->setWidth(700)
            ->setOffsetX(50)
            ->setOffsetY(25);
        $textRun = $shape->createTextRun('Achievement IT Problem by Week');
        $textRun->getFont()->setBold(true)
            ->setSize(30);


        // Data untuk timeline
        $week4 = Data::select(DB::raw('DATE(created)'), 'summary')->whereBetween(DB::raw('DATE(created)'), [Carbon::parse($end_date)->subDays(7), $end_date])->get();
        $week3 = Data::select(DB::raw('DATE(created)'), 'summary')->whereBetween(DB::raw('DATE(created)'), [Carbon::parse($end_date)->subDays(14), Carbon::parse($end_date)->subDays(7)])->get();
        $week2 = Data::select(DB::raw('DATE(created)'), 'summary')->whereBetween(DB::raw('DATE(created)'), [Carbon::parse($end_date)->subDays(21), Carbon::parse($end_date)->subDays(14)])->get();
        $week1 = Data::select(DB::raw('DATE(created)'), 'summary')->whereBetween(DB::raw('DATE(created)'), [Carbon::parse($end_date)->subDays(28), Carbon::parse($end_date)->subDays(21)])->get();

        $data_week1 = [];
        $data_week2 = [];
        $data_week3 = [];
        $data_week4 = [];

        $index1 = 1;
        foreach ($week1 as $key => $value) {
            $data_week1[$key] = $index1 . ". " . $value->summary . "\n\n";
            $index1++;
        }
        $index2 = 1;
        foreach ($week2 as $key => $value) {
            $data_week2[$key] = $index2 . ". " . $value->summary . "\n\n";
            $index2++;
        }
        $index3 = 1;
        foreach ($week3 as $key => $value) {
            $data_week3[$key] = $index3 . ". " . $value->summary . "\n\n";
            $index3++;
        }
        $index4 = 1;
        foreach ($week4 as $key => $value) {
            $data_week4[$key] = $index4 . ". " . $value->summary . "\n\n";
            $index4++;
        }

        $implodeweek1 = implode($data_week1);
        $implodeweek2 = implode($data_week2);
        $implodeweek3 = implode($data_week3);
        $implodeweek4 = implode($data_week4);

        $week = ['Week 1', 'Week 2', 'Week 3', 'Week 4'];
        $descriptions = [$implodeweek1, $implodeweek2, $implodeweek3, $implodeweek4];
        $positions = [100, 400, 700, 1000]; // X positions for the timeline elements

        // Buat timeline
        foreach ($week as $index => $w_index) {
            // Menambahkan tahun
            $shape = $slide4->createRichTextShape();
            $shape->setHeight(50)
                ->setWidth(120)
                ->setOffsetX($positions[$index])
                ->setOffsetY(150);
            $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

            $textRun = $shape->createTextRun($w_index);
            $textRun->getFont()->setBold(true)->setSize(20)->setColor(new Color('FFFFB003'));

            // Menambahkan deskripsi
            $descShape = $slide4->createRichTextShape();
            $descShape->setHeight(450)
                ->setWidth(250)
                ->setOffsetX($positions[$index] - 65)
                ->setOffsetY(200);
            $descShape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
            $descShape->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('CCCCCC'));
            $descShape->getBorder()->setLineStyle(Border::LINE_SINGLE);

            $descTextRun = $descShape->createTextRun($descriptions[$index]);
            $descTextRun->getFont()->setSize(12)->setColor(new Color(Color::COLOR_BLACK));

            // Tambahkan garis penghubung jika bukan elemen terakhir
            $position = [285, 585, 885];
            if ($index < 3) {
                $lineShape = $slide4->createLineShape($position[$index], 420, $position[$index] + 50, 420);
                $lineShape->getBorder()->setLineStyle(Border::LINE_SINGLE);
                $lineShape->getBorder()->setLineWidth(2);
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
        // $filename = 'Report IT Problem ' . Carbon::parse($end_date)->format('F Y') . '.pptx';
        // $savePath = storage_path($filename);
        // $writer = IOFactory::createWriter($objPHPPresentation, 'PowerPoint2007');
        // $writer->save($savePath);
        // // Return file sebagai response download
        // return response()->download($savePath)->deleteFileAfterSend(true);

        // Simpan presentasi ke dalam file
        $filename = 'Report Monthly IT Problem' . ' - ' . Carbon::parse($start_date)->format('F Y') . '.pptx';
        $savePath = storage_path($filename);
        $writer = IOFactory::createWriter($objPHPPresentation, 'PowerPoint2007');
        $writer->save($savePath);

        // Simpan file Excel sementara
        $excelPath = 'exports/List Problem Monthly - ' .  Carbon::parse($start_date)->format('F Y') . '.xlsx';
        Excel::store(new DataExport($start_date, $end_date), $excelPath, 'local');

        // 3. Buat file ZIP yang berisi kedua file tersebut
        $zipFilename = 'Monthly Report - ' . Carbon::parse($start_date)->format('F Y') . '.zip';
        $zipFilePath = storage_path('app/exports/' . $zipFilename);
        $zip = new ZipArchive;
        if ($zip->open($zipFilePath, ZipArchive::CREATE) === TRUE) {
            $zip->addFile(storage_path('app/' . $excelPath), 'List Problem Monthly - ' .  Carbon::parse($start_date)->format('F Y') . '.xlsx');
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
