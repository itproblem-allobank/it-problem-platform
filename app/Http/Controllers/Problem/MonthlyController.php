<?php

namespace App\Http\Controllers\Problem;

use App\Http\Controllers\Controller;
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

        return view('problem/p-monthly');
    }

    public function download(Request $request)
    {

        $start_date = $request->start_date;
        $end_date = $request->end_date;

        $objPHPPresentation = new PhpPresentation();
        // Set Layout
        $objPHPPresentation->getLayout()->setDocumentLayout(
            DocumentLayout::LAYOUT_CUSTOM,
            true // true = landscape, false = portrait
        );

        // Set ukuran slide sesuai kebutuhan
        $objPHPPresentation->getLayout()->setCX(12193200); // width: 33.87 cm
        $objPHPPresentation->getLayout()->setCY(6886800);  // height: 19.13 cm

        // ---------------------------- SLIDE 1 ----------------------------------------------
        $slide1 = $objPHPPresentation->getActiveSlide();
        $backgroundImagePath = storage_path('image/background.png');
        $backgroundImage = new File();
        $backgroundImage->setPath($backgroundImagePath);
        $backgroundImage->setWidth(1280);
        $backgroundImage->setOffsetX(0);
        $backgroundImage->setOffsetY(0);
        $slide1->addShape($backgroundImage);

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
        $textRun = $shape->createTextRun('Monthly Report IT Problem');
        $textRun->getFont()->setBold(true)
            ->setSize(28);

        //Divider
        $lineShape1 = $slide1->createLineShape(50, 355, 1150, 355);
        $lineShape1->getBorder()->setColor(new Color('FF000000'));
        $lineShape1->getBorder()->setLineWidth(2);


        //Text
        $shape = $slide1->createRichTextShape()
            ->setHeight(50)
            ->setWidth(1150)
            ->setOffsetX(50)
            ->setOffsetY(360);
        $textRun1 = $shape->createTextRun('Information Technology Infrastructure & Operations No ');
        $textRun1->getFont()->setBold(true)
            ->setSize(20);
        $textRun2 = $shape->createTextRun('013/DIV-IFO/REP/25');
        $textRun2->getFont()->setBold(true)
            ->setSize(20);

        //Text
        $shape = $slide1->createRichTextShape()
            ->setHeight(50)
            ->setWidth(280)
            ->setOffsetX(980)
            ->setOffsetY(640);
        $textRun = $shape->createTextRun('PT Allo Bank Indonesia');
        $textRun->getFont()->setSize(20);

        //------------------------ SLIDE 2 -----------------------------
        $slide2 = $objPHPPresentation->createSlide();
        $backgroundImagePath = storage_path('image/background.png');
        $backgroundImage = new File();
        $backgroundImage->setPath($backgroundImagePath);
        $backgroundImage->setWidth(1280);
        $backgroundImage->setOffsetX(0);
        $backgroundImage->setOffsetY(0);
        $slide2->addShape($backgroundImage);


        $shape = $slide2->createRichTextShape()
            ->setHeight(50)
            ->setWidth(1000)
            ->setOffsetX(25)
            ->setOffsetY(25);
        $textRun = $shape->createTextRun('Document Control');
        $textRun->getFont()->setBold(true)
            ->setSize(30);

        $imagePath = storage_path('image/Line.png');
        $pictureShape = new File();
        $pictureShape->setPath($imagePath);
        $pictureShape->setWidth(1200);  // Ubah ukuran gambar sesuai kebutuhan
        $pictureShape->setOffsetX(20); // Posisi horizontal gambar
        $pictureShape->setOffsetY(100); // Posisi vertikal gambar
        $slide2->addShape($pictureShape);

        // Add a table for document control details
        $tableShape = $slide2->createTableShape(2);
        $tableShape->setWidth(600);

        // Position the table on the slide
        $tableShape->setOffsetX(50);
        $tableShape->setOffsetY(135);

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

        // Create the bold text run for "Tri Intan Siska P."
        $boldTextRun = $textShape1->createTextRun("Tri Intan Siska P.\n");
        $boldTextRun->getFont()->setSize(15);
        $boldTextRun->getFont()->setColor(new Color('FF000000')); // Black color
        $boldTextRun->getFont()->setBold(true); // Set the text to bold

        // Create the text run for "IT infra Operation"
        $textRun3 = $textShape1->createTextRun("IT Operations Dept. Head");
        $textRun3->getFont()->setSize(15);
        $textRun3->getFont()->setColor(new Color('FF000000')); // Black color

        //Text Shape 2
        $textShape2 = $slide2->createRichTextShape();
        $textShape2->setHeight(250);
        $textShape2->setWidth(300);
        $textShape2->setOffsetX(480);
        $textShape2->setOffsetY(420);
        $textShape2->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);

        // Create the text run for the left-aligned text
        $textRun2 = $textShape2->createTextRun("\n\nDiperiksa oleh,\n\n\n\n\n");
        $textRun2->getFont()->setSize(15);
        $textRun2->getFont()->setColor(new Color('FF000000')); // Black color

        // Create the bold text run for "Tri Intan Siska Permatasari"
        $boldTextRun = $textShape2->createTextRun("Fachri\n");
        $boldTextRun->getFont()->setSize(15);
        $boldTextRun->getFont()->setColor(new Color('FF000000')); // Black color
        $boldTextRun->getFont()->setBold(true); // Set the text to bold

        // Create the text run for "IT Problem Lead"
        $textRun3 = $textShape2->createTextRun("IT Problem Section Head");
        $textRun3->getFont()->setSize(15);
        $textRun3->getFont()->setColor(new Color('FF000000')); // Black color

        //Text Shape 3
        $textShape2 = $slide2->createRichTextShape();
        $textShape2->setHeight(250);
        $textShape2->setWidth(300);
        $textShape2->setOffsetX(900);
        $textShape2->setOffsetY(420);
        $textShape2->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);

        // Create the text run for the left-aligned text
        $textRun2 = $textShape2->createTextRun("\n\nDibuat oleh,\n\n\n\n\n");
        $textRun2->getFont()->setSize(15);
        $textRun2->getFont()->setColor(new Color('FF000000')); // Black color

        // Create the bold text run for "Tri Intan Siska Permatasari"
        $boldTextRun = $textShape2->createTextRun("Ahmad Syauqi\n");
        $boldTextRun->getFont()->setSize(15);
        $boldTextRun->getFont()->setColor(new Color('FF000000')); // Black color
        $boldTextRun->getFont()->setBold(true); // Set the text to bold

        // Create the text run for "IT Problem Lead"
        $textRun3 = $textShape2->createTextRun("IT Problem Engineer");
        $textRun3->getFont()->setSize(15);
        $textRun3->getFont()->setColor(new Color('FF000000')); // Black color




        // ------------------- SLIDE 3 --------------------------
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
        $textRun = $shape->createTextRun('Problem Management');
        $textRun->getFont()->setBold(true)
            ->setSize(30);

        $shape = $slide3->createRichTextShape()
            ->setHeight(25)
            ->setWidth(400)
            ->setOffsetX(25)
            ->setOffsetY(65);
        $date = Carbon::parse($end_date)->format('F Y');
        $textRun = $shape->createTextRun('As of ' . $date);
        $textRun->getFont()->setSize(14);

        $shape = $slide3->createRichTextShape()
            ->setHeight(25)
            ->setWidth(400)
            ->setOffsetX(25)
            ->setOffsetY(110);
        $textRun = $shape->createTextRun('PROBLEM OVERVIEW');
        $textRun->getFont()->setSize(10)->setBold(true);

        $imagePath = storage_path('image/Line.png');
        $pictureShape = new File();
        $pictureShape->setPath($imagePath);
        $pictureShape->setWidth(1200);
        $pictureShape->setOffsetX(20);
        $pictureShape->setOffsetY(100);
        $slide3->addShape($pictureShape);



        //Source Data
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
            if ($value->problem == 'Core Surrounding') {
                $color = 'ff89a64e';
            } else if ($value->problem == 'Ekosistem MPC') {
                $color = 'ff00b0f0';
            } else if ($value->problem == 'Loan') {
                $color = 'ffa6a6a6';
            } else if ($value->problem == 'Onboarding') {
                $color = 'ff81ff63';
            } else if ($value->problem == 'Online Payment') {
                $color = 'ff09b1a7';
            } else if ($value->problem == 'Switching 3rdparty') {
                $color = 'ffee52e1';
            } else if ($value->problem == 'Transaction') {
                $color = 'ff8380ee';
            } else if ($value->problem == 'Wholesale') {
                $color = 'ff8064a2';
            } else if ($value->problem == 'Cybersecurity') {
                $color = 'ffb9cd96';
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
        $offsety = 135;
        //loop category data
        foreach ($total as $key => $data) {
            // Tambahkan tabel dengan 4 baris dan 3 kolom
            $tableShape = $slide3->createTableShape(3);
            $tableShape->setHeight(100);
            $tableShape->setWidth(128);
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
            $offsetx = $offsetx + 137.5;
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


        // Total All IT Problem 
        $shape = $slide3->createRichTextShape();
        $shape->setHeight(25)
            ->setWidth(40)
            ->setOffsetX(1247)
            ->setOffsetY(70);
        $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $textRun = $shape->createTextRun($total_high + $total_medium + $total_low);
        $textRun->getFont()->setBold(true)
            ->setSize(12)
            ->setColor(new Color(Color::COLOR_BLACK));

        // Icon +
        $shape = $slide3->createRichTextShape();
        $shape->setHeight(25)
            ->setWidth(40)
            ->setOffsetX(-5)
            ->setOffsetY(210);
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
            ->setOffsetY(230);
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
            ->setOffsetY(190);
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
            ->setOffsetY(210);
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
            ->setOffsetY(230);
        $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $textRun = $shape->createTextRun($total_closed_this_month);
        $textRun->getFont()->setBold(true)
            ->setSize(12)
            ->setColor(new Color(Color::COLOR_BLACK));



        // -------------------- CHART 1 ---------------------
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
            if ($value->problem == 'Core Surrounding') {
                $color = 'ff89a64e';
            } else if ($value->problem == 'Ekosistem MPC') {
                $color = 'ff00b0f0';
            } else if ($value->problem == 'Loan') {
                $color = 'ffa6a6a6';
            } else if ($value->problem == 'Onboarding') {
                $color = 'ff81ff63';
            } else if ($value->problem == 'Online Payment') {
                $color = 'ff09b1a7';
            } else if ($value->problem == 'Switching 3rdparty') {
                $color = 'ffee52e1';
            } else if ($value->problem == 'Transaction') {
                $color = 'ff8380ee';
            } else if ($value->problem == 'Wholesale') {
                $color = 'ff8064a2';
            } else if ($value->problem == 'Cybersecurity') {
                $color = 'ffb9cd96';
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
        $chartShape->setHeight(215)
            ->setWidth(820)
            ->setOffsetX(25)
            ->setOffsetY(260);
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
            $closed = Data::whereBetween(DB::raw('DATE(created)'), [Carbon::parse($start_date)->subMonths(3), Carbon::parse($end_date)->subMonths(1)])->where('status', '=', 'Closed')->where(DB::raw('MONTH(created)'), '=', $value->month)->where('problem', '!=', 'Enhancement')->get()->count();
            $rcidentified = Data::whereBetween(DB::raw('DATE(created)'), [Carbon::parse($start_date)->subMonths(3), Carbon::parse($end_date)->subMonths(1)])->where('status', '=', 'Root Cause Identified')->where(DB::raw('MONTH(created)'), '=', $value->month)->where('problem', '!=', 'Enhancement')->get()->count();
            $pending = Data::whereBetween(DB::raw('DATE(created)'), [Carbon::parse($start_date)->subMonths(3), Carbon::parse($end_date)->subMonths(1)])->where('status', '=', 'Pending')->where(DB::raw('MONTH(created)'), '=', $value->month)->where('problem', '!=', 'Enhancement')->get()->count();
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
        $chartShape->setHeight(215)
            ->setWidth(820)
            ->setOffsetX(25)
            ->setOffsetY(475);
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


        //Chart 6 Container
        $shape = $slide3->createRichTextShape()
            ->setHeight(215)
            ->setWidth(205)
            ->setOffsetX(845)
            ->setOffsetY(260);
        $shape->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFFFFF'));
        $shape->getBorder()->setLineStyle(Border::LINE_SINGLE)->setColor(new Color('FF000000'));

        // Set Data
        $curr_created = Data::whereYear('created', 2026)
            ->where(DB::raw('DATE(created)'), '<=', $end_date)
            ->where('problem', '!=', 'Enhancement')
            ->count();

        $prev_created = Data::whereYear('created', 2026)
            ->where(DB::raw('DATE(created)'), '<', $start_date)
            ->where('problem', '!=', 'Enhancement')
            ->count();

        $curr_closed = Data::whereYear('closed_time', 2026)
            ->where(DB::raw('DATE(closed_time)'), '<=', $end_date)
            ->where('problem', '!=', 'Enhancement')
            ->where('status', 'Closed')
            ->count();

        $prev_closed = Data::whereYear('closed_time', 2026)
            ->where(DB::raw('DATE(closed_time)'), '<', $start_date)
            ->where('problem', '!=', 'Enhancement')
            ->where('status', 'Closed')
            ->count();
        // dd($curr_created, $prev_created, $curr_closed, $prev_closed);
        $percen_created = $prev_created > 0
            ? (($curr_created - $prev_created) / $prev_created) * 100
            : 0;

        $percen_closed = $prev_closed > 0
            ? (($curr_closed - $prev_closed) / $prev_closed) * 100
            : 0;


        // Menambahkan teks ke kotak pertama
        $percentage = $shape->createTextRun("▲ " . number_format($percen_created, 2) . "%");
        $percentage->getFont()->setBold(true)->setSize(24)->setColor(new Color('FFC00000'));
        $title = $shape->createTextRun("\nIssues Created");
        $title->getFont()->setBold(true)->setSize(20)->setColor(new Color('FFC00000'));
        $c_month = $shape->createTextRun("\n\nCurrent Month : ");
        $c_month->getFont()->setBold(true)->setSize(12);
        $vc_month = $shape->createTextRun("\n" . $curr_created);
        $vc_month->getFont()->setBold(true)->setSize(18);
        $p_month = $shape->createTextRun("\nPrevious Month : ");
        $p_month->getFont()->setBold(true)->setSize(12);
        $vp_month = $shape->createTextRun("\n" . $prev_created);
        $vp_month->getFont()->setBold(true)->setSize(18);

        // Menambahkan kotak kedua untuk "Issues Closed"
        $shape2 = $slide3->createRichTextShape()
            ->setHeight(215)
            ->setWidth(205)
            ->setOffsetX(1050)
            ->setOffsetY(260);
        $shape2->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFFFFF'));
        $shape2->getBorder()->setLineStyle(Border::LINE_SINGLE)->setColor(new Color('FF000000'));

        // Menambahkan teks ke kotak kedua
        $percentage2 = $shape2->createTextRun("▲ " . number_format($percen_closed, 2) . "%");
        $percentage2->getFont()->setBold(true)->setSize(24)->setColor(new Color('FF00C000'));
        $title2 = $shape2->createTextRun("\nIssues Closed");
        $title2->getFont()->setBold(true)->setSize(20)->setColor(new Color('FF00C000'));
        $c_month2 = $shape2->createTextRun("\n\nCurrent Month : ");
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
            ->setOffsetX(845)
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


        // ---------- SLIDE 3 PROBLEM DETAILS -------------- //

        $slide_additional = $objPHPPresentation->createSlide();
        $backgroundImagePath = storage_path('image/background.png');
        $backgroundImage = new File();
        $backgroundImage->setPath($backgroundImagePath);
        $backgroundImage->setWidth(1280);
        $backgroundImage->setOffsetX(0);
        $backgroundImage->setOffsetY(0);
        $slide_additional->addShape($backgroundImage);

        $imagePath = storage_path('image/allobank.png');
        $pictureShape = new File();
        $pictureShape->setPath($imagePath);
        $pictureShape->setWidth(200);
        $pictureShape->setOffsetX(1050);
        $pictureShape->setOffsetY(20);
        $slide_additional->addShape($pictureShape);

        $objPHPPresentation->getLayout()->setDocumentLayout(['cx' => 1280, 'cy' => 700], true)
            ->setCX(1280, DocumentLayout::UNIT_PIXEL)
            ->setCY(700, DocumentLayout::UNIT_PIXEL);

        $shape = $slide_additional->createRichTextShape()
            ->setHeight(50)
            ->setWidth(1000)
            ->setOffsetX(25)
            ->setOffsetY(15);
        $textRun = $shape->createTextRun('Problem Management');
        $textRun->getFont()->setBold(true)
            ->setSize(30);

        $shape = $slide_additional->createRichTextShape()
            ->setHeight(25)
            ->setWidth(400)
            ->setOffsetX(25)
            ->setOffsetY(65);
        $date = Carbon::parse($end_date)->format('F Y');
        $textRun = $shape->createTextRun('As of ' . $date);
        $textRun->getFont()->setSize(14);

        $imagePath = storage_path('image/Line.png');
        $pictureShape = new File();
        $pictureShape->setPath($imagePath);
        $pictureShape->setWidth(1200);  // Ubah ukuran gambar sesuai kebutuhan
        $pictureShape->setOffsetX(20); // Posisi horizontal gambar
        $pictureShape->setOffsetY(100); // Posisi vertikal gambar
        $slide_additional->addShape($pictureShape);


        // SUMMARY ALL PROBLEM
        $shape = $slide_additional->createRichTextShape()
            ->setHeight(25)
            ->setWidth(400)
            ->setOffsetX(25)
            ->setOffsetY(110);
        $textRun = $shape->createTextRun('SUMMARY ALL PROBLEM ON THIS MONTH');
        $textRun->getFont()->setSize(10)->setBold(true);

        // Define data
        $detaildata = Data::where('problem', '!=', 'Enhancement')->where('status', '=', 'Pending')
            ->union(Data::where('problem', '!=', 'Enhancement')->where('status',  '=', 'Root Cause Identified'))
            ->orderByRaw("
        CASE 
            WHEN target_version = '' THEN 3  -- Kosong di paling bawah
            WHEN target_version = 'Backlog' THEN 2  -- Backlog di atas kosong
            ELSE 1  -- Lainnya di atas semua
            END, target_version ASC
            ")
            ->orderByRaw("
            CASE problem
            WHEN 'Loan' THEN 1
            WHEN 'Onboarding' THEN 2
            WHEN 'Core Surrounding' THEN 3
            ELSE 4 
        END
    ")
            ->get();

        // dd(json_encode($detaildata, JSON_PRETTY_PRINT));

        // ----------------- Create Table ------------------------------ 
        $columns = 12;
        $table = $slide_additional->createTableShape($columns);
        $table->getBorder()->setLineStyle(Border::LINE_SINGLE);

        // Set table position & Size
        $table->setheight(210);
        $table->setwidth(1200);
        $table->setOffsetX(25);
        $table->setOffsetY(135);

        $tempdata = [
            ['', 'No', 'Category', 'No Ticket', 'Summary', 'Level', 'Target Version', 'Version Type', 'Team', 'SLA', "Status\nCreated Date", 'Created - RCA Time', 'Ticket Age'],
        ];

        $no = 1;

        foreach ($detaildata as $value) {

            // ===== STATUS =====
            $tempstatus = $value->status === 'Root Cause Identified'
                ? 'RC Identified'
                : $value->status;

            $status = $tempstatus . "\n" . Carbon::parse($value->created)->format('d/m/y');

            // ===== RCA TIME =====
            if ($value->rca_time) {
                $createdDate = Carbon::parse($value->created);
                $rcaDate = Carbon::parse($value->rca_time);
                $rca_days = (int) $createdDate->diffInDays($rcaDate, false);
                $rca_time = $rca_days . " days\n" . Carbon::parse($value->rca_time)->format('d/m/y');
            } else {
                $rca_time = '-';
            }

            // ===== SLA =====
            $priority = strtolower($value->priority);
            $limitMonth = match ($priority) {
                'high' => 2,
                'medium' => 4,
                'low' => 6,
                default => null,
            };

            if ($limitMonth) {
                $slaStatus = Carbon::parse($value->created)
                    ->addMonths($limitMonth)
                    ->lt($value->rca_time ? Carbon::parse($value->rca_time) : Carbon::now())
                    ? '🔴 Over'
                    : '🟢 Met';
            } else {
                $slaStatus = '-';
            }

            // ===== ROW COLOR =====
            $rowColor = match ($value->problem) {
                'Core Surrounding' => 'ff89a64e',
                'Ekosistem MPC' => 'ff00b0f0',
                'Loan' => 'ffa6a6a6',
                'Onboarding' => 'ff81ff63',
                'Online Payment' => 'ff09b1a7',
                'Switching 3rdparty' => 'ffee52e1',
                'Transaction' => 'ff8380ee',
                'Wholesale' => 'ff8064a2',
                'Cybersecurity' => 'ffb9cd96',
                default => 'ffffffff',
            };

            // ===== MAIN ROW =====
            $tempdata[] = [
                $value->problem,
                $no,
                $value->category,
                $value->code_jira,
                $value->summary,
                $value->priority,
                $value->target_version ?? '-',
                $value->version_type ?? '-',
                $value->team ?? '-',
                $slaStatus,
                $status,
                $rca_time,
                Carbon::parse($value->created)->diffForHumans(null, true),
                '__ROWCOLOR__' => $rowColor
            ];

            // ===== RCA ROW =====
            $tempdata[] = [
                'RCA',
                '',
                'RCA',
                '',
                $value->root_cause,
                '',
                '',
                '',
                '',
                '',
                '',
                '',
                '',
                '__ROWCOLOR__' => $rowColor
            ];

            $no++;
        }


        foreach ($tempdata as $rowIndex => $row) {

            $isHeader = ($rowIndex === 0);
            $isRcaRow = ($row[0] === 'RCA');
            $rowColor = $row['__ROWCOLOR__'] ?? 'ffffffff';

            $tableRow = $table->createRow();
            $tableRow->setHeight($isRcaRow ? 45 : 25);

            foreach ($row as $cellIndex => $cellText) {

                if ($cellIndex === 0 || $cellIndex === '__ROWCOLOR__') continue;

                $cell = $tableRow->nextCell();

                // ===== WIDTH =====
                if ($cellIndex == 1) $cell->setWidth(40);
                elseif ($cellIndex == 2) $cell->setWidth(120);
                elseif ($cellIndex == 3) $cell->setWidth(90);
                elseif ($cellIndex == 4) $cell->setWidth(420);
                else $cell->setWidth(70);

                // ===== TEXT =====
                $textRun = $cell->createTextRun($cellText);
                $textRun->getFont()->setBold($isHeader);

                // ===== ALIGN =====
                if ($cellIndex == 4) {
                    $cell->getActiveParagraph()->getAlignment()
                        ->setHorizontal(Alignment::HORIZONTAL_LEFT)
                        ->setMarginLeft(3);
                } else {
                    $cell->getActiveParagraph()->getAlignment()
                        ->setHorizontal(Alignment::HORIZONTAL_CENTER);
                }

                $cell->getActiveParagraph()->getAlignment()
                    ->setVertical(Alignment::VERTICAL_CENTER);

                // ===== HEADER =====
                if ($isHeader) {
                    $cell->getFill()->setFillType(Fill::FILL_SOLID);
                    $cell->getFill()->setStartColor(new Color(Color::COLOR_BLACK));
                    $textRun->getFont()->setColor(new Color(Color::COLOR_WHITE));
                    continue;
                }

                // ===== RCA ROW =====
                if ($isRcaRow) {

                    if ($cellIndex != 4) {
                        $cell->createTextRun('');
                    } else {
                        $textRun->getFont()->setItalic(true)->setSize(10);
                    }

                    $cell->getFill()->setFillType(Fill::FILL_SOLID);
                    $cell->getFill()->setStartColor(new Color($rowColor));
                    continue;
                }

                // ===== NORMAL ROW COLOR =====
                $cell->getFill()->setFillType(Fill::FILL_SOLID);
                $cell->getFill()->setStartColor(new Color($rowColor));
            }
        }


        // ----------- SLIDE CLOSED TICKET ------------------------
        $slideclosedticket = $objPHPPresentation->createSlide();
        $backgroundImagePath = storage_path('image/background.png');
        $backgroundImage = new File();
        $backgroundImage->setPath($backgroundImagePath);
        $backgroundImage->setWidth(1280);
        $backgroundImage->setOffsetX(0);
        $backgroundImage->setOffsetY(0);
        $slideclosedticket->addShape($backgroundImage);


        $imagePath = storage_path('image/allobank.png');
        $pictureShape = new File();
        $pictureShape->setPath($imagePath);
        $pictureShape->setWidth(200);
        $pictureShape->setOffsetX(1050);
        $pictureShape->setOffsetY(20);
        $slideclosedticket->addShape($pictureShape);

        $objPHPPresentation->getLayout()->setDocumentLayout(['cx' => 1280, 'cy' => 700], true)
            ->setCX(1280, DocumentLayout::UNIT_PIXEL)
            ->setCY(700, DocumentLayout::UNIT_PIXEL);

        $shape = $slideclosedticket->createRichTextShape()
            ->setHeight(50)
            ->setWidth(1000)
            ->setOffsetX(25)
            ->setOffsetY(15);
        $textRun = $shape->createTextRun('Problem Management');
        $textRun->getFont()->setBold(true)
            ->setSize(30);

        $shape = $slideclosedticket->createRichTextShape()
            ->setHeight(25)
            ->setWidth(400)
            ->setOffsetX(25)
            ->setOffsetY(65);

        $textRun = $shape->createTextRun('As of ' . $date);
        $textRun->getFont()->setSize(14);

        $imagePath = storage_path('image/Line.png');
        $pictureShape = new File();
        $pictureShape->setPath($imagePath);
        $pictureShape->setWidth(1200);  // Ubah ukuran gambar sesuai kebutuhan
        $pictureShape->setOffsetX(20); // Posisi horizontal gambar
        $pictureShape->setOffsetY(100); // Posisi vertikal gambar
        $slideclosedticket->addShape($pictureShape);

        // SUMMARY ALL PROBLEM CLOSED
        $shape = $slideclosedticket->createRichTextShape()
            ->setHeight(25)
            ->setWidth(400)
            ->setOffsetX(25)
            ->setOffsetY(110);
        $textRun = $shape->createTextRun('DETAIL PROBLEM CLOSED ON THIS MONTH');
        $textRun->getFont()->setSize(10)->setBold(true);

        // Define data
        $detaildata = Data::where('problem', '!=', 'Enhancement')->where('status', '=', 'Closed')->whereBetween(DB::raw('DATE(closed_time)'), [$start_date, $end_date])->get();

        // dd(json_encode($detaildata, JSON_PRETTY_PRINT));

        // ----------------- Create Table ------------------------------ 
        $columns = 4;
        $table = $slideclosedticket->createTableShape($columns);
        $table->getBorder()->setLineStyle(Border::LINE_SINGLE);

        // Set table position & Size
        $table->setheight(210);
        $table->setwidth(1200);
        $table->setOffsetX(25);
        $table->setOffsetY(135);

        // DEFINE ARRAY
        $tempdata = [
            ['', 'Category', 'No Ticket', 'Summary', 'Solution'],
        ];

        // ADD ARRAY DATA
        if ($detaildata->isEmpty()) {
            $tempdata[] = ['-', '-', '-', '-', '-'];
        } else {
            foreach ($detaildata as $key => $value) {
                $tempstatus = $value->status;
                if ($value->status == 'Root Cause Identified') {
                    $tempstatus = 'RC Identified';
                }

                $status = $tempstatus . "\n" . Carbon::parse($value->created)->format('d/m/y');

                $no_ticket = $value->code_jira;
                $summary = $value->summary;
                // $summary = "[" . $value->code_jira . "]" . " " . $value->summary;

                $tempdata[] = [$value->problem, $value->category, $no_ticket, $summary, $value->work_around];
            }
        }


        // INSERT ARRAY TO TABLE
        foreach ($tempdata as $rowIndex => $row) {
            $tableRow = $table->createRow();
            $tableRow->setHeight(25); // Set the height of the row
            foreach ($row as $cellIndex => $cellText) {
                if ($cellIndex == 0) {
                    continue; // Lewati kolom yang disembunyikan
                }

                //set width
                $cell = $tableRow->nextCell();
                if ($cellIndex == 1) {
                    $cell->setWidth(100);
                } else if ($cellIndex == 2) {
                    $cell->setWidth(80);
                } else if ($cellIndex == 3) {
                    $cell->setWidth(520);
                } else if ($cellIndex == 4) {
                    $cell->setWidth(500);
                }

                //set status
                $problem = $row[0];
                $cellText = $cellText ?? '';
                $textRun = $cell->createTextRun($cellText);
                $textRun->getFont()->setBold($rowIndex == 0);
                $cell->getFill()->setFillType(Fill::FILL_SOLID);
                $cell->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
                // kolom 3 & 4 harus left
                if (in_array($cellIndex, [3, 4])) {
                    $cell->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
                    $cell->getActiveParagraph()->getAlignment()->setMarginLeft(2.8);
                } else {
                    $cell->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                }


                if ($rowIndex == 0) {
                    $cell->getFill()->setStartColor(new Color(Color::COLOR_BLACK));
                    $textRun->getFont()->setColor(new Color(Color::COLOR_WHITE));
                } else {
                    // CEK jika ini baris dummy kosong (semua '-')
                    $isEmptyRow = count(array_unique($row)) === 1 && $row[0] === '-';
                    if ($isEmptyRow) {
                        $cell->getFill()->setStartColor(new Color('ffffffff')); // putih
                    } else {
                        if ($problem == 'Core Surrounding') {
                            $cell->getFill()->setStartColor(new Color('ff89a64e'));
                        } else if ($problem == 'Ekosistem MPC') {
                            $cell->getFill()->setStartColor(new Color('ff00b0f0'));
                        } else if ($problem == 'Loan') {
                            $cell->getFill()->setStartColor(new Color('ffa6a6a6'));
                        } else if ($problem == 'Onboarding') {
                            $cell->getFill()->setStartColor(new Color('ff81ff63'));
                        } else if ($problem == 'Online Payment') {
                            $cell->getFill()->setStartColor(new Color('ff09b1a7'));
                        } else if ($problem == 'Switching 3rdparty') {
                            $cell->getFill()->setStartColor(new Color('ffee52e1'));
                        } else if ($problem == 'Transaction') {
                            $cell->getFill()->setStartColor(new Color('ff8380ee'));
                        } else if ($problem == 'Wholesale') {
                            $cell->getFill()->setStartColor(new Color('ff8064a2'));
                        } else if ($problem == 'Cybersecurity') {
                            $cell->getFill()->setStartColor(new Color('ffb9cd96'));
                        } else {
                            $cell->getFill()->setStartColor(new Color('ffffffff'));
                        }
                    }
                }
            }
        }


        // ----------- SLIDE ENHANCEMENT ------------------------
        $slideEnhancement = $objPHPPresentation->createSlide();
        $backgroundImagePath = storage_path('image/background.png');
        $backgroundImage = new File();
        $backgroundImage->setPath($backgroundImagePath);
        $backgroundImage->setWidth(1280);
        $backgroundImage->setOffsetX(0);
        $backgroundImage->setOffsetY(0);
        $slideEnhancement->addShape($backgroundImage);


        $imagePath = storage_path('image/allobank.png');
        $pictureShape = new File();
        $pictureShape->setPath($imagePath);
        $pictureShape->setWidth(200);
        $pictureShape->setOffsetX(1050);
        $pictureShape->setOffsetY(20);
        $slideEnhancement->addShape($pictureShape);

        $objPHPPresentation->getLayout()->setDocumentLayout(['cx' => 1280, 'cy' => 700], true)
            ->setCX(1280, DocumentLayout::UNIT_PIXEL)
            ->setCY(700, DocumentLayout::UNIT_PIXEL);

        $shape = $slideEnhancement->createRichTextShape()
            ->setHeight(50)
            ->setWidth(1000)
            ->setOffsetX(25)
            ->setOffsetY(15);
        $textRun = $shape->createTextRun('Problem Management');
        $textRun->getFont()->setBold(true)
            ->setSize(30);

        $shape = $slideEnhancement->createRichTextShape()
            ->setHeight(25)
            ->setWidth(400)
            ->setOffsetX(25)
            ->setOffsetY(65);
        $startdate = Carbon::parse($start_date)->format('d F Y');
        $enddate = Carbon::parse($end_date)->format('d F Y');
        $textRun = $shape->createTextRun('As of ' . $date);
        $textRun->getFont()->setSize(14);

        $shape = $slideEnhancement->createRichTextShape()
            ->setHeight(25)
            ->setWidth(400)
            ->setOffsetX(25)
            ->setOffsetY(110);
        $textRun = $shape->createTextRun('PRODUCT ENHANCEMENT (STRENGTHEN)');
        $textRun->getFont()->setSize(10)->setBold(true);

        $imagePath = storage_path('image/Line.png');
        $pictureShape = new File();
        $pictureShape->setPath($imagePath);
        $pictureShape->setWidth(1200);  // Ubah ukuran gambar sesuai kebutuhan
        $pictureShape->setOffsetX(20); // Posisi horizontal gambar
        $pictureShape->setOffsetY(100); // Posisi vertikal gambar
        $slideEnhancement->addShape($pictureShape);

        //TABLE OPEN PROBLEM ENHANCEMENT
        $columns = 12; // Number of columns
        $tableShape = $slideEnhancement->createTableShape($columns);
        $tableShape->getBorder()->setLineStyle(Border::LINE_SINGLE);

        // Set the table's position and size
        $tableShape->setHeight(210);
        $tableShape->setWidth(1030);
        $tableShape->setOffsetX(25);
        $tableShape->setOffsetY(135);

        // GET DATA FROM DATABASE
        $data_table = Data::
            // whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])
            where('problem', '=', 'Enhancement')
            ->whereIn('status', ['Pending', 'Root Cause Identified'])
            ->select('code_jira', 'problem', 'category', 'summary', 'status', 'created', 'target_version', 'target_date', 'priority', 'changed_at', 'rca_time', 'closed_time', 'team', 'aspect')
            ->orderBy('created', 'ASC')
            ->get();

        // DEFINE ARRAY
        $tempdata = [
            ['', 'No', 'Category', 'No Ticket', 'Summary', 'Created Date', 'Target Version', 'Version Type', 'Target Date', 'Level', 'Team', 'Aspect', 'Status'],
        ];

        // ADD ARRAY DATA
        $i = 1;
        foreach ($data_table as $key => $value) {
            $status = $value->status;
            if ($value->status == 'Root Cause Identified') {
                $status = 'RC Identified';
            }

            $summary = $value->summary;
            $no_ticket = $value->code_jira;

            //convert date to carbon parse
            $created = Carbon::parse($value->created);
            $rcatime = Carbon::parse($value->rca_time);
            $closed_time = Carbon::parse($value->closed_time);

            $target_version = $value->target_version;

            //declare rca time
            if ($value->rca_time == null) {
                $rca_time = '-';
            } else {
                $rca_days = intval($created->diffInDays($rcatime));
                $rca_days_string = strval($rca_days) . ' days';
                $rca_time = $rca_days_string . "\n" . Carbon::parse($value->rca_time)->format('d/m/y');
            }

            //declare team
            if ($value->team == null) {
                $team = '-';
            } else {
                $team = $value->team;
            }

            //declare completion time
            if ($value->closed_time == null) {
                $completion_time = '-';
            } else {
                $completion_days = intval($created->diffInDays($closed_time));
                $completion_days_string = strval($completion_days) . ' Days';
                $completion_time = $completion_days_string . "\n" . Carbon::parse($value->closed_time)->format('d/m/y');
            }

            //declare target date
            if ($value->target_date == null) {
                $target_date = '-';
            } else {
                $target_date = Carbon::parse($value->target_date)->format('d/m/y');
            }

            if ($value->aspect == null) {
                $aspect = 'Others';
            } else {
                $aspect = $value->aspect;
            }

            $tempdata[] = [$value->problem, strval($i), $value->category, $no_ticket, $summary,  $created->format('d/m/y'), $target_version, $target_version, $target_date, $value->priority,  $team, $aspect, $status];
            $i++;
        }

        // dd($tempdata);

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
                    $cell->setWidth(30);
                } else if ($cellIndex == 2) {
                    $cell->setWidth(80);
                } else if ($cellIndex == 3) {
                    $cell->setWidth(80);
                } else if ($cellIndex == 4) {
                    $cell->setWidth(290);
                } else if ($cellIndex == 5) {
                    $cell->setWidth(70);
                } else if ($cellIndex == 6) {
                    $cell->setWidth(70);
                } else if ($cellIndex == 7) {
                    $cell->setWidth(70);
                } else if ($cellIndex == 8) {
                    $cell->setWidth(70);
                } else if ($cellIndex == 9) {
                    $cell->setWidth(70);
                } else if ($cellIndex == 10) {
                    $cell->setWidth(70);
                } else if ($cellIndex == 11) {
                    $cell->setWidth(70);
                } else if ($cellIndex == 12) {
                    $cell->setWidth(60);
                }

                //set status
                $problem = $row[0];
                $status = explode("\n", $row[12]);
                $firstStatus = $status[0];
                // $cell = $tableRow->nextCell();
                $textRun = $cell->createTextRun($cellText);
                $textRun->getFont()->setBold($rowIndex == 0);
                $cell->getFill()->setFillType(Fill::FILL_SOLID);
                if ($cellIndex == 4) { // jangan override untuk kolom ke-4
                    $cell->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
                    $cell->getActiveParagraph()->getAlignment()->setMarginLeft(2.8);
                } else {
                    $cell->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                }
                // vertical center
                $cell->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
                $cell->getFill()->setStartColor(new Color('ffffffff'));
                //
                if ($rowIndex == 0) {
                    $cell->getFill()->setStartColor(new Color(Color::COLOR_BLACK));
                    $textRun->getFont()->setColor(new Color(Color::COLOR_WHITE));
                } else {
                    if ($cellIndex == 12) {
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
                        $cell->getFill()->setStartColor(new Color('ffffffff'));
                    }
                }
            }
        }

        //TABLE CLOSED ENHANCEMENT
        $columns = 12; // Number of columns
        $tableShape = $slideEnhancement->createTableShape($columns);
        $tableShape->getBorder()->setLineStyle(Border::LINE_SINGLE);

        $tableShape->setHeight(210);
        $tableShape->setWidth(1030);
        $tableShape->setOffsetX(25);
        $tableShape->setOffsetY(400);

        //get lastdate
        $last_date = Carbon::parse($end_date)->endOfDay();
        // GET DATA FROM DATABASE
        $data_table = Data::where('problem', '=', 'Enhancement')
            ->whereBetween('closed_time', [$start_date, $last_date])
            ->where('status', '=', 'Closed')
            ->select('code_jira', 'problem', 'category', 'summary', 'status', 'created', 'target_version', 'target_date', 'priority', 'changed_at', 'rca_time', 'closed_time', 'team', 'aspect')
            ->get();

        // DEFINE ARRAY
        $tempdata = [
            ['', 'No', 'Category', 'No Ticket', 'Summary', 'Created Date', 'Target Version', 'Version Type', 'Target Date', 'Level', 'Team', 'Aspect', 'Status'],
        ];

        // ADD ARRAY DATA
        $i = 1;
        foreach ($data_table as $key => $value) {
            $status = $value->status;
            if ($value->status == 'Root Cause Identified') {
                $status = 'RC Identified';
            }

            $summary = $value->summary;
            $no_ticket = $value->code_jira;

            //convert date to carbon parse
            $created = Carbon::parse($value->created);
            $rcatime = Carbon::parse($value->rca_time);
            $closed_time = Carbon::parse($value->closed_time);
            $target_version = $value->target_version;

            //declare rca time
            if ($value->rca_time == null) {
                $rca_time = '-';
            } else {
                $rca_days = intval($created->diffInDays($rcatime));
                $rca_days_string = strval($rca_days) . ' days';
                $rca_time = $rca_days_string . "\n" . Carbon::parse($value->rca_time)->format('d/m/y');
            }

            //declare team
            if ($value->team == null) {
                $team = '-';
            } else {
                $team = $value->team;
            }

            //declare completion time
            if ($value->closed_time == null) {
                $completion_time = '-';
            } else {
                $completion_days = intval($created->diffInDays($closed_time));
                $completion_days_string = strval($completion_days) . ' Days';
                $completion_time = $completion_days_string . "\n" . Carbon::parse($value->closed_time)->format('d/m/y');
            }

            //declare target date
            if ($value->target_date == null) {
                $target_date = '-';
            } else {
                $target_date = Carbon::parse($value->target_date)->format('d/m/y');
            }

            //aspect
            if ($value->aspect == null) {
                $aspect = 'Others';
            } else {
                $aspect = $value->aspect;
            }

            $tempdata[] = [$value->problem, strval($i), $value->category, $no_ticket, $summary,  $created->format('d/m/y'), $target_version, $target_version, $target_date, $value->priority,  $team, $aspect, $status];
            $i++;
        }

        // INSERT ARRAY TO TABLE
        foreach ($tempdata as $rowIndex => $row) {
            $tableRow = $tableShape->createRow();
            $tableRow->setHeight(25); // Set the height of the row
            foreach ($row as $cellIndex => $cellText) {
                if ($cellIndex == 0) {
                    continue; // Lewati kolom yang disembunyikan
                }
                $cell = $tableRow->nextCell();
                if ($cellIndex == 1) {
                    $cell->setWidth(30);
                } else if ($cellIndex == 2) {
                    $cell->setWidth(80);
                } else if ($cellIndex == 3) {
                    $cell->setWidth(80);
                } else if ($cellIndex == 4) {
                    $cell->setWidth(290);
                } else if ($cellIndex == 5) {
                    $cell->setWidth(70);
                } else if ($cellIndex == 6) {
                    $cell->setWidth(70);
                } else if ($cellIndex == 7) {
                    $cell->setWidth(70);
                } else if ($cellIndex == 8) {
                    $cell->setWidth(70);
                } else if ($cellIndex == 9) {
                    $cell->setWidth(70);
                } else if ($cellIndex == 10) {
                    $cell->setWidth(70);
                } else if ($cellIndex == 11) {
                    $cell->setWidth(70);
                } else if ($cellIndex == 12) {
                    $cell->setWidth(60);
                }

                $problem = $row[0];
                $status = explode("\n", $row[12]);
                $firstStatus = $status[0];
                $textRun = $cell->createTextRun($cellText);
                $textRun->getFont()->setBold($rowIndex == 0);
                $cell->getFill()->setFillType(Fill::FILL_SOLID);
                if ($cellIndex == 4) { // jangan override untuk kolom ke-4
                    $cell->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
                    $cell->getActiveParagraph()->getAlignment()->setMarginLeft(2.8);
                } else {
                    $cell->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                }
                // vertical center
                $cell->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
                $cell->getFill()->setStartColor(new Color('ffffffff'));
                //
                if ($rowIndex == 0) {
                    $cell->getFill()->setStartColor(new Color(Color::COLOR_BLACK));
                    $textRun->getFont()->setColor(new Color(Color::COLOR_WHITE));
                } else {
                    if ($cellIndex == 12) {
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
                        $cell->getFill()->setStartColor(new Color('ffffffff'));
                    }
                }
            }
        }


        // Detail High, Medium, Low Enhancement
        $high_lastweek_enhancement = Data::where(DB::raw('DATE(created)'), '<', $start_date)
            ->where('problem', '=', 'Enhancement')
            ->where('priority', '=', 'High')
            ->whereIn('status', ['Root Cause Identified', 'Pending'])
            ->union(Data::where(DB::raw('DATE(created)'), '<', $start_date)
                ->whereBetween(DB::raw('DATE(closed_time)'), [$start_date, $end_date])
                ->where('problem', '=', 'Enhancement')
                ->where('priority', '=', 'High')
                ->where('status', '=', 'Closed'))
            ->count();

        $medium_lastweek_enhancement = Data::where(DB::raw('DATE(created)'), '<', $start_date)
            ->where('problem', '=', 'Enhancement')
            ->where('priority', '=', 'Medium')
            ->whereIn('status', ['Root Cause Identified', 'Pending'])
            ->union(Data::where(DB::raw('DATE(created)'), '<', $start_date)
                ->whereBetween(DB::raw('DATE(closed_time)'), [$start_date, $end_date])
                ->where('problem', '=', 'Enhancement')
                ->where('priority', '=', 'Medium')
                ->where('status', '=', 'Closed'))
            ->count();

        $low_lastweek_enhancement = Data::where(DB::raw('DATE(created)'), '<', $start_date)
            ->where('problem', '=', 'Enhancement')
            ->where('priority', '=', 'Low')
            ->whereIn('status', ['Root Cause Identified', 'Pending'])
            ->union(Data::where(DB::raw('DATE(created)'), '<', $start_date)
                ->whereBetween(DB::raw('DATE(closed_time)'), [$start_date, $end_date])
                ->where('problem', '=', 'Enhancement')
                ->where('priority', '=', 'Low')
                ->where('status', '=', 'Closed'))
            ->count();

        $high_thisweek_enhancement = Data::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])
            ->where('problem', '=', 'Enhancement')
            ->where('priority', '=', 'High')
            ->count();

        $medium_thisweek_enhancement = Data::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])
            ->where('problem', '=', 'Enhancement')
            ->where('priority', '=', 'Medium')
            ->count();

        $low_thisweek_enhancement = Data::whereBetween(DB::raw('DATE(created)'), [$start_date, $end_date])
            ->where('problem', '=', 'Enhancement')
            ->where('priority', '=', 'Low')
            ->count();

        $high_closed_thisweek_enhancement = Data::whereBetween(DB::raw('DATE(changed_at)'), [$start_date, $end_date])
            ->where('problem', '=', 'Enhancement')
            ->where('priority', '=', 'High')
            ->where('status', '=', 'Closed')
            ->count();

        $medium_closed_thisweek_enhancement = Data::whereBetween(DB::raw('DATE(changed_at)'), [$start_date, $end_date])
            ->where('problem', '=', 'Enhancement')
            ->where('priority', '=', 'Medium')
            ->where('status', '=', 'Closed')
            ->count();

        $low_closed_thisweek_enhancement = Data::whereBetween(DB::raw('DATE(changed_at)'), [$start_date, $end_date])
            ->where('problem', '=', 'Enhancement')
            ->where('priority', '=', 'Low')
            ->where('status', '=', 'Closed')
            ->count();

        // Count Enhancement
        $enhancement_high = $high_lastweek_enhancement + $high_thisweek_enhancement - $high_closed_thisweek_enhancement;
        $enhancement_medium = $medium_lastweek_enhancement + $medium_thisweek_enhancement - $medium_closed_thisweek_enhancement;
        $enhancement_low = $low_lastweek_enhancement + $low_thisweek_enhancement - $low_closed_thisweek_enhancement;
        $enhancement_count = $enhancement_high + $enhancement_medium + $enhancement_low;

        // Counting existing, this week, closed
        $total_existing_enhancement = $low_lastweek_enhancement + $medium_lastweek_enhancement + $high_lastweek_enhancement;
        $total_thisweek_enhancement = $low_thisweek_enhancement + $medium_thisweek_enhancement + $high_thisweek_enhancement;
        $total_closed_enhancement = $low_closed_thisweek_enhancement + $medium_closed_thisweek_enhancement + $high_closed_thisweek_enhancement;

        $tableShape = $slideEnhancement->createTableShape(3);
        $tableShape->setHeight(100);
        $tableShape->setWidth(144);
        $tableShape->setOffsetX(1100);
        $tableShape->setOffsetY(135);

        //row judul
        $rowShape = $tableShape->createRow();
        $rowShape->setHeight(40);
        $cell = $rowShape->nextCell();
        $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFFFFFFF'));
        $cell->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $cell->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
        $cell->setColSpan(3);
        $textRun = $cell->createTextRun($enhancement_count . "\n" . truncateString('Enhancement'));
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
            $high_lastweek_enhancement,
            $medium_lastweek_enhancement,
            $low_lastweek_enhancement
        ];
        foreach ($value as $key => $v) {
            $cell = $rowShape->nextCell();
            $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFFFFFFF'));
            $cell->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            $cell->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
            $cell->createTextRun($v);
        }

        $rowShape = $tableShape->createRow();
        $rowShape->setHeight(20);
        $value = [
            $high_thisweek_enhancement,
            $medium_thisweek_enhancement,
            $low_thisweek_enhancement
        ];

        foreach ($value as $key => $v) {
            $cell = $rowShape->nextCell();
            $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFFFFFFF'));
            $cell->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            $cell->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
            $cell->createTextRun($v);
        }

        $rowShape = $tableShape->createRow();
        $rowShape->setHeight(20);
        $value = [
            $high_closed_thisweek_enhancement,
            $medium_closed_thisweek_enhancement,
            $low_closed_thisweek_enhancement
        ];

        foreach ($value as $key => $v) {
            $cell = $rowShape->nextCell();
            $cell->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFFFFFFF'));
            $cell->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            $cell->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
            $cell->createTextRun($v);
        }

        // Icon +
        $shape = $slideEnhancement->createRichTextShape();
        $shape->setHeight(25)
            ->setWidth(40)
            ->setOffsetX(1070)
            ->setOffsetY(210);
        $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $textRun = $shape->createTextRun('+');
        $textRun->getFont()->setBold(true)
            ->setSize(16)
            ->setColor(new Color(Color::COLOR_BLACK));

        // Icon -
        $shape = $slideEnhancement->createRichTextShape();
        $shape->setHeight(25)
            ->setWidth(40)
            ->setOffsetX(1070)
            ->setOffsetY(230);
        $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $textRun = $shape->createTextRun('-');
        $textRun->getFont()->setBold(true)
            ->setSize(16)
            ->setColor(new Color(Color::COLOR_BLACK));

        // Total Existing
        $shape = $slideEnhancement->createRichTextShape();
        $shape->setHeight(25)
            ->setWidth(40)
            ->setOffsetX(1235)
            ->setOffsetY(190);
        $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $textRun = $shape->createTextRun($total_existing_enhancement);
        $textRun->getFont()->setBold(true)
            ->setSize(12)
            ->setColor(new Color(Color::COLOR_BLACK));

        //Total Created
        $shape = $slideEnhancement->createRichTextShape();
        $shape->setHeight(25)
            ->setWidth(40)
            ->setOffsetX(1235)
            ->setOffsetY(210);
        $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $textRun = $shape->createTextRun($total_thisweek_enhancement);
        $textRun->getFont()->setBold(true)
            ->setSize(12)
            ->setColor(new Color(Color::COLOR_BLACK));

        //Total Closed
        $shape = $slideEnhancement->createRichTextShape();
        $shape->setHeight(25)
            ->setWidth(40)
            ->setOffsetX(1235)
            ->setOffsetY(230);
        $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $textRun = $shape->createTextRun($total_closed_enhancement);
        $textRun->getFont()->setBold(true)
            ->setSize(12)
            ->setColor(new Color(Color::COLOR_BLACK));

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

        $imagePath = storage_path('image/Line.png');
        $pictureShape = new File();
        $pictureShape->setPath($imagePath);
        $pictureShape->setWidth(1200);  // Ubah ukuran gambar sesuai kebutuhan
        $pictureShape->setOffsetX(20); // Posisi horizontal gambar
        $pictureShape->setOffsetY(100); // Posisi vertikal gambar
        $additionalslide->addShape($pictureShape);

        $objPHPPresentation->getLayout()->setDocumentLayout(['cx' => 1280, 'cy' => 700], true)
            ->setCX(1280, DocumentLayout::UNIT_PIXEL)
            ->setCY(700, DocumentLayout::UNIT_PIXEL);

        $shape = $additionalslide->createRichTextShape()
            ->setHeight(50)
            ->setWidth(1000)
            ->setOffsetX(25)
            ->setOffsetY(15);
        $textRun = $shape->createTextRun('Problem Management');
        $textRun->getFont()->setBold(true)
            ->setSize(30);

        $shape = $additionalslide->createRichTextShape()
            ->setHeight(25)
            ->setWidth(400)
            ->setOffsetX(25)
            ->setOffsetY(65);
        $textRun = $shape->createTextRun('As of ' . $date);
        $textRun->getFont()->setSize(14);

        $shape = $additionalslide->createRichTextShape()
            ->setHeight(25)
            ->setWidth(400)
            ->setOffsetX(25)
            ->setOffsetY(110);
        $textRun = $shape->createTextRun('IT PROBLEM CHART');
        $textRun->getFont()->setSize(10)->setBold(true);

        // -------------- SET CHART RCA TIME --------------------------

        // Define data
        $days1 = Data::where('created', '>=', Carbon::now()->subMonth()->format('Y-m-d'))
            ->whereNotNull('rca_time')
            ->where('rca_days', '=', 1)
            ->count();
        $days2 = Data::where('created', '>=', Carbon::now()->subMonth()->format('Y-m-d'))
            ->whereNotNull('rca_time')
            ->where('rca_days', '=', 2)
            ->count();
        $days3 = Data::where('created', '>=', Carbon::now()->subMonth()->format('Y-m-d'))
            ->whereNotNull('rca_time')
            ->where('rca_days', '=', 3)
            ->count();
        $days4 = Data::where('created', '>=', Carbon::now()->subMonth()->format('Y-m-d'))
            ->whereNotNull('rca_time')
            ->where('rca_days', '=', 4)
            ->count();
        $days5 = Data::where('created', '>=', Carbon::now()->subMonth()->format('Y-m-d'))
            ->whereNotNull('rca_time')
            ->where('rca_days', '=', 5)
            ->count();
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
            ->setHeight(250)
            ->setWidth(600)
            ->setOffsetX(25)
            ->setOffsetY(135);
        // Set judul chart
        $shape->getTitle()->setText('Ticket RCA Time');
        $shape->getLegend()->getBorder()->setLineStyle(Border::LINE_NONE); // Menghilangkan kotak pada legenda
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

        // ------------ DETAIL LIST RCA TIME TICKET ------------------
        $columns = 6; // Number of columns
        $tableShape = $additionalslide->createTableShape($columns);
        $tableShape->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $tableShape->setHeight(275);
        $tableShape->setWidth(600);
        $tableShape->setOffsetX(25);
        $tableShape->setOffsetY(385);

        // QUERY
        $data_table = Data::where('created', '>=', Carbon::now()->subMonth()->format('Y-m-d'))
            ->whereNotNull('rca_time')
            ->Orderby('rca_days', 'desc')
            ->get();

        // DEFINE ARRAY
        $tempdata = [
            ['', 'Category', 'Summary', 'Created Date', 'Created-RCA Time', 'Resolved Time', 'Status & Complete Time'],
        ];

        // ADD ARRAY DATA
        foreach ($data_table as $key => $value) {
            $tempstatus = $value->status;
            if ($value->status == 'Root Cause Identified') {
                $tempstatus = 'RC Iden';
            }

            if ($value->status == 'Closed') {
                $status = $tempstatus . "\n" . Carbon::parse($value->changed_at)->format('d/m/y');
            } else {
                $status = $tempstatus . "\n" . '-';
            }

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

            $tempdata[] = [$value->problem, $value->category, $summary,  $created->format('d/m/y'), $rca_time,  $completion_time, $status];
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
                    $cell->setWidth(80);
                } else if ($cellIndex == 2) {
                    $cell->setWidth(240);
                } else if ($cellIndex == 3) {
                    $cell->setWidth(70);
                } else if ($cellIndex == 4) {
                    $cell->setWidth(70);
                } else if ($cellIndex == 5) {
                    $cell->setWidth(70);
                } else if ($cellIndex == 6) {
                    $cell->setWidth(70);
                }

                //set status
                $problem = $row[0];
                $status = explode("\n", $row[6]);
                $firstStatus = $status[0];
                $textRun = $cell->createTextRun($cellText);
                $textRun->getFont()->setSize(8);
                $textRun->getFont()->setBold($rowIndex == 0);
                $cell->getFill()->setFillType(Fill::FILL_SOLID);
                if ($cellIndex == 2) { // jangan override untuk kolom ke-4
                    $cell->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
                    $cell->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
                    $cell->getActiveParagraph()->getAlignment()->setMarginLeft(2.8);
                } else {
                    $cell->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                    $cell->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
                }
                if ($rowIndex == 0) {
                    $cell->getFill()->setStartColor(new Color(Color::COLOR_BLACK));
                    $textRun->getFont()->setColor(new Color(Color::COLOR_WHITE));
                } else {
                    $cell->getFill()->setStartColor(new Color('ff95b3d7'));
                }
            }
        }

        // ================== CHART PROBLEM BY ASSIGNEE & STATUS ==================

        // --------------------------------------------------
        // 1. Ambil daftar assignee
        // --------------------------------------------------
        $assignees = Data::select('nickname')
            ->where('problem', '!=', 'Enhancement')
            ->groupBy('nickname')
            ->pluck('nickname');

        // --------------------------------------------------
        // 2. Mapping status & warna
        // --------------------------------------------------
        $statusMap = [
            'Closed' => [
                'field' => 'closed_time',
                'color' => 'FF00B050' // Hijau
            ],
            'Root Cause Identified' => [
                'field' => 'created',
                'color' => 'FFFFC000' // Oranye
            ],
            'Pending' => [
                'field' => 'created',
                'color' => 'FFFFFF00' // Kuning
            ],
        ];

        // --------------------------------------------------
        // 3. Hitung data per assignee & status
        // --------------------------------------------------
        $chartData = [];

        foreach ($assignees as $nickname) {
            foreach ($statusMap as $status => $config) {
                $count = Data::whereBetween(DB::raw('DATE(' . $config['field'] . ')'), [$start_date, $end_date])
                    ->where('nickname', $nickname)
                    ->where('status', $status)
                    ->where('problem', '!=', 'Enhancement')
                    ->count();

                if ($count > 0) {
                    $chartData[$status][$nickname] = $count;
                }
            }
        }

        // --------------------------------------------------
        // 4. Buat chart
        // --------------------------------------------------
        $chartShape = $additionalslide->createChartShape();
        $chartShape->setHeight(250)
            ->setWidth(600)
            ->setOffsetX(625)
            ->setOffsetY(135);

        // Tipe chart
        $chartType = new Bar();
        $chartShape->getPlotArea()->setType($chartType);

        // Judul
        $chartShape->getTitle()->setText('Problem By Assignee & Status');

        // Legend
        $chartShape->getLegend()->getBorder()->setLineStyle(Border::LINE_NONE);

        // Axis
        $chartShape->getPlotArea()->getAxisX()->setTitle('');
        $chartShape->getPlotArea()->getAxisY()->setTitle('');

        // Border chart
        $chartShape->getBorder()
            ->setLineStyle(Border::LINE_SINGLE)
            ->setLineWidth(1)
            ->setColor(new Color('FF000000'));

        // --------------------------------------------------
        // 5. Tambahkan series (status → warna)
        // --------------------------------------------------
        foreach ($statusMap as $status => $config) {
            if (empty($chartData[$status])) {
                continue;
            }

            $series = new Series($status, $chartData[$status]);
            $series->getFill()
                ->setFillType(Fill::FILL_SOLID)
                ->setStartColor(new Color($config['color']));

            $chartType->addSeries($series);
        }

        // ================== CHART JIRA SERVICE REQUEST (CLEAN) ==================

        // 1️⃣ Ambil data sekali (aggregate)
        $services = Service::whereBetween(
            DB::raw('DATE(created)'),
            [$start_date, $end_date]
        )
            ->select(
                'issue_type',
                'status',
                DB::raw('COUNT(*) as total')
            )
            ->groupBy('issue_type', 'status')
            ->get();

        // 2️⃣ Ambil status dan urutkan berdasarkan total terbanyak
        $statuses = $services
            ->groupBy('status')
            ->map(fn($items) => $items->sum('total'))
            ->sortDesc()
            ->keys()
            ->values()
            ->toArray();

        // 3️⃣ Ambil issue_type dan urutkan berdasarkan total terbanyak
        $issueTypes = $services
            ->groupBy('issue_type')
            ->map(fn($items) => $items->sum('total'))
            ->sortDesc()
            ->keys()
            ->values();

        // 4️⃣ Build series data (issue_type sebagai series)
        $seriesData = [];

        foreach ($issueTypes as $issueType) {
            $data = [];

            foreach ($statuses as $status) {
                $data[$status] = $services
                    ->where('issue_type', $issueType)
                    ->where('status', $status)
                    ->sum('total');
            }

            // hanya tambahkan kalau ada datanya
            if (array_sum($data) > 0) {
                $seriesData[$issueType] = $data;
            }
        }

        // ================== GENERATE CHART ==================

        $chartShape = $additionalslide->createChartShape();
        $chartShape->setHeight(250)
            ->setWidth(600)
            ->setOffsetX(625)
            ->setOffsetY(385);

        // Chart type
        $chartType = new Bar();
        $chartShape->getPlotArea()->setType($chartType);

        // Title
        $chartShape->getTitle()->setText('Ticket Jira Service Request');

        // Axis & styling
        $chartShape->getLegend()->getBorder()->setLineStyle(Border::LINE_NONE);
        $chartShape->getPlotArea()->getAxisX()->setTitle('');
        $chartShape->getPlotArea()->getAxisY()->setTitle('');
        $chartShape->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $chartShape->getBorder()->setColor(new Color('FF000000'));
        $chartShape->getBorder()->setLineWidth(1);

        // Add series ke chart
        foreach ($seriesData as $issueType => $statusCounts) {
            $chartType->addSeries(
                new Series($issueType, $statusCounts)
            );
        }


        // ------------------------------------------------------------------------------------

        //Slide 5
        $slide5 = $objPHPPresentation->createSlide();
        $backgroundImagePath = storage_path('image/background_end.png');
        $backgroundImage = new File();
        $backgroundImage->setPath($backgroundImagePath);
        $backgroundImage->setWidth(1280);
        $backgroundImage->setHeight(723);
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
