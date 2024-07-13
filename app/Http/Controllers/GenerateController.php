<?php

namespace App\Http\Controllers;

use App\Models\Data;
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

use Exception;

class GenerateController extends Controller
{
    public function __construct()
    {
        $this->middleware('auth');
    }

    public function index()
    {

        return view('generate_ppt');
    }

    public function generateppt(Request $request)
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
            ->setOffsetX(50)
            ->setOffsetY(25);
        $textRun = $shape->createTextRun('Report IT Problem');
        $textRun->getFont()->setBold(true)
            ->setSize(30);

        $shape = $slide3->createRichTextShape()
            ->setHeight(25)
            ->setWidth(400)
            ->setOffsetX(50)
            ->setOffsetY(75);
        $date = Carbon::parse($end_date)->format('d F Y');
        $textRun = $shape->createTextRun('As of ' . $date);
        $textRun->getFont()->setSize(14);

        //data container category
        $problem = Data::whereBetween('created', [$start_date, $end_date])->select('problem', DB::raw('count(*) as count'))->groupBy('problem')->get();
        // dd($problem);
        $total = [];
        foreach ($problem as $key => $value) {
            $highest = Data::whereBetween('created', [$start_date, $end_date])->where('problem', '=', $value->problem)->where('priority', '=', 'Highest')->get()->count();
            $high = Data::whereBetween('created', [$start_date, $end_date])->where('problem', '=', $value->problem)->where('priority', '=', 'High')->get()->count();
            $medium = Data::whereBetween('created', [$start_date, $end_date])->where('problem', '=', $value->problem)->where('priority', '=', 'Medium')->get()->count();
            $low = Data::whereBetween('created', [$start_date, $end_date])->where('problem', '=', $value->problem)->where('priority', '=', 'Low')->get()->count();
            $lowest = Data::whereBetween('created', [$start_date, $end_date])->where('problem', '=', $value->problem)->where('priority', '=', 'Lowest')->get()->count();
            $total[] = [
                'problem' => $value->problem,
                'total' => $value->count,
                'high' => $highest + $high,
                'medium' => $medium,
                'low' => $low + $lowest,
            ];
        }

        // dd($total);

        //set tempat
        $offsetx = 50;
        $offsety = 120;
        //loop category data
        foreach ($total as $key => $data) {
            // Tambahkan tabel dengan 4 baris dan 3 kolom
            $tableShape = $slide3->createTableShape(3);
            $tableShape->setHeight(100);
            $tableShape->setWidth(135);
            $tableShape->setOffsetX($offsetx);
            $tableShape->setOffsetY($offsety);

            //row judul
            $rowShape = $tableShape->createRow();
            $rowShape->setHeight(40);
            $cell = $rowShape->nextCell();
            $cell->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            $cell->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
            $cell->setColSpan(3);
            $textRun = $cell->createTextRun($data['problem'] . ' ' . $data['total']);
            $textRun->getFont()->setBold(true);
            $textRun->getFont()->setSize(12);

            //row title
            $rowShape = $tableShape->createRow();
            $rowShape->setHeight(25);
            $value = ['High', 'Med', 'Low'];
            foreach ($value as $key => $v) {
                $cell = $rowShape->nextCell();
                $cell->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                $cell->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
                $textRun = $cell->createTextRun($v);
                $textRun->getFont()->setBold(true);
            }

            $rowShape = $tableShape->createRow();
            $rowShape->setHeight(25);
            $value = [$data['high'], $data['medium'], $data['low']];
            foreach ($value as $key => $v) {
                $cell = $rowShape->nextCell();
                $cell->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                $cell->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
                $cell->createTextRun($v);
            }

            //set tempat box selanjutnya
            $offsetx = $offsetx + 145;
        }

        //set data chart 1
        $data_chart1 = Data::where('created', '<=', $end_date)->select('problem', DB::raw('count(*) as count'))->groupBy('problem')->get();
        $resultdata_chart1 = [];
        foreach ($data_chart1 as $key => $value) {
            $status_closed = Data::where('created', '<=', $end_date)->where('problem', '=', $value->problem)->where('status', '=', 'Closed')->get()->count();
            $status_pending = Data::where('created', '<=', $end_date)->where('problem', '=', $value->problem)->where('status', '=', 'Pending')->get()->count();
            $resultdata_chart1[] =
                [
                    'problem' => $value->problem,
                    'total' => $value->count,
                    'count_closed' => $status_closed,
                    'count_pending' => $status_pending,
                ];
        }

        // Chart 1 
        $chartShape = $slide3->createChartShape();
        $chartShape->setHeight(400)
            ->setWidth(600)
            ->setOffsetX(20)
            ->setOffsetY(250);
        // Define tipe chart
        $chartType = new Bar();
        $chartShape->getPlotArea()->setType($chartType);
        // Set judul chart
        $chartShape->getTitle()->setText('Ticket by Category');
        // Tambahkan seri data ke chart
        foreach ($resultdata_chart1 as $key => $value) {
            $series = new Series($value['problem'], ['Closed' =>  $value['count_closed'], 'Pending' => $value['count_pending']]);
            $chartType->addSeries($series);
        }

        // set Data Chart 2
        $data_chart2 = Data::whereBetween('created', [Carbon::parse($start_date)->subMonths(3), $end_date])->select(DB::raw('MONTH(created) as month'), DB::raw('count(*) as count'))
            ->groupBy(DB::raw('MONTH(created)'))
            ->get();
        $resultdata_chart2 = [];
        foreach ($data_chart2 as $key => $value) {
            $closed = Data::whereBetween('created', [Carbon::parse($start_date)->subMonths(3), $end_date])->where('status', '=', 'Closed')->where(DB::raw('MONTH(created)'), '=', $value->month)->get()->count();
            $pending = Data::whereBetween('created', [Carbon::parse($start_date)->subMonths(3), $end_date])->where('status', '=', 'Pending')->where(DB::raw('MONTH(created)'), '=', $value->month)->get()->count();
            $resultdata_chart2[] = [
                'month' => Carbon::create()->month($value->month)->format('F'),
                'count' => $value->count,
                'closed' => $closed,
                'pending' => $pending
            ];
        }

        // Chart 2
        $chartShape = $slide3->createChartShape();
        $chartShape->setHeight(400)
            ->setWidth(600)
            ->setOffsetX(650)
            ->setOffsetY(250);
        // Define tipe chart
        $chartType = new Bar();
        $chartShape->getPlotArea()->setType($chartType);

        // Set judul chart
        $chartShape->getTitle()->setText('Ticket by Last 3 Months');

        $dataclosed = [];
        foreach ($resultdata_chart2 as $key => $value) {
            $dataclosed[$value['month']] = $value['closed'];
        }
        $datapending = [];
        foreach ($resultdata_chart2 as $key => $value) {
            $datapending[$value['month']] = $value['pending'];
        }

        $series = new Series('Closed', $dataclosed);
        $series2 = new Series('Pending', $datapending);
        $chartType->addSeries($series);
        $chartType->addSeries($series2);


        //Slide 4
        $slide4 = $objPHPPresentation->createSlide();

        // Contoh data timeline
        $timelineData = [
            ['date' => '2023-01', 'event' => 'Event 1'],
            ['date' => '2023-02', 'event' => 'Event 2'],
            ['date' => '2023-03', 'event' => 'Event 3'],
            ['date' => '2023-04', 'event' => 'Event 4'],
        ];

        // Set posisi awal untuk timeline
        $x = 100;
        $y = 50;

        // Tambahkan garis horizontal sebagai garis waktu
        $timelineLine = $slide4->createLineShape($x, $y + 20, $x + 400, $y + 20);
        $timelineLine->getBorder()->setLineWidth(2)->setColor(new Color('FF000000'));

        // Tambahkan item timeline
        foreach ($timelineData as $data) {
            // Tambahkan lingkaran untuk titik waktu
            // $circle = $slide4->createShape(Circle::class)
            //     ->setHeight(20)
            //     ->setWidth(20)
            //     ->setOffsetX($x)
            //     ->setOffsetY($y + 10);
            // $circle->getBorder()->setLineWidth(2)->setColor(new Color('FF000000'));

            // Tambahkan shape kotak untuk tanggal
            $shape = $slide4->createRichTextShape()
                ->setHeight(50)
                ->setWidth(100)
                ->setOffsetX($x - 40)
                ->setOffsetY($y - 40);
            $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            $textRun = $shape->createTextRun($data['date']);
            $textRun->getFont()->setBold(true)
                ->setSize(12)
                ->setColor(new Color('FF000000'));

            // Tambahkan shape kotak untuk event
            $shape = $slide4->createRichTextShape()
                ->setHeight(50)
                ->setWidth(200)
                ->setOffsetX($x - 40)
                ->setOffsetY($y + 40);
            $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            $textRun = $shape->createTextRun($data['event']);
            $textRun->getFont()->setSize(12)
                ->setColor(new Color('FF000000'));

            // Update posisi X untuk item berikutnya
            $x += 100;
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
        $filename = 'Report IT Problem ' . date('d F Y') . '.pptx';
        $savePath = storage_path($filename);
        $writer = IOFactory::createWriter($objPHPPresentation, 'PowerPoint2007');
        $writer->save($savePath);
        // Return file sebagai response download
        return response()->download($savePath)->deleteFileAfterSend(true);
    }
}
