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
        setCellText($row, $cell, 'Oktober 2023', 15);

        $row = $tableShape->createRow();
        $cell = $row->nextCell();
        setCellText($row, $cell, 'Review date', 15);
        $cell = $row->nextCell();
        setCellText($row, $cell, '30 Oktober 2023', 15);

        //Text Shape 1
        $textShape1 = $slide2->createRichTextShape();
        $textShape1->setHeight(250);
        $textShape1->setWidth(300);
        $textShape1->setOffsetX(50);
        $textShape1->setOffsetY(420);
        $textShape1->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);

        // Create the text run for the left-aligned text
        $textRun2 = $textShape1->createTextRun("Jakarta, 11 Desember 2023\n\nDisetujui oleh,\n\n\n\n\n");
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
        $date = date('d F Y');
        $textRun = $shape->createTextRun('As of ' . $date);
        $textRun->getFont()->setSize(14);

        //data container category
        $problem = Data::select('problem', DB::raw('count(*) as count'))->groupBy('problem')->get();
        $total = [];
        foreach ($problem as $key => $value) {
            $highest = Data::where('problem', '=', $value->problem)->where('priority', '=', 'Highest')->get()->count();
            $high = Data::where('problem', '=', $value->problem)->where('priority', '=', 'High')->get()->count();
            $medium = Data::where('problem', '=', $value->problem)->where('priority', '=', 'Medium')->get()->count();
            $low = Data::where('problem', '=', $value->problem)->where('priority', '=', 'Low')->get()->count();
            $lowest = Data::where('problem', '=', $value->problem)->where('priority', '=', 'Lowest')->get()->count();
            $highestmonthly = Data::where('problem', '=', $value->problem)->where('priority', '=', 'Highest')->where('created', '>', now()->subDays(30))->get()->count();
            $highmonthly = Data::where('problem', '=', $value->problem)->where('priority', '=', 'High')->where('created', '>', now()->subDays(30))->get()->count();
            $mediummonthly = Data::where('problem', '=', $value->problem)->where('priority', '=', 'Medium')->where('created', '>', now()->subDays(30))->get()->count();
            $lowmonthly = Data::where('problem', '=', $value->problem)->where('priority', '=', 'Low')->where('created', '>', now()->subDays(30))->get()->count();
            $lowestmonthly = Data::where('problem', '=', $value->problem)->where('priority', '=', 'Lowest')->where('created', '>', now()->subDays(30))->get()->count();
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

            //row value //dibatalin karna munculin bulanan aja 

            // $rowShape = $tableShape->createRow();
            // $rowShape->setHeight(25);
            // $value = [$data['high'], $data['medium'], $data['low']];
            // foreach ($value as $key => $v) {
            //     $cell = $rowShape->nextCell();
            //     $cell->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            //     $cell->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
            //     $cell->createTextRun($v);
            // }

            //row value
            $rowShape = $tableShape->createRow();
            $rowShape->setHeight(25);
            $value = [$data['highmonthly'], $data['mediummonthly'], $data['lowmonthly']];
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
        $data_chart1 = Data::select('problem', DB::raw('count(*) as count'))->groupBy('problem')->get();
        $resultdata_chart1 = [];
        foreach ($data_chart1 as $key => $value) {
            $status_closed = Data::where('problem', '=', $value->problem)->where('status', '=', 'Closed')->get()->count();
            $status_pending = Data::where('problem', '=', $value->problem)->where('status', '=', 'Pending')->get()->count();
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
        $data_chart2 = Data::select(DB::raw('MONTH(created) as month'), DB::raw('count(*) as count'))
            ->groupBy(DB::raw('MONTH(created)'))
            ->get();
        $resultdata_chart2 = [];
        foreach ($data_chart2 as $key => $value) {
            $closed = Data::where('status', '=', 'Closed')->where(DB::raw('MONTH(created)'), '=', $value->month)->get()->count();
            $pending = Data::where('status', '=', 'Pending')->where(DB::raw('MONTH(created)'), '=', $value->month)->get()->count();
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
        $chartShape->getTitle()->setText('Ticket by 3 Months');

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

        // foreach ($resultdata_chart2 as $key => $value) {
        //     $series = new Series($value['month'], ['Closed' =>  $value['closed'], 'Pending' => $value['pending']]);
        //     $chartType->addSeries($series);
        // }


        // //set data chart 2
        // $total2024 = Data::where('created', 'like', '%2024%')->get()->count();
        // $closed2024 = Data::where('created', 'like', '%2024%')->where('status', '=', 'Closed')->get()->count();
        // $pending2024 = Data::where('created', 'like', '%2024%')->where('status', '=', 'Pending')->get()->count();
        // $wip2024 = Data::where('created', 'like', '%2024%')->where('status', '=', 'Work In Progress')->get()->count();
        // $total2023 = Data::where('created', 'like', '%2023%')->get()->count();
        // $closed2023 = Data::where('created', 'like', '%2023%')->where('status', '=', 'Closed')->get()->count();
        // $pending2023 = Data::where('created', 'like', '%2023%')->where('status', '=', 'Pending')->get()->count();
        // $wip2023 = Data::where('created', 'like', '%2023%')->where('status', '=', 'Work In Progress')->get()->count();

        // // Chart 2
        // $chartShape = $slide3->createChartShape();
        // $chartShape->setHeight(400)
        //     ->setWidth(600)
        //     ->setOffsetX(650)
        //     ->setOffsetY(250);
        // // Define tipe chart
        // $chartType = new Bar();
        // $chartShape->getPlotArea()->setType($chartType);

        // // Set judul chart
        // $chartShape->getTitle()->setText('Ticket by Yearly');
        // // Tambahkan seri data ke chart  
        // $total = new Series('Total', ['2024' =>  $total2024, '2023' => $total2023]);
        // $closed = new Series('Closed', ['2024' =>  $closed2024, '2023' => $closed2023]);
        // $pending = new Series('Pending', ['2024' =>  $pending2024, '2023' => $pending2023]);
        // $wik = new Series('Work In Progress', ['2024' =>  $wip2024, '2023' => $wip2023]);
        // $chartType->addSeries($total);
        // $chartType->addSeries($closed);
        // $chartType->addSeries($pending);
        // $chartType->addSeries($wik);


        //Slide 4
        $slide4 = $objPHPPresentation->createSlide();
        $backgroundImagePath = storage_path('image/background_end.png');
        $backgroundImage = new File();
        $backgroundImage->setPath($backgroundImagePath);
        $backgroundImage->setWidth(1280);
        $backgroundImage->setOffsetX(0);
        $backgroundImage->setOffsetY(0);
        $slide4->addShape($backgroundImage);

        $shape = $slide4->createRichTextShape()
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
