<?php

namespace App\Http\Controllers\Incident;

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
use App\Models\Incident;
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

        return view('incident/i-monthly');
    }

    public function download(Request $request)
    {
        $start_date = $request->start_date;
        $end_date = $request->end_date;
        // dd($request->end_date);
        $objPHPPresentation = new PhpPresentation();
        // Set layout custom
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
        $textRun = $shape->createTextRun('Monthly Report IT Incident');
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
            ->setSize(20)
            // ->setColor(new Color('FFFF0000'))
        ;

        //Text
        $shape = $slide1->createRichTextShape()
            ->setHeight(50)
            ->setWidth(280)
            ->setOffsetX(980)
            ->setOffsetY(640);
        $textRun = $shape->createTextRun('PT Allo Bank Indonesia');
        $textRun->getFont()->setSize(20);

        // ------------------------ SLIDE 2 ----------------------------------------------
        $slide2 = $objPHPPresentation->createSlide();
        $backgroundImagePath = storage_path('image/background.png');
        $backgroundImage = new File();
        $backgroundImage->setPath($backgroundImagePath);
        $backgroundImage->setWidth(1280);
        $backgroundImage->setOffsetX(0);
        $backgroundImage->setOffsetY(0);
        $slide2->addShape($backgroundImage);


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
        setCellText($row, $cell, 'Report Monthly IT Incident', 15);

        $row = $tableShape->createRow();
        $cell = $row->nextCell();
        setCellText($row, $cell, 'Version', 15);
        $cell = $row->nextCell();
        setCellText($row, $cell, Carbon::parse($end_date)->format('F Y'), 15);

        $row = $tableShape->createRow();
        $cell = $row->nextCell();
        setCellText($row, $cell, 'Review date', 15);
        $cell = $row->nextCell();
        // setCellText($row, $cell, Carbon::parse($end_date)->format('d F Y'), 15);
        setCellText($row, $cell, Carbon::now()->format('d F Y'), 15);

        //Text Shape 1
        $textShape1 = $slide2->createRichTextShape();
        $textShape1->setHeight(250);
        $textShape1->setWidth(300);
        $textShape1->setOffsetX(50);
        $textShape1->setOffsetY(420);
        $textShape1->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);

        // Create the text run for the left-aligned text
        // $date = Carbon::parse($end_date)->format('d F Y');
        $date = Carbon::now()->format('d F Y');
        $textRun2 = $textShape1->createTextRun("Jakarta, " . $date . "\n\nDisetujui oleh,\n\n\n\n\n");
        $textRun2->getFont()->setSize(15);
        $textRun2->getFont()->setColor(new Color('FF000000')); // Black color

        // Create the bold text run for "Tri Intan Siska P."
        $boldTextRun = $textShape1->createTextRun("Tri Intan Siska P.\n");
        $boldTextRun->getFont()->setSize(15);
        $boldTextRun->getFont()->setColor(new Color('FF000000')); // Black color
        $boldTextRun->getFont()->setBold(true); // Set the text to bold

        // Create the text run for "IT infra Operation"
        $textRun3 = $textShape1->createTextRun("Plt. IT Operations Dept. Head");
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

        // Create the text run for "IT Incident Lead"
        $textRun3 = $textShape2->createTextRun("IT Incident Section Head");
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

        // Create the text run for "IT Incident Lead"
        $textRun3 = $textShape2->createTextRun("IT Incident Engineer");
        $textRun3->getFont()->setSize(15);
        $textRun3->getFont()->setColor(new Color('FF000000')); // Black color

        // ----------------------- SLIDE 3 -------------------------------------------------
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
        $pictureShape->setWidth(200);
        $pictureShape->setOffsetX(1050);
        $pictureShape->setOffsetY(20);
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
        $textRun = $shape->createTextRun('Report IT Incident');
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

        // Line
        $imagePath = storage_path('image/Line.png');
        $pictureShape = new File();
        $pictureShape->setPath($imagePath);
        $pictureShape->setWidth(1200);
        $pictureShape->setOffsetX(20);
        $pictureShape->setOffsetY(100);
        $slide3->addShape($pictureShape);


        // Chart 1 - Incident by Category
        // Data
        $incidentByCategory = Incident::select('category', DB::raw('count(*) as total'))
            ->whereBetween('created_time', [$start_date, $end_date])
            ->groupBy('category')
            ->get();
        // dd(json_encode($incidentByCategory, JSON_PRETTY_PRINT));

        $jsonData = $incidentByCategory->toArray();
        // dd($jsonData);

        // Generate Chart
        $chartShape = $slide3->createChartShape();
        $chartShape->setHeight(250)
            ->setWidth(820)
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
        foreach ($jsonData as $key => $value) {
            $count = strval($value['total']);
            $series = new Series($value['category'], ['Total' => $count]);
            $chartType->addSeries($series);
        }




        // ----------------------- SLIDE CLOSING / SLIDE 5 ---------------------------------
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
        $filename = 'Report Monthly IT Incident' . ' - ' . Carbon::parse($start_date)->format('F Y') . '.pptx';
        $savePath = storage_path($filename);
        $writer = IOFactory::createWriter($objPHPPresentation, 'PowerPoint2007');
        $writer->save($savePath);

        // Simpan file Excel sementara
        $excelPath = 'exports/List Incident Monthly - ' .  Carbon::parse($start_date)->format('F Y') . '.xlsx';
        Excel::store(new DataExport($start_date, $end_date), $excelPath, 'local');

        // 3. Buat file ZIP yang berisi kedua file tersebut
        $zipFilename = 'Monthly Report - ' . Carbon::parse($start_date)->format('F Y') . '.zip';
        $zipFilePath = storage_path('app/exports/' . $zipFilename);
        $zip = new ZipArchive;
        if ($zip->open($zipFilePath, ZipArchive::CREATE) === TRUE) {
            $zip->addFile(storage_path('app/' . $excelPath), 'List Incident Monthly - ' .  Carbon::parse($start_date)->format('F Y') . '.xlsx');
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
