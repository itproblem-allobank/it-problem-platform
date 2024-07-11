<?php

namespace App\Http\Controllers;

use App\Models\Data;
use Illuminate\Support\Facades\DB;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\DocumentLayout;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar;
use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Shape\Drawing\File;

class PPTController extends Controller
{
    public function generateppt()
    {
        $objPHPPresentation = new PhpPresentation();
        $currentSlide = $objPHPPresentation->getActiveSlide();

        $backgroundImagePath = storage_path('image/background.png');
        $backgroundImage = new File();
        $backgroundImage->setPath($backgroundImagePath);
        $backgroundImage->setWidth(1280);
        $backgroundImage->setOffsetX(0);
        $backgroundImage->setOffsetY(0);
        $currentSlide->addShape($backgroundImage);


        $imagePath = storage_path('image/allobank.png');
        $pictureShape = new File();
        $pictureShape->setPath($imagePath);
        $pictureShape->setWidth(200);  // Ubah ukuran gambar sesuai kebutuhan
        $pictureShape->setOffsetX(1050); // Posisi horizontal gambar
        $pictureShape->setOffsetY(20); // Posisi vertikal gambar
        $currentSlide->addShape($pictureShape);

        $objPHPPresentation->getLayout()->setDocumentLayout(['cx' => 1280, 'cy' => 700], true)
            ->setCX(1280, DocumentLayout::UNIT_PIXEL)
            ->setCY(700, DocumentLayout::UNIT_PIXEL);

        // Tambahkan teks judul slide
        $shape = $currentSlide->createRichTextShape()
            ->setHeight(50)
            ->setWidth(400)
            ->setOffsetX(50)
            ->setOffsetY(25);
        $textRun = $shape->createTextRun('Report IT Problem');
        $textRun->getFont()->setBold(true)
            ->setSize(30);

        $shape = $currentSlide->createRichTextShape()
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
            $tableShape = $currentSlide->createTableShape(3);
            $tableShape->setHeight(100);
            $tableShape->setWidth(150);
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
            $offsetx = $offsetx + 200;
        }

        //set data chart 1
        $datachart = Data::select('problem', DB::raw('count(*) as count'))->groupBy('problem')->get();
        $resultdata = [];
        foreach ($datachart as $key => $value) {
            $status_closed = Data::where('problem', '=', $value->problem)->where('status', '=', 'Closed')->get()->count();
            $status_pending = Data::where('problem', '=', $value->problem)->where('status', '=', 'Pending')->get()->count();
            $resultdata[] =
                [
                    'problem' => $value->problem,
                    'total' => $value->count,
                    'count_closed' => $status_closed,
                    'count_pending' => $status_pending,
                ];
        }

        // Chart 1 
        $chartShape = $currentSlide->createChartShape();
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
        foreach ($resultdata as $key => $value) {
            $series = new Series($value['problem'], ['Closed' =>  $value['count_closed'], 'Pending' => $value['count_pending']]);
            $chartType->addSeries($series);
        }

        //set data chart 2
        $total2024 = Data::where('created', 'like', '%2024%')->get()->count();
        $closed2024 = Data::where('created', 'like', '%2024%')->where('status', '=', 'Closed')->get()->count();
        $pending2024 = Data::where('created', 'like', '%2024%')->where('status', '=', 'Pending')->get()->count();
        $wip2024 = Data::where('created', 'like', '%2024%')->where('status', '=', 'Work In Progress')->get()->count();
        $total2023 = Data::where('created', 'like', '%2023%')->get()->count();
        $closed2023 = Data::where('created', 'like', '%2023%')->where('status', '=', 'Closed')->get()->count();
        $pending2023 = Data::where('created', 'like', '%2023%')->where('status', '=', 'Pending')->get()->count();
        $wip2023 = Data::where('created', 'like', '%2023%')->where('status', '=', 'Work In Progress')->get()->count();

        // Chart 2
        $chartShape = $currentSlide->createChartShape();
        $chartShape->setHeight(400)
            ->setWidth(600)
            ->setOffsetX(650)
            ->setOffsetY(250);
        // Define tipe chart
        $chartType = new Bar();
        $chartShape->getPlotArea()->setType($chartType);

        // Set judul chart
        $chartShape->getTitle()->setText('Ticket by Yearly');
        // Tambahkan seri data ke chart  
        $total = new Series('Total', ['2024' =>  $total2024, '2023' => $total2023]);
        $closed = new Series('Closed', ['2024' =>  $closed2024, '2023' => $closed2023]);
        $pending = new Series('Pending', ['2024' =>  $pending2024, '2023' => $pending2023]);
        $wik = new Series('Work In Progress', [ '2024' =>  $wip2024, '2023' => $wip2023]);
        $chartType->addSeries($total);
        $chartType->addSeries($closed);
        $chartType->addSeries($pending);
        $chartType->addSeries($wik);

        // Simpan presentasi ke dalam file
        $filename = 'presentation_' . time() . '.pptx';
        $savePath = storage_path($filename);
        $writer = IOFactory::createWriter($objPHPPresentation, 'PowerPoint2007');
        $writer->save($savePath);
        // Return file sebagai response download
        return response()->download($savePath)->deleteFileAfterSend(true);
    }
}
