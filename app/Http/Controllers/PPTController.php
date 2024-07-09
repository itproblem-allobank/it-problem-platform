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
use PhpOffice\PhpPresentation\Shape\Chart\Type\Line;

class PPTController extends Controller
{
    public function generateppt()
    {
        $objPHPPresentation = new PhpPresentation();
        $currentSlide = $objPHPPresentation->getActiveSlide();
        $objPHPPresentation->getLayout()->setDocumentLayout(['cx' => 1280, 'cy' => 700], true)
            ->setCX(1280, DocumentLayout::UNIT_PIXEL)
            ->setCY(700, DocumentLayout::UNIT_PIXEL);

        // Tambahkan teks judul slide
        $shape = $currentSlide->createRichTextShape()
            ->setHeight(50)
            ->setWidth(600)
            ->setOffsetX(170)
            ->setOffsetY(50);
        $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $textRun = $shape->createTextRun('Report IT Problem');
        $textRun->getFont()->setBold(true)
            ->setSize(24)
            ->setColor(new Color('FFE06B20'));

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
            ->setOffsetY(180);
        // Define tipe chart
        $chartType = new Bar();
        $chartShape->getPlotArea()->setType($chartType);
        // Set judul chart
        $chartShape->getTitle()->setText('Ticket by Category');
        // Tambahkan seri data ke chart
        foreach ($resultdata as $key => $value) {
            $series = new Series($value['problem'], [$value['count_closed'], $value['count_pending']]);
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
            ->setOffsetY(180);
        // Define tipe chart
        $chartType = new Bar();
        $chartShape->getPlotArea()->setType($chartType);
        // Set judul chart
        $chartShape->getTitle()->setText('Ticket by Yearly');
        // Tambahkan seri data ke chart  
        $total = new Series('Total', [$total2024, $total2023]);
        $closed = new Series('Closed', [$closed2024, $closed2023]);
        $pending = new Series('Pending', [$pending2024, $pending2023]);
        $wik = new Series('Work In Progress', [$wip2024, $wip2023]);
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
