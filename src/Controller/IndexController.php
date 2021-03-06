<?php

namespace App\Controller;

use Symfony\Component\HttpFoundation\Request;
use Symfony\Component\HttpFoundation\Response;
use Symfony\Component\HttpFoundation\ResponseHeaderBag;
use Symfony\Component\Routing\Annotation\Route;
use Symfony\Component\Filesystem\Filesystem;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Symfony\Bundle\FrameworkBundle\Controller\Controller;
use PhpOffice\PhpSpreadsheet\Cell\DataType;

class IndexController extends Controller
{

    /**
     * return render template
     *
     * @Route("/" , name="index" )
     * @param Request $request
     * @return Response
     */
    public function RenderAction(Request $request)
    {
        if($request->isMethod('post')) {
//            $filename = 'f:\object.txt';

            $filesystem = new Filesystem();

            if ($filesystem->exists($_FILES['file_import']['tmp_name'])) {
                $contents = file($_FILES['file_import']['tmp_name']);

                $open_bracket = 0;
                $close_bracket = 0;
                $is_kontrahent = false;
                $ctr = 0;
                $i = 0;

                $kontrahents = [];

                foreach ($contents as $value) {
                    if ((strstr($value, "Kontrahent") != false) && ($is_kontrahent == false)) {
                        $is_kontrahent = true;
                        $open_bracket++;
                    }
                    else if ((strstr($value, "Kontrahent") == false) && ($is_kontrahent == true)) {
                        if (strpos($value, "{") != false) {
                            $open_bracket++;
                        }
                        else if (strstr($value, "}") != false) {
                            $close_bracket++;

                            if ($open_bracket == $close_bracket) {
                                $ctr++;
                                $is_kontrahent = false;
                                $open_bracket = 0;
                                $close_bracket = 0;
                            }
                        }
                        else
                            $kontrahents[$ctr][] = trim($value);

                    }
                    $i++;
                }
            }

//            echo '<pre>';
//            print_r(explode('=', $kontrahents[0][1])[1]);
//            echo '</pre>';
//            exit();


            $header = [];

            //get header array
            foreach ($kontrahents as $ahent)
                foreach ($ahent as $item)
                    $header[] = trim(explode('=', $item)[0]);

            $header = array_values(array_unique($header));

//            echo '<pre>';
//            print_r($header);
//            echo '</pre>';

            //export to excel
            $spreadsheet = new Spreadsheet();

            $letter = 1;

            //output header
            for ($i = 1; $i < count($header); $i++)
                if ($i <= 26)
                    $spreadsheet->getActiveSheet()->setCellValue(chr(64 + (int)$i) . '1', $header[$i-1]);
                else {
                    $spreadsheet->getActiveSheet()->setCellValue(chr(65) . chr(64 + (int)$letter) . '1', $header[$i-1]);
                    $letter++;
                }

            $letter = 1;
            //output contents
            for ($k = 0; $k < count($kontrahents); $k++) {
                $letter = 1;
                for ($j = 0; $j < count($kontrahents[$k]); $j++)
                    for ($i = 1; $i < count($header); $i++) {
                        if (trim(explode('=', trim($kontrahents[$k][$j]))[0]) == $header[$i-1])
                            if ($i <= 26) {
                                $spreadsheet->getActiveSheet()
                                    ->setCellValueExplicit(chr(64 + (int)$i) . (string)($k + 2),
                                        (string)mb_convert_encoding(trim(explode('=', $kontrahents[$k][$j])[1]), 'UTF-8'),
                                        DataType::TYPE_STRING);

                            } else {
                                $spreadsheet->getActiveSheet()
                                    ->setCellValueExplicit(chr(65) . chr(64 + (int)$letter) . (string)($k + 2),
                                        (string)mb_convert_encoding(trim(explode('=', $kontrahents[$k][$j])[1]), 'UTF-8'),
                                        DataType::TYPE_STRING);
                                $letter++;
                            }
                    }
            }

            $writer = new Xlsx($spreadsheet);

            $fileName = 'export.xlsx';
            $temp_file = tempnam(sys_get_temp_dir(), $fileName);

            $writer->save($temp_file);

            return $this->file($temp_file, $fileName, ResponseHeaderBag::DISPOSITION_INLINE);
        }
        else
            return $this->render('index.html.twig');
    }
    /**
     * @Route("submit" , name="submit")
     * @param Request $request
     * @return Response
     */
    public function SubmitAction(Request $request)
    {

    }

}