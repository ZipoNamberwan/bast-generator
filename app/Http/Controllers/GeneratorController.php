<?php

namespace App\Http\Controllers;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpWord\TemplateProcessor;

class GeneratorController extends Controller
{
    public function generate()
    {
        $date_array = [
            '_1' => ['begin' => '2023-06-01', 'end' => '2023-06-07'],
            '_2' => ['begin' => '2023-06-08', 'end' => '2023-06-14'],
            '_3' => ['begin' => '2023-06-15', 'end' => '2023-06-21'],
            '_4' => ['begin' => '2023-06-22', 'end' => '2023-06-26'],
            '_5' => ['begin' => '2023-06-27', 'end' => '2023-06-30'],
        ];

        $reader_wilkerstat = IOFactory::createReaderForFile("assets/wilkerstat.xlsx");
        $reader_wilkerstat->setReadDataOnly(true);
        $wilkerstat = $reader_wilkerstat->load("assets/wilkerstat.xlsx");

        $reader_regsosek = IOFactory::createReaderForFile("assets/regsosek.xlsx");
        $reader_regsosek->setReadDataOnly(true);
        $regsosek = $reader_regsosek->load("assets/regsosek.xlsx");

        $reader_rekap_wilkerstat = IOFactory::createReaderForFile("assets/rekap wilkerstat.xlsx");
        $reader_rekap_wilkerstat->setReadDataOnly(true);
        $rekap_wilkerstat = $reader_rekap_wilkerstat->load("assets/rekap wilkerstat.xlsx");

        $repo = IOFactory::createReaderForFile("assets/repo.xlsx");
        $repo->setReadDataOnly(true);
        $repo = $repo->load("assets/repo.xlsx");

        $prelist = IOFactory::createReaderForFile("assets/kk tani prelist.xlsx");
        $prelist->setReadDataOnly(true);
        $prelist = $prelist->load("assets/kk tani prelist.xlsx");

        $master_sls = IOFactory::createReaderForFile("assets/master sls.xlsx");
        $master_sls->setReadDataOnly(true);
        $master_sls = $master_sls->load("assets/master sls.xlsx");

        $struktur = IOFactory::createReaderForFile("assets/struktur.xlsx");
        $struktur->setReadDataOnly(true);
        $struktur = $struktur->load("assets/struktur.xlsx");

        $loop = true;
        $i = 2;
        $rekap_result = [];
        do {
            if ($rekap_wilkerstat->getActiveSheet()->getCell('A' . $i)->getValue() == null) {
                $loop = false;
            }

            $idsls = $rekap_wilkerstat->getActiveSheet()->getCell('A' . $i)->getValue();
            if (!key_exists($idsls, $rekap_result)) {
                $rekap_result[$idsls] = [$rekap_wilkerstat->getActiveSheet()->getCell('B' . $i)->getValue(), $rekap_wilkerstat->getActiveSheet()->getCell('C' . $i)->getValue()];
            }

            $i++;
        } while ($loop);

        $loop = true;
        $i = 2;
        $regsosek_result = [];
        do {
            if ($regsosek->getActiveSheet()->getCell('A' . $i)->getValue() == null) {
                $loop = false;
            }

            $idsls = $regsosek->getActiveSheet()->getCell('A' . $i)->getValue();
            if (!key_exists($idsls, $regsosek_result)) {
                $regsosek_result[$idsls] = $regsosek->getActiveSheet()->getCell('R' . $i)->getValue();
            }
            $i++;
        } while ($loop);

        $loop = true;
        $i = 2;
        $repo_result = [];
        do {
            if ($repo->getActiveSheet()->getCell('A' . $i)->getValue() == null) {
                $loop = false;
            }

            $idsls = $repo->getActiveSheet()->getCell('A' . $i)->getValue();
            if (!key_exists($idsls, $repo_result)) {
                $repo_result[$idsls] = $repo->getActiveSheet()->getCell('M' . $i)->getValue() != null;
            }
            $i++;
        } while ($loop);

        $loop = true;
        $i = 2;
        $prelist_result = [];
        do {
            if ($prelist->getActiveSheet()->getCell('A' . $i)->getValue() == null) {
                $loop = false;
            }

            $idsls = $prelist->getActiveSheet()->getCell('A' . $i)->getValue();
            if (!key_exists($idsls, $prelist_result)) {
                $prelist_result[$idsls] = $prelist->getActiveSheet()->getCell('B' . $i)->getValue();
            }
            $i++;
        } while ($loop);

        $loop = true;
        $i = 2;
        $struktur_result = [];
        do {
            if ($struktur->getActiveSheet()->getCell('A' . $i)->getValue() == null) {
                $loop = false;
            }

            $ppl = $struktur->getActiveSheet()->getCell('A' . $i)->getValue();
            if (!key_exists($ppl, $struktur_result)) {
                $struktur_result[$ppl] = $struktur->getActiveSheet()->getCell('B' . $i)->getValue();
            }
            $i++;
        } while ($loop);

        $loop = true;
        $i = 2;
        $master_sls_result = [];
        do {
            if ($master_sls->getActiveSheet()->getCell('A' . $i)->getValue() == null) {
                $loop = false;
            }

            $idsls = $master_sls->getActiveSheet()->getCell('L' . $i)->getCalculatedValue();
            if (!key_exists($idsls, $master_sls_result)) {
                $master_sls_result[$idsls] = [
                    $master_sls->getActiveSheet()->getCell('G' . $i)->getValue(),
                    $master_sls->getActiveSheet()->getCell('H' . $i)->getValue(),
                    $master_sls->getActiveSheet()->getCell('J' . $i)->getValue(),
                    $master_sls->getActiveSheet()->getCell('K' . $i)->getValue()
                ];
            }
            $i++;
        } while ($loop);

        $loop = true;
        $i = 2;

        $matrix_result = [];
        $matrix_temporary = [];

        do {
            if ($wilkerstat->getActiveSheet()->getCell('A' . $i)->getValue() == null) {
                $loop = false;
            }

            if ($loop) {
                $petugas = $wilkerstat->getActiveSheet()->getCell('M' . $i)->getValue();

                if (!key_exists($petugas, $matrix_result))
                    $matrix_result[$petugas] = [];

                $date = explode(" ", $wilkerstat->getActiveSheet()->getCell('I' . $i)->getValue());
                if (!key_exists($date[0], $matrix_result[$petugas])) {
                    $matrix_result[$petugas][$date[0]] = [];
                }

                $kode_sls = $wilkerstat->getActiveSheet()->getCell('L' . $i)->getValue() . $wilkerstat->getActiveSheet()->getCell('D' . $i)->getValue();
                if (!key_exists($kode_sls, $matrix_result[$petugas][$date[0]])) {
                    if (!key_exists($kode_sls, $matrix_temporary)) {
                        $matrix_temporary[$kode_sls] = ['rutp' => 1, 'utp' => $wilkerstat->getActiveSheet()->getCell('P' . $i)->getValue()];
                    } else {
                        $matrix_temporary[$kode_sls]['rutp'] = $matrix_temporary[$kode_sls]['rutp'] + 1;
                        $matrix_temporary[$kode_sls]['utp'] = $matrix_temporary[$kode_sls]['utp'] + $wilkerstat->getActiveSheet()->getCell('P' . $i)->getValue();
                    }
                    $matrix_result[$petugas][$date[0]][$kode_sls] = ['pemutakhiran' => (int) floor($matrix_temporary[$kode_sls]['rutp'] / $rekap_result[$kode_sls][0] * $regsosek_result[$kode_sls]), 'rutp' => $matrix_temporary[$kode_sls]['rutp'], 'utp' => $matrix_temporary[$kode_sls]['utp']];
                } else {
                    $matrix_temporary[$kode_sls]['rutp'] = $matrix_temporary[$kode_sls]['rutp'] + 1;
                    $matrix_temporary[$kode_sls]['utp'] = $matrix_temporary[$kode_sls]['utp'] + $wilkerstat->getActiveSheet()->getCell('P' . $i)->getValue();
                    $matrix_result[$petugas][$date[0]][$kode_sls]['rutp'] = $matrix_temporary[$kode_sls]['rutp'];
                    $matrix_result[$petugas][$date[0]][$kode_sls]['utp'] = $matrix_temporary[$kode_sls]['utp'];
                    $matrix_result[$petugas][$date[0]][$kode_sls]['pemutakhiran'] = (int) floor($matrix_temporary[$kode_sls]['rutp'] / $rekap_result[$kode_sls][0] * $regsosek_result[$kode_sls]);
                }

                $matrix_result[$petugas][$date[0]][$kode_sls]['status'] = $repo_result[$kode_sls] ? ($matrix_result[$petugas][$date[0]][$kode_sls]['rutp'] >= $rekap_result[$kode_sls][0] ? 1 : 2) : 2;
                $matrix_result[$petugas][$date[0]][$kode_sls]['prelist'] = (int) floor($matrix_temporary[$kode_sls]['rutp'] / $rekap_result[$kode_sls][0] * $prelist_result[$kode_sls]);
                $matrix_result[$petugas][$date[0]][$kode_sls]['desa'] = "[" . $master_sls_result[$kode_sls][0] . "] " . $master_sls_result[$kode_sls][1];
                $matrix_result[$petugas][$date[0]][$kode_sls]['sls'] = "[" . $master_sls_result[$kode_sls][2] . "] " . $master_sls_result[$kode_sls][3];
            }

            $i++;
        } while ($loop);

        foreach ($date_array as $period => $date) {
            foreach ($matrix_result as $petugas => $matrix) {
                $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('assets/template.xlsx');
                $sheet = $spreadsheet->getActiveSheet();
                $sheet->setCellValue('J31', strtoupper($petugas));
                $sheet->setCellValue('D25', "Diterima tanggal: " . date('d', strtotime($date['end'])) . ' Juni 2023');
                $sheet->setCellValue('D31', strtoupper($struktur_result[$petugas]));
                $row = 13;

                foreach ($matrix as $d => $sls) {
                    $datecacah = date('Y-m-d', strtotime($d));
                    $begin = date('Y-m-d', strtotime($date['begin']));
                    $end = date('Y-m-d', strtotime($date['end']));

                    if (($datecacah > $begin) && ($datecacah <= $end)) {
                        foreach ($sls as $kode_sls => $value) {
                            $sheet->setCellValue('B' . $row, ($row - 12));
                            $sheet->setCellValue('C' . $row, $value['desa']);
                            $sheet->setCellValue('E' . $row, $value['sls']);
                            $sheet->setCellValue('F' . $row, 'SLS');
                            $sheet->setCellValue('G' . $row, $d);
                            $sheet->setCellValue('H' . $row, $value['prelist']);
                            $sheet->setCellValue('I' . $row, $value['pemutakhiran']);
                            $sheet->setCellValue('J' . $row, $value['utp']);
                            $sheet->setCellValue('K' . $row, $value['rutp']);
                            $sheet->setCellValue('L' . $row, $value['status']);

                            $row++;
                        }
                    }
                }

                $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
                $writer->save("assets/result/" . $petugas . $period . ".xlsx");
            }
        }

        dd($matrix_result);

        return "done";
    }
}
