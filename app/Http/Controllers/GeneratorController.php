<?php

namespace App\Http\Controllers;

use PhpOffice\PhpSpreadsheet\IOFactory;

class GeneratorController extends Controller
{
    public function generate()
    {
        $reader_wilkerstat = IOFactory::createReaderForFile("assets/wonomerto taging st2023.xlsx");
        $reader_wilkerstat->setReadDataOnly(true);
        $wilkerstat = $reader_wilkerstat->load("assets/wonomerto taging st2023.xlsx");

        $reader_regsosek = IOFactory::createReaderForFile("assets/regsosek.xlsx");
        $reader_regsosek->setReadDataOnly(true);
        $regsosek = $reader_regsosek->load("assets/regsosek.xlsx");

        $reader_rekap_wilkerstat = IOFactory::createReaderForFile("assets/rekap wilkerstat.xlsx");
        $reader_rekap_wilkerstat->setReadDataOnly(true);
        $rekap_wilkerstat = $reader_rekap_wilkerstat->load("assets/rekap wilkerstat.xlsx");


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

        $matrix_result = [];

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
                    $matrix_result[$petugas][$date[0]][$kode_sls] = ['pemutakhiran' => 1, 'rutp' => 1, 'utp' => $wilkerstat->getActiveSheet()->getCell('P' . $i)->getValue()];
                } else {
                    $matrix_result[$petugas][$date[0]][$kode_sls]['rutp'] = $matrix_result[$petugas][$date[0]][$kode_sls]['rutp'] + 1;
                    $matrix_result[$petugas][$date[0]][$kode_sls]['utp'] = $matrix_result[$petugas][$date[0]][$kode_sls]['utp'] + $wilkerstat->getActiveSheet()->getCell('P' . $i)->getValue();
                    $matrix_result[$petugas][$date[0]][$kode_sls]['pemutakhiran'] = (int) floor($matrix_result[$petugas][$date[0]][$kode_sls]['rutp'] / $rekap_result[$kode_sls][0] * $regsosek_result[$kode_sls]);
                }
            }

            $i++;
        } while ($loop);

        dd($matrix_result);

        return "done";
    }
}
