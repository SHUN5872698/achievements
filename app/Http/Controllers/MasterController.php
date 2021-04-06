<?php

namespace App\Http\Controllers;

use App\User;
use App\Achievement;
use App\Master;
use Carbon\Carbon;
use Illuminate\Support\Facades\DB;
use App\Exports\Export;
use App\Http\Requests\AchievementFormRequest;
use Illuminate\Http\Request;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\File;

class MasterController extends Controller
{
    /**
     * 実績閲覧ページ
     *
     * @param Request $request
     * @return void
     * 本校と2校別で利用者情報を取得
     */
    public function master_index(Request $request)
    {
        //本校の利用者情報の取得
        $school_1 = new User();
        $school_1 = $school_1->School_1();

        //2校の利用者情報の取得
        $school_2 = new User();
        $school_2 = $school_2->School_2();

        $bmonth = Carbon::now()->firstOfMonth();

        $data = [
            'school_1' => $school_1,
            'school_2' => $school_2,
            'bmonth' => $bmonth,
        ];
        return view('master.master_index', $data);
    }

    /**
     * 実績閲覧ページ
     *
     * @param Request $request
     * @return void
     * 利用者情報一覧と選択された利用者の実績データを取得
     */
    public function check_records(Request $request)
    {
        $data = new Master();
        $data = $data->Records($request);

        return view('master.master_index', $data);
    }

    /**
     * 個別で実績データをダウンロード
     *
     * @param Request $request
     * @return void
     *
     */
    public function one_data(Request $request)
    {
        //月初を取得
        $date = new Carbon($request->month);

        //月初を取得
        $bmonth = new Carbon($request->month);

        //利用者の情報を取得
        $user = new User();
        $user = $user->getUser($request);

        //execelの日付と曜日、サービス提供の状況欄に記入する配列を作成
        $days = new Master();
        $days = $days->Excel_Days($request);

        //利用者の一ヶ月分の実績データを取得
        $records = new Master();
        $records = $records->Month_Records($request);

        //テンプレートシートを選択
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('./excel/sample.xlsx');
        //シートを情報を取得
        $sheet = $spreadsheet->getActiveSheet();

        //シートに日付と曜日、在籍校名をセット
        $sheet->setCellValue('A1', $bmonth->isoformat('Y'));
        $sheet->setCellValue('A2', $bmonth->isoformat('M'));
        $sheet->setCellValue('A4', $user->first_name . $user->last_name);
        if ($user->school_id == 1) {
            $sheet->setCellValue('j4', "未来のかたち 本町本校");
        } else if ($user->school_id == 2) {
            $sheet->setCellValue('j4', "未来のかたち 本町第２校");
        }
        $sheet->fromArray($days, null, 'A9');

        //利用者の実績データをセットする
        $offset = 9;
        foreach ($records as $i => $record) {
            $rowNum = $i + $offset;
            if (
                array_key_exists('insert_date', $record) and
                array_key_exists('start_time', $record) and
                array_key_exists('end_time', $record) and
                array_key_exists('food', $record) and
                array_key_exists('outside_support', $record) and
                array_key_exists('medical_support', $record) and
                array_key_exists('note', $record)
            ) {
                $sheet->setCellValueByColumnAndRow(3, $rowNum, "");
                $sheet->setCellValueByColumnAndRow(4, $rowNum, $record['start_time']);
                $sheet->setCellValueByColumnAndRow(5, $rowNum, $record['end_time']);
                $sheet->setCellValueByColumnAndRow(7, $rowNum, $record['food']);
                $sheet->setCellValueByColumnAndRow(8, $rowNum, $record['outside_support']);
                $sheet->setCellValueByColumnAndRow(9, $rowNum, $record['medical_support']);
                $sheet->setCellValueByColumnAndRow(10, $rowNum, $record['note']);
            }
        }

        //Excelデータを新規作成
        $writer = new Xlsx($spreadsheet);
        //Excelの保存先のディレクトリ名とファイル名の変数を作成
        $excel_name = './excel/school_' . $user->school_id . '/' . $date->isoFormat('Y年M月分') . $user->first_name . $user->last_name . '.xlsx';
        //Excelファイルに名前をつけて保存
        $writer->save($excel_name);
        //Excelファイルのダウンロード
        return response()->download($excel_name);
    }

    public function dl_school1(Request $request)
    {

        //本校の利用者情報の取得
        $user = new User();
        $user = $user->ExcleSchool_1();

        //当月から一年間分の月初を取得
        $months = new Master();
        $months = $months->Exele_Months();

        $data = [
            'user' => $user,
            'months' => $months,
        ];

        return view('master.dl_excel', $data);
    }

    public function dl_school2(Request $request)
    {

        //２校の利用者情報の取得
        $user = new User();
        $user = $user->ExcleSchool_2();

        //当月から一年間分の月初を取得
        $months = new Master();
        $months = $months->Exele_Months();

        $data = [
            'user' => $user,
            'months' => $months,
        ];
        return view('master.dl_excel', $data);
    }

    /**
     * 在籍校別で利用者情報をExcelで一括出力させる
     *
     * @param Request $request
     * @return void
     */
    public function bulk_creation(Request $request)
    {
        //在籍校別の利用者一覧を取得
        $school_id = $request->school_id;
        $users = User::where('school_id', $school_id)->get();

        //月初を取得
        $bmonth = new Carbon($request->month);

        //月の日数が何日かを取得
        $dmonth = new Achievement();
        $dmonth = $dmonth->D_Month($request);

        //一ヶ月分の日数を取得
        $mdays = new Achievement();
        $mdays = $mdays->M_Days($request);

        //execelのファイル名の変数を作成

        //execelの日付と曜日、サービス提供の状況欄に記入する配列を作成
        $days = new Master();
        $days = $days->Excel_Days($request);

        //ユーザー毎の一ヶ月間の実績データを繰り返し連想配列に登録してExcelファイルに出力する
        foreach ($users as $user) {
            //実績を格納する配列を初期化
            $records = null;
            //配列を日数分作成
            for ($n = 0; $n < $dmonth; $n++) {
                $records[][$n] = null;
            }
            //利用者の一ヶ月間の実績データを取得してExcel出力するために配列に変換
            $achievements = Achievement::where('user_id', $user->id)
                ->whereYear('insert_date', $bmonth)
                ->whereMonth('insert_date', $bmonth)
                ->orderBy('insert_date', 'asc')
                ->select(
                    'insert_date',
                    'start_time',
                    'end_time',
                    'visit_support',
                    'food',
                    'outside_support',
                    'medical_support',
                    'note',
                    'stamp',
                )
                ->get()
                ->toarray();

            //実績登録日と一ヶ月の日数を比較して一致した場合$recordsの配列にレコードを上書きする
            foreach ($achievements as $achievement) {
                //開始時間と終了時間を変換
                $achievement['start_time'] = substr($achievement['start_time'], 0, 5);
                $achievement['end_time'] = substr($achievement['end_time'], 0, 5);
                for ($n = 0; $n < $dmonth; $n++) {
                    if ($mdays[$n] == $achievement['insert_date']) {
                        $records[$n] = $achievement;
                    }
                }
            }

            //テンプレートシートを選択
            $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('./excel/sample.xlsx');
            //シートを情報を取得
            $sheet = $spreadsheet->getActiveSheet();

            //シートに日付と曜日、在籍校名をセット
            $sheet->setCellValue('A1', $bmonth->isoformat('Y'));
            $sheet->setCellValue('A2', $bmonth->isoformat('M'));
            $sheet->setCellValue('A4', $user->first_name . $user->last_name);
            if ($user->school_id == 1) {
                $sheet->setCellValue('j4', "未来のかたち 本町本校");
            } else if ($user->school_id == 2) {
                $sheet->setCellValue('j4', "未来のかたち 本町第２校");
            }
            $sheet->fromArray($days, null, 'A9');

            //利用者の実績データをセットする
            $offset = 9;
            foreach ($records as $i => $record) {
                $rowNum = $i + $offset;
                if (
                    array_key_exists('insert_date', $record) and
                    array_key_exists('start_time', $record) and
                    array_key_exists('end_time', $record) and
                    array_key_exists('food', $record) and
                    array_key_exists('outside_support', $record) and
                    array_key_exists('medical_support', $record) and
                    array_key_exists('note', $record)
                ) {
                    $sheet->setCellValueByColumnAndRow(3, $rowNum, "");
                    $sheet->setCellValueByColumnAndRow(4, $rowNum, $record['start_time']);
                    $sheet->setCellValueByColumnAndRow(5, $rowNum, $record['end_time']);
                    $sheet->setCellValueByColumnAndRow(7, $rowNum, $record['food']);
                    $sheet->setCellValueByColumnAndRow(8, $rowNum, $record['outside_support']);
                    $sheet->setCellValueByColumnAndRow(9, $rowNum, $record['medical_support']);
                    $sheet->setCellValueByColumnAndRow(10, $rowNum, $record['note']);
                }
            }
            //Excelデータを新規作成
            $writer = new Xlsx($spreadsheet);

            //Excelの保存先のディレクトリ名とファイル名の変数を作成
            $excel_name = './excel/school_' . $user->school_id . '/' . $bmonth->isoFormat('Y年M月分') . $user->first_name . $user->last_name . '.xlsx';
            //Excelファイルに名前をつけて作成
            $writer->save($excel_name);
            //aaa
        }
        //Excelファイルのダウンロード
        return response()->download($excel_name);
    }
}
