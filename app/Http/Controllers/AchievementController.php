<?php

namespace App\Http\Controllers;

use App\User;
use App\Master;
use App\Achievement;
use Carbon\Carbon;
use Illuminate\Http\Request;

class AchievementController extends Controller
{
    /**
     * 利用者の実績ページ
     * @param Request $request
     * @return void
     * 実績データと当月の日数、曜日、使用する引数を取得して登録ページへ移動
     */
    public function selection(Request $request)
    {
        //利用者の情報を取得
        $user = new User();
        $user = $user->getUser($request);

        //月初を取得
        $bmonth = new Carbon($request->month);

        $days = new Achievement();
        $days = $days->M_Days($request);

        //今月から過去1年間の月を取得
        $months = new Achievement();
        $months = $months->Months($request);

        //利用者の当日の実績データを取得
        $one_record = new Achievement();
        $one_record = $one_record->One_Record($request);

        //利用者の月の実績データの取得
        $records = new Achievement();
        $records = $records->Month_Records($request);

        $data = [
            'user' => $user,
            'bmonth' => $bmonth,
            'days' => $days,
            'months' => $months,
            'one_record' => $one_record,
            'records' => $records,
        ];
        return view('achievement.record', $data);
    }

    /**
     * 実績テーブルに新規レコードを作成
     * @param Request $request
     * @return void
     * 当日の開始時刻と登録日を作成
     */
    public function new_record(Request $request)
    {
        //利用者の今日の実績データを検索
        $one_record = new Achievement();
        $one_record = $one_record->One_Record($request);

        //データが存在した場合戻る
        if ($one_record != null) {
            return back();
        } else {
            //存在しなかった場合レコードを新規作成して戻る
            $new_record = new Achievement();
            $new_record = $new_record->New_Record($request);
            return back();
        }
    }

    /**
     * 当日の終了時刻を取得してレコードに追加
     * @param Request $request
     * @return void
     */
    public function end_time(Request $request)
    {

        //利用者の今日の実績データを検索
        $one_record = new Achievement();
        $one_record = $one_record->One_Record($request);

        //当日のデータのend_timeの値がnullでなかった場合は戻る
        if ($one_record->end_time != null) {
            return back();
        } else {
            //nullだった場合レコードに終了時刻を登録して戻る
            $end_time = new Achievement();
            $end_time = $end_time->End_Time($request);
            return back();
        }
    }

    /**
     * 食事提供加算の更新処理
     * @param Request $request
     * @return void
     */
    public function food_up(Request $request)
    {
        //利用者の今日の実績データを検索
        $one_record = new Achievement();
        $one_record = $one_record->One_Record($request);

        //データが存在した場合食事提供加算のレコードを更新して戻る
        if ($one_record != null) {
            $food = new Achievement();
            $food = $food->Food_Up($request);
            return back();
        } else {
            //存在しなかった場合戻る
            return back();
        }
    }

    /**
     * 施設外支援の更新処理
     * @param Request $request
     * @return void
     */
    public function outside_up(Request $request)
    {
        //利用者の今日の実績データを検索
        $one_record = new Achievement();
        $one_record = $one_record->One_Record($request);

        //データが存在した場合食事提供加算のレコードを更新して戻る
        if ($one_record != null) {
            $outside = new Achievement();
            $outside = $outside->Outside_Up($request);
            return back();
        } else {
            //存在しなかった場合戻る
            return back();
        }
    }

    /**
     * 医療連携体制加算の更新処理
     * @param Request $request
     * @return void
     */
    public function medical_up(Request $request)
    {
        //利用者の今日の実績データを検索
        $one_record = new Achievement();
        $one_record = $one_record->One_Record($request);

        //データが存在した場合食事提供加算のレコードを更新して戻る
        if ($one_record != null) {
            $medical = new Achievement();
            $medical = $medical->Medical_up($request);
            return back();
        } else {
            //存在しなかった場合戻る
            return back();
        }
    }

    /**
     * 備考の更新処理
     * @param Request $request
     * @return void
     */
    public function note_up(Request $request)
    {
        //利用者の今日の実績データを検索
        $one_record = new Achievement();
        $one_record = $one_record->One_Record($request);

        //データが存在した場合食事提供加算のレコードを更新して戻る
        if ($one_record != null) {
            $note = new Achievement();
            $note = $note->Note_Up($request);
            return back();
        } else {
            //存在しなかった場合戻る
            return back();
        }
    }

    /**
     * 利用者実績の作成ページへ移動
     * @param Request $request
     * @return void
     */
    public function add_achievement(Request $request)
    {
        //利用者情報を取得
        $user = new User();
        $user = $user->getUser($request);

        return view('master.add_achievement', $user);
    }

    /**
     * 利用者実績の作成を実行
     * @param Request $request
     * @return void
     */
    public function create_achievement(Request $request)
    {
        //利用者の当日レコードが存在しないかをチェック
        $one_record = Achievement::where('user_id', $request->user_id)
            ->wheredate('insert_date', $request->insert_date)
            ->first();
        //当日のレコードが存在していた場合セッションにメッセージを保存して戻る
        if ($one_record != null) {
            return back()
                ->withInput()
                ->with('error', '既に実績データが登録されています。');
        } else {
            //バリデーションを実行
            $this->validate($request, Achievement::$rules, Achievement::$messages);

            //Achievementモデルのオブジェクト作成
            $achievement = new Achievement();

            //formの内容を全て取得
            $form = $request->all();
            //内容を更新して保存
            $achievement->fill($form)->save();

            $month = new Carbon($request->insert_date);
            $month = $month->firstOfMonth();
            $request->merge(['insert_date' => $month]);

            //利用者の当月の実績データを取得
            $data = new Master();
            $data = $data->DRecords($request);

            //実績閲覧ページにリダイレクト
            return view('master.master_index', $data);
        }
    }

    /**
     * 実績編集ページ
     * @pahram Request $request
     * @return void
     * idからレコードを抽出して編集ページに渡す
     */
    public function edit_achievement(Request $request)
    {
        //利用者情報の取得
        $user = new User();
        $user = $user->getUser($request);

        //実績データがの取得
        $achievement = Achievement::where('achievements.id', $request->id)
            ->first();

        $data = [
            'user' => $user,
            'achievement' => $achievement,
        ];
        return view('master.edit_achievement', $data);
    }

    /**
     * 実績編集を実行
     * @pahram Request $request
     * @return void
     * idからレコードを抽出して各項目を編集
     */
    public function update_achievement(Request $request)
    {
        //Achievementモデルのオブジェクト作成
        $achievement = Achievement::where('id', $request->id)
            ->first();
        //formの内容を全て取得
        $form = $request->all();
        //内容を更新して保存
        $achievement->fill($form)->save();

        $month = new Carbon($request->insert_date);
        $month = $month->firstOfMonth();
        $request->merge(['insert_date' => $month]);

        //利用者の当月の実績データを取得
        $data = new Master();
        $data = $data->DRecords($request);

        //実績閲覧ページにリダイレクト
        return view('master.master_index', $data);
    }

    /**
     * 実績削除を実行
     * @pahram Request $request
     * @return void
     * ダイアログ
     */
    public function delete_achievement(Request $request)
    {
        $achievement = Achievement::where('id', $request->id)
            ->first();
        if (is_null($achievement)) {
            //実績データがない場合
            return redirect('/master');
        } else {
            //存在していた場合削除処理を実行
            $achievement->delete();
            //実績閲覧ページにリダイレクト
            return back();
        }
    }
}
