<?php

namespace App;

use Illuminate\Http\Request;
use Carbon\Carbon;
use Illuminate\Database\Eloquent\Model;

class Achievement extends Model
{
    /** ガードするフィールド */
    protected $guarded = array('id');

    /** 変更を許可するフィールド */
    protected $fillable = array('user_id', 'insert_date', 'start_time', 'end_time', 'food', 'outside_support', 'medical_support', 'note');

    /** date型にキャストするフィールド */
    protected $dates = array('insert_date');

    //実績登録、編集用のバリデーション
    public static $rules = [
        'start_time' => ['required', 'date_format:H:i'],
        'end_time' => ['required', 'date_format:H:i', 'after:start_time'],
    ];

    /** バリデーションのエラーメッセージ */
    public static $messages = [
        'start_time.required' => '必須項目です',
        'start_time.date_format' => 'H:i形式で入力して下さい。',

        'end_time.required' => '必須項目です',
        'end_time.date_format' => 'H:i形式で入力して下さい。',
        'end_time.after' => '開始時間より後の時間を入力して下さい。',
    ];

    /** usersテーブルとリレーション処理 */
    public function user()
    {
        return $this->belongsTo('App\User', 'id');
    }

    /**
     * 月の日数が何日かを取得
     *
     * @param Request $request
     * @return void
     */
    public function D_Month(Request $request)
    {
        $dmonth = new Carbon($request->month);
        $dmonth = $dmonth->daysInMonth;
        return $dmonth;
    }

    /**
     * 一ヶ月の日数の取得処理
     *
     * @param Request $request
     * @return void
     */
    /**
     * 一ヶ月分の日数を取得
     */
    public function M_Days(Request $request)
    {
        //月初を取得
        $bmonth = new Carbon($request->month);

        //月の日数が何日かを取得
        $dmonth = new Master();
        $dmonth = $dmonth->D_Month($request);

        //日付を連想配列に格納
        for ($i = 0; $i <= $dmonth; $i++) {
            $days[$i] = $bmonth->copy()->addDay($i);
        }
        return ($days);
    }

    /**
     * 今月から過去1年間の月を取得
     *
     * @return void
     */
    public function Months(Request $request)
    {
        //今月の月初を取得
        $bmonth = Carbon::now()->firstOfMonth();

        //requestで送信された日付をカーボンにフォーマット
        $submonth = new Carbon($request->month);

        //配列に格納する
        $month = array($submonth);

        //一年分の月初を配列に格納
        for ($i = 0; $i < 12; $i++) {
            $months[$i] = $bmonth->copy()->subMonth($i);
        }

        //配列を結合する
        $months = array_merge($month, $months);

        //重複する配列を削除
        $months = array_unique($months);

        return $months;
    }

    /**
     * 利用者の当日の実績データを取得
     * @param Request $request
     * @return void
     */
    public function One_Record(Request $request)
    {
        $now = Carbon::now()->isoformat("Y-M-D");

        $one_record = Achievement::where('user_id', $request->user_id)
            ->wheredate('insert_date', $now)
            ->first();
        return $one_record;
    }

    /**
     * 利用者の当月の実績データを取得
     *
     * @param Request $request
     * @return void
     * リクエストで送られてきた値から何月かを取得して一ヶ月分のレコードを取得
     * 実績データの登録日と一ヶ月の日数を比較して一致した日数にレコードを代入
     */
    public function Month_Records(Request $request)
    {
        //月初を取得
        $bmonth = new Carbon($request->month);

        //月の日数を作成
        $dmonth = new Achievement();
        $dmonth = $dmonth->D_Month($request);

        //一ヶ月の日数の配列を取得
        $records = new Achievement();
        $records = $records->M_Days($request);

        //利用者の一ヶ月間のレコードを取得
        $achievements = Achievement::where('achievements.user_id', $request->user_id)
            ->whereYear('insert_date', $bmonth)
            ->whereMonth('insert_date', $bmonth)
            ->orderBy('insert_date', 'asc')
            ->get();

        //実績データの登録日と一ヶ月の日数を比較して一致した日数にレコードを代入
        foreach ($achievements as $achievement) {
            for ($n = 0; $n < $dmonth; $n++) {
                if ($records[$n] == $achievement->insert_date) {
                    $records[$n] = $achievement;
                }
            }
        }
        return $records;
    }

    /**
     * 当日の実績レコードの作成
     *
     * @param Request $request
     * @return void
     * 現在時刻を15分切り上げてフォーマット
     * 9時30分以前ならダミーの時刻を登録
     */
    public function New_Record(Request $request)
    {
        //今日の登録日を取得してフォーマット
        $insert_date = Carbon::now()->format("Y-m-d");

        //現在時刻を取得して15分切り上げる
        $start_time = Carbon::now();
        $start_time->addMinutes(15 - $start_time->minute % 15);
        //時刻をフォーマット
        $start_time = $start_time->format("H:i");

        //比較するダミーの時刻を作成
        $fake_time = new Carbon('09:30:00');
        $fake_time = $fake_time->format("H:i");

        //$start_timeが9時30分以前なら9時30分で登録
        if ($start_time > $fake_time) {

            Achievement::create(
                [
                    'user_id' => $request->user_id,
                    'insert_date' => $insert_date,
                    'start_time' => $start_time,
                ],
            );
            //それ以降は$start_timeの時間で登録
        } elseif ($start_time < $fake_time) {
            Achievement::create(
                [
                    'user_id' => $request->user_id,
                    'insert_date' => $insert_date,
                    'start_time' => $fake_time,
                ],
            );
        }
    }

    /**
     * 終了時刻を作成してレコードを更新
     *
     * @param Request $request
     * @return void
     * 現在時刻を15分切り下げてフォーマット
     * 16時以後ならダミーの時刻を登録
     */
    public function End_Time(Request $request)
    {
        //今日の登録日を取得してフォーマット
        $insert_date = Carbon::now()->format("Y-m-d");

        //現在時刻を取得して15分切り下げる
        $end_time = Carbon::now();
        $end_time->subMinutes($end_time->minute % 15);

        //時刻をフォーマット
        $end_time = $end_time->format("H:i");

        //比較するダミーの時刻を作成
        $fake_time = new Carbon('16:00:00');
        $fake_time = $fake_time->format("H:i");

        //終了時刻が16時より前なら$end_timeで終了時間を登録
        if ($end_time < $fake_time) {
            Achievement::where('user_id', $request->user_id)
                ->where('insert_date', $insert_date)
                ->update(
                    ['end_time' => $end_time,]
                );
            //16時以降なら16時で登録
        } elseif ($end_time > $fake_time) {
            Achievement::where('user_id', $request->user_id)
                ->where('insert_date', $insert_date)
                ->update(
                    ['end_time' => $fake_time,]
                );
        }
    }

    /**
     * 食事提供加算を更新
     *
     * @param Request $request
     * @return void
     */
    public function Food_Up(Request $request)
    {
        //今日の登録日を取得してフォーマット
        $insert_date = Carbon::now()->format("Y-m-d");

        Achievement::where('user_id', $request->user_id)
            ->where('insert_date', $insert_date)
            ->update(
                ['food' => $request->food,]
            );
    }

    /**
     * 施設外支援を更新
     *
     * @param Request $request
     * @return void
     */
    public function Outside_Up(Request $request)
    {
        //今日の登録日を取得してフォーマット
        $insert_date = Carbon::now()->format("Y-m-d");

        Achievement::where('user_id', $request->user_id)
            ->where('insert_date', $insert_date)
            ->update(
                ['outside_support' => $request->outside,]
            );
    }

    /**
     * 医療連携体制加算を更新
     *
     * @param Request $request
     * @return void
     */
    public function Medical_Up(Request $request)
    {
        //今日の登録日を取得してフォーマット
        $insert_date = Carbon::now()->format("Y-m-d");

        Achievement::where('user_id', $request->user_id)
            ->where('insert_date', $insert_date)
            ->update(
                ['medical_support' => $request->medical,]
            );
    }

    /**
     * 備考を更新
     *
     * @param Request $request
     * @return void
     */
    public function Note_Up(Request $request)
    {
        //今日の登録日を取得してフォーマット
        $insert_date = Carbon::now()->format("Y-m-d");

        Achievement::where('user_id', $request->user_id)
            ->where('insert_date', $insert_date)
            ->update(
                ['note' => $request->note,]
            );
    }
}
