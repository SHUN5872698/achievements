<?php

namespace App;

use Illuminate\Contracts\Auth\MustVerifyEmail;
use Illuminate\Foundation\Auth\User as Authenticatable;
use App\Achievement;
use Illuminate\Http\Request;
use Carbon\Carbon;
use Illuminate\Notifications\Notifiable;

class Master extends Authenticatable
{
    use Notifiable;

    /**
     * The attributes that are mass assignable.
     *
     * @var array
     */
    protected $fillable = [
        'name', 'email', 'password',
    ];

    /**
     * The attributes that should be hidden for arrays.
     *
     * @var array
     */
    protected $hidden = [
        'password', 'remember_token',
    ];

    /**
     * The attributes that should be cast to native types.
     *
     * @var array
     */
    protected $casts = [
        'email_verified_at' => 'datetime',
    ];

    /**
     * 利用者一覧と当月の実績データと実績表に必要なデータを取得
     * @pahram Request $request
     * @return void
     * 在籍校別で利用者情報を取得する
     */
    public function Records(Request $request)
    {
        //本校の利用者情報の取得
        $school_1 = new User();
        $school_1 = $school_1->School_1();

        //2校の利用者情報の取得
        $school_2 = new User();
        $school_2 = $school_2->School_2();

        //利用者の情報をを取得
        $user = new User();
        $user = $user->getUser($request);

        //月初を取得
        $bmonth =  new Carbon($request->month);

        //当月の日数を取得
        $days = new Achievement();
        $days = $days->M_Days($request);

        //今月から過去1年間の月を取得
        $months = new Achievement();
        $months = $months->Months($request);

        //利用者の月の実績データの取得
        $records = new Achievement();
        $records = $records->Month_Records($request);

        $data = [
            'school_1' => $school_1,
            'school_2' => $school_2,
            'user' => $user,
            'bmonth' => $bmonth,
            'days' => $days,
            'months' => $months,
            'records' => $records,
        ];
        return ($data);
    }

    public function DRecords(Request $request)
    {
        //本校の利用者情報の取得
        $school_1 = new User();
        $school_1 = $school_1->School_1();

        //2校の利用者情報の取得
        $school_2 = new User();
        $school_2 = $school_2->School_2();

        //利用者の情報をを取得
        $user = new User();
        $user = $user->getUser($request);

        //月初を取得
        $bmonth =  new Carbon($request->insert_date);

        $dmonth = new Master();
        $dmonth = $dmonth->D_Month($request);

        //当月の日数を取得
        $days = new Master();
        $days = $days->M_Days($request);

        $records = new Master();
        $records = $records->M_Days($request);

        //今月から過去1年間の月を取得
        $months = new Master();
        $months = $months->Months($request);

        //利用者の月の実績データの取得
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

        $data = [
            'school_1' => $school_1,
            'school_2' => $school_2,
            'user' => $user,
            'bmonth' => $bmonth,
            'days' => $days,
            'months' => $months,
            'records' => $records,
        ];
        return ($data);
    }

    /**
     * 月の日数が何日かを取得
     */
    public function D_Month(Request $request)
    {
        $dmonth = new Carbon($request->insert_date);
        $dmonth = $dmonth->daysInMonth;
        return $dmonth;
    }

    /**
     * 一ヶ月分の日数を取得
     */
    public function M_Days(Request $request)
    {
        //月初を取得
        $bmonth = new Carbon($request->insert_date);
        $bmonth = $bmonth->firstOfMonth();

        //月の日数が何日かを取得
        $dmonth = new Master();
        $dmonth = $dmonth->D_Month($request);

        //日付を連想配列に格納
        for ($i = 0; $i < $dmonth; $i++) {
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
        $submonth = new Carbon($request->insert_date);

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
     * 今月から過去1年間の月を取得
     *
     * @return void
     */
    public function Exele_Months()
    {
        //今月の月初を取得
        $bmonth = Carbon::now()->firstOfMonth();

        //一年分の月初を配列に格納
        for ($i = 0; $i < 12; $i++) {
            $months[$i] = $bmonth->copy()->subMonth($i);
        }
        return $months;
    }

    /**
     * Excelの日付、曜日、サービス提供欄出力用の連想配列を作成
     *
     */
    public function Excel_Days(Request $request)
    {
        // 一ヶ月分の日数を取得
        $mdays = new Achievement();
        $mdays = $mdays->M_Days($request);

        //mdaysをネストして新たな配列に
        //0 => array:3 [▼
        //  0 => "日付、"
        //  1 => "曜日"
        //  2 => "欠or空文字
        //を格納
        foreach ($mdays as $day) {
            if ($day->isSaturday() || $day->isSunday()) {
                $days[] = array(
                    $day->isoFormat('D日'),
                    $day->isoFormat('ddd'),
                    ""
                );
            } else {
                $days[] = array(
                    $day->isoFormat('D日'),
                    $day->isoFormat('ddd'),
                    "欠"
                );
            }
        }
        return $days;
    }

    //一ヶ月分の実績データの配列を作成
    public function Month_Records(Request $request)
    {
        //月初を取得
        $bmonth = new Carbon($request->month);

        //月の日数が何日かを取得
        $dmonth = new Achievement();
        $dmonth = $dmonth->D_Month($request);

        //一ヶ月分の日数を取得
        $mdays = new Achievement();
        $mdays = $mdays->M_Days($request);

        //中身がnullの配列を作成
        $records = null;

        //配列を日数分作成
        for ($n = 0; $n < $dmonth; $n++) {
            $records[][$n] = null;
        }

        //利用者の一ヶ月間のレコードを取得
        $achievements = Achievement::where('achievements.user_id', $request->user_id)
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

        //実績データの登録日と一ヶ月の日数を比較して一致した場合$recordsの配列にレコードを上書きする
        foreach ($achievements as $achievement) {
            //開始時間と終了時間を切り取って変換
            $achievement['start_time'] = substr($achievement['start_time'], 0, 5);
            $achievement['end_time'] = substr($achievement['end_time'], 0, 5);
            for ($n = 0; $n < $dmonth; $n++) {
                if ($mdays[$n] == $achievement['insert_date']) {
                    $records[$n] = $achievement;
                }
            }
        }
        return $records;
    }
}
