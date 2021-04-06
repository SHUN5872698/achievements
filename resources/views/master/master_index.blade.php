@extends('layouts.master_app')
@section('css')
<link rel="stylesheet" href="{{ asset('js/jquery-ui-1.12.1/jquery-ui.min.css') }}">
<link rel="stylesheet" href="{{ asset('js/jquery-ui-1.12.1/jquery-ui.structure.min.css') }}">
<link rel="stylesheet" href="{{ asset('js/jquery-ui-1.12.1/jquery-ui.theme.min.css') }}">
@yield('css')
<style>
    .ui-dialog-titlebar {
        color: white;
        background: red;
    }

</style>
@section('title', '管理者実績閲覧')
@section('content')
<div class="row justify-content-center my-3">
    <div class="col-md-10">
        <div class="card">
            <div class="card-header">
                <nav class="navbar navbar-expand-lg navbar-light bg-light">
                    <font class="master_title">実績閲覧</font>
                    <div class="collapse navbar-collapse" id="navbarSupportedContent">
                        <!-- Left Side Of Navbar -->
                        <ul class="navbar-nav mr-auto"></ul>
                        <!-- Right Side Of Navbar -->
                        <ul class="navbar-nav ml-auto">
                            <!-- Authentication Links -->
                            @guest
                            <li class="nav-item">
                                <a class="nav-link" href="{{ route('login') }}">{{ __('Login') }}</a>
                            </li>
                            @if (Route::has('register'))
                            <li class="nav-item">
                                <a class="nav-link" href="{{ route('register') }}">{{ __('Register') }}</a>
                            </li>
                            @endif
                            @else
                            <li class="nav-item dropdown">
                                <a id="navbarDropdown" class="nav-link dropdown-toggle mt-1 mx-2" href="#" role="button" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false" v-pre>
                                    管理者機能一覧<span class="caret"></span>
                                </a>

                                <div class="dropdown-menu dropdown-menu-right" aria-labelledby="navbarDropdown">
                                    <a class="dropdown-item text-black" href="/master">実績閲覧</a>
                                    <a class="dropdown-item text-black" href="/add_user">新規利用者登録</a>
                                    <a class="dropdown-item text-black" href="/edit_user">利用者情報の編集</a>
                                    <a class="dropdown-item" href="{{ route('logout') }}" onclick="event.preventDefault();
                                                             document.getElementById('logout-form').submit();">
                                        {{ __('Logout') }}
                                    </a>

                                    <form id="logout-form" action="{{ route('logout') }}" method="POST" style="display: none;">
                                        @csrf
                                    </form>
                                </div>
                            </li>
                            @endguest
                        </ul>
                    </div>
                </nav>
            </div>
            <div class="card-body">
                <form method="GET" action="/check_records">
                    @csrf
                    <div class="form-group row">
                        <label for="school-1" class="col-md-4 col-form-label text-md-center">
                            <font class="school_1">本校の利用者</font>
                            <br>
                            <a class="dl_excel" href="/dl_school1">全員分をダウンロード</a>
                        </label>
                        <div class="col-md-4">
                            <select class="form-control" name="user_id">
                                @foreach ($school_1 as $sch1)
                                <option value={{$sch1->id}}>{{$sch1->first_name}}　{{$sch1->last_name}}</option>
                                @endforeach
                            </select>
                            <input type="hidden" name="month" value={{$bmonth}}>
                            <input type="submit" value="実績表示" class="button">
                        </div>
                    </div>
                </form>

                <form method="GET" action="/check_records">
                    @csrf
                    <div class="form-group row">
                        <label for="school-2" class="col-md-4 col-form-label text-md-center">
                            <font class="school_2">２校の利用者</font>
                            <br>
                            <a class="dl_excel" href="/dl_school2">全員分をダウンロード</a>
                        </label>
                        <div class="col-md-4">
                            <select class="form-control" name="user_id">
                                @foreach ($school_2 as $sch2)
                                <option value={{$sch2->id}}>{{$sch2->first_name}}　{{$sch2->last_name}}</option>
                                @endforeach
                            </select>
                            <input type="hidden" name="month" value={{$bmonth}}>
                            <input type="submit" value="実績表示" class="button">
                        </div>
                    </div>
                </form>
            </div>
        </div>
    </div>

    @if (isset($user))
    <div class="col-10 mt-3">
        <div class="card">
            <div class="card-header mb-2">
                <div class="row">
                    <ul class="nav card-header-tabs mr-auto">
                        <li class="user_name mx-3">
                            <font class="name_title">支給決定障害者氏名</font>
                            <br>
                            <font class="name">{{$user->first_name}} {{$user->last_name}}</font>
                        </li>
                        <li class="month-select mx-1 mt-2">
                            <form action="/check_records" method="get">
                                @csrf
                                <select class="past my-1" name="month">
                                    @foreach ($months as $month)
                                    <option value={{$month}}>{{$month->isoformat('Y年M月')}}</option>
                                    @endforeach
                                </select>
                                <input type="hidden" name="user_id" value={{$user->id}}>
                                <input type="submit" class="btn btn-outline-secondary btn-sm mx-2" value="変更">
                            </form>
                        </li>
                        <li class="card-title mx-1 mt-2">
                            <font>実績記録表</font>
                        </li>
                    </ul>
                    <ul class="nav card-header-tabs ml-auto">
                        <li class="add_achievement mx-3">
                            <button type="button" class="btn btn-outline-secondary btn-sm mx-1 mt-2" onclick="location.href='/add_achievement?user_id={{$user->id}}';">実績を追加する</button>
                        </li>
                        <li class="add_achievement mr-5">
                            <button type="button" class="btn btn-outline-primary btn-sm mx-1 mt-2" onclick="location.href='one_data?user_id={{$user->id}}&month={{$bmonth}}';">ダウンロード</button>

                        </li>
                    </ul>
                </div>
            </div>
            <div class="card-body">
                <h1>{{$bmonth->isoformat('Y年M月')}}分</h1>
                <table class="table table-bordered">
                    <thead>
                        <tr class="text-center">
                            <th class="align-middle" rowspan="2">日付</th>
                            <th class="align-middle" rowspan="2">曜日</th>
                            <th colspan="5">サービス提供実績</th>
                            <th class="align-middle" rowspan="2">備考</th>
                            <th rowspan="2"></th>
                            <th rowspan="2"></th>
                        </tr>
                        <tr class="text-center">
                            <th>開始時間</th>
                            <th>終了時間</th>
                            <th>食事提供加算</th>
                            <th>施設外支援</th>
                            <th>医療連携体制加算</th>
                        </tr>
                    </thead>
                    <tbody>
                        @foreach (array_map(null, $days, $records) as [$day, $record])
                        <tr align="center">
                            @if ($day == $record)
                            <td>{{$day->isoformat('D')}}</td>
                            <td>{{$day->isoformat('ddd')}}</td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td></td>
                            @else
                            <td>{{$day->isoformat('D')}}</td>
                            <td>{{$day->isoformat('ddd')}}</td>
                            <td>{{substr($record->start_time,0, 5)}}</td>
                            <td>{{substr($record->end_time,0, 5)}}</td>
                            @if ($record->food == 0)
                            <td>無</td>
                            @elseif ($record->food == 1)
                            <td>
                                <font color="red">有</font>
                            </td>
                            @endif

                            @if ($record->outside_support == 0)
                            <td>無</td>
                            @elseif ($record->outside_support == 2)
                            <td>
                                <font color="red">有</font>
                            </td>
                            @endif

                            @if ($record->medical_support == 0)
                            <td>無</td>
                            @elseif ($record->medical_support == 2)
                            <td>
                                <font color="red">有</font>
                            </td>
                            @endif

                            @if ($record->note)
                            <td>{{$record->note}}</td>
                            @else
                            <td></td>
                            @endif
                            <td>
                                <button type="button" class="btn btn-primary btn-sm" onclick="location.href='/edit_achievement?id={{$record->id}}&user_id={{$record->user_id}}';">編集</button>
                            </td>
                            <td>
                                <input type="button" class="btn btn-danger btn-sm" onclick="dAchievement({{$record->id }}, '{{ $record->insert_date->isoformat('Y年M月D日') }}');" value="削除">
                            </td>
                            @endif
                        </tr>
                        @endforeach
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    @endif
</div>
<script src="{{ asset('js/jquery-ui-1.12.1/external/jquery/jquery.js') }}"></script>
<script src="{{ asset('js/jquery-ui-1.12.1/jquery-ui.min.js') }}"></script>
<script>
    $(function () {
        $("#dialog-confirm").hide();
    });

    function dAchievement(id, insert_date) {
        $("#dAchievement").text(insert_date);
        $("#dialog-confirm").dialog({
            resizable: false,
            height: "auto",
            width: 400,
            modal: true,
            buttons: {
                "削除": function () {
                    $(this).dialog("close");
                    location.href = '/del_achievement?id=' + id;
                },
                "キャンセル": function () {
                    $(this).dialog("close");
                }
            }
        });
    }

</script>

<div id="dialog-confirm" title="実績を削除">
    <p><span class=" ui-icon ui-icon-alert" style="float:left; margin:3px 2px;"></span>
        下記の実績を削除してもいいですか？
        <p id="dAchievement"></p>
    </p>
</div>
@endsection
