@extends('layouts.ach_app')
@section('title', '実績画面')
@yield('css')
@section('content')
<div class="row my-3">
    <div class="col-1"></div>
    <div class="col-md-10">
        <div class="card">
            <nav class="navbar navbar-expand-lg navbar-light bg-light">
                <div class="collapse navbar-collapse" id="navbarSupportedContent">
                    <!-- Left Side Of Navbar -->
                    <ul class="navbar-nav mr-auto">
                        <a class="navbar-brand" href="/achievement?user_id={{$user->id}}&month=0">実績ページ</a>
                    </ul>
                    <!-- Right Side Of Navbar -->
                    <ul class="navbar-nav ml-auto mt-2">
                        <li class="nav-item active mt-2 mr-3">
                            <p class="user_name">利用者名： {{$user->first_name}}{{$user->last_name}} </p>
                        </li>
                        <li class="nav-item">
                            <a class="nav-link" href="/">ログアウト</a>
                        </li>
                    </ul>
                </div>
            </nav>

            @if (is_null($one_record))
            <div class="alert alert-danger mt-4 mx-3" role="alert">
                <div class="row">
                    <b class="mt-3 mx-2 mr-auto">今日の開始時間が登録されていません</b>
                    <a href="/new_record?user_id={{$user->id}}" class="btn btn-primary my-2 mr-5">出席</a>

                </div>
            </div>
            @else
            <div class="card mt-4 mx-3">
                <h5 class="card-header mb-3">本日の利用状況</h5>
                <div class="card-body">
                    <table class="table table-bordered">
                        <thead>
                            <tr align="center">
                                <th>開始時間</th>
                                <th>終了時間</th>
                                <th>食事提供加算</th>
                                <th>施設外支援</th>
                                <th>医療連携体制加算</th>
                                <th>備考</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr align="center">
                                <td>
                                    <div class="start mt-4">{{substr($one_record->start_time,0, 5)}}</div>
                                </td>
                                @if (is_null($one_record->end_time))
                                <td>
                                    <div class="end mt-3"><a href="/end_time?user_id={{$user->id}}"
                                            class="btn btn-outline-danger">退勤</a></div>
                                </td>
                                @else
                                <td>
                                    <div class="end mt-4">{{substr($one_record->end_time,0, 5)}}</div>
                                </td>
                                @endif
                                <td>
                                    <form action="/food_up" method="get">
                                        @csrf
                                        <div class="food_check mt-2">
                                            <label class="mr-3">
                                                <input class="" type="radio" name="food" value=0 @if($one_record->food
                                                == 0) checked @endif>無
                                            </label>
                                            <label>
                                                <input class="" type="radio" name="food" value=1 @if($one_record->food
                                                == 1) checked @endif>有
                                            </label>
                                        </div>
                                        <div class="food_up">
                                            <input type="hidden" name="user_id" value={{$user->id}}>
                                            <input type="submit" class="btn btn-outline-secondary btn-sm" value="変更">
                                        </div>
                                    </form>
                                </td>
                                <td>
                                    <form action="/outside_up" method="get">
                                        @csrf
                                        <div class="food_check mt-2">
                                            <label class="mr-3">
                                                <input class="" type="radio" name="outside" value=0
                                                    @if($one_record->outside_support
                                                == 0) checked @endif>無
                                            </label>
                                            <label>
                                                <input class="" type="radio" name="outside" value=2
                                                    @if($one_record->outside_support
                                                == 2) checked @endif>有
                                            </label>
                                        </div>
                                        <div class="outside_up">
                                            <input type="hidden" name="user_id" value={{$user->id}}>
                                            <input type="submit" class="btn btn-outline-secondary btn-sm" value="変更">
                                        </div>
                                    </form>
                                </td>
                                <td>
                                    <form action="/medical_up" method="get">
                                        @csrf
                                        <div class="food_check mt-2">
                                            <label class="mr-3">
                                                <input class="" type="radio" name="medical" value=0
                                                    @if($one_record->medical_support
                                                == 0) checked @endif>無
                                            </label>
                                            <label>
                                                <input class="" type="radio" name="medical" value=2
                                                    @if($one_record->medical_support
                                                == 2) checked @endif>有
                                            </label>
                                        </div>
                                        <div class="medical_up">
                                            <input type="hidden" name="user_id" value={{$user->id}}>
                                            <input type="submit" class="btn btn-outline-secondary btn-sm" value="変更">
                                        </div>
                                    </form>
                                </td>
                                @if ($one_record->note == null)
                                <td>
                                    <form action="/note_up" method="get">
                                        @csrf
                                        <select class="note my-1" name="note">
                                            <option value=""></option>
                                            <option value="通所">通所　</option>
                                            <option value="Skype">Skype　</option>
                                            <option value="メール">メール　</option>
                                            <option value="訪問">訪問　</option>
                                        </select>
                                        <div class="note_up my-1">
                                            <input type="hidden" name="user_id" value={{$user->id}}>
                                            <input type="submit" class="btn btn-outline-secondary btn-sm" value="変更">
                                        </div>
                                    </form>
                                </td>

                                @elseif($one_record->note == "通所")
                                <td>
                                    <form action="/note_up" method="get">
                                        @csrf
                                        <select class="note my-1" name="note">
                                            <option value={{$one_record->note}}>{{$one_record->note}}　</option>
                                            <option value="Skype">Skype　</option>
                                            <option value="メール">メール　</option>
                                            <option value="訪問">訪問　</option>
                                        </select>
                                        <div class="note_up my-1">
                                            <input type="hidden" name="user_id" value={{$user->id}}>
                                            <input type="submit" class="btn btn-outline-secondary btn-sm" value="変更">
                                        </div>
                                    </form>
                                </td>

                                @elseif($one_record->note == "Skype")
                                <td>
                                    <form action="/note_up" method="get">
                                        @csrf
                                        <select class="note my-1" name="note">
                                            <option value={{$one_record->note}}>{{$one_record->note}}　</option>
                                            <option value="通所">通所　</option>
                                            <option value="メール">メール　</option>
                                            <option value="訪問">訪問　</option>
                                        </select>
                                        <div class="note_up my-1">
                                            <input type="hidden" name="user_id" value={{$user->id}}>
                                            <input type="submit" class="btn btn-outline-secondary btn-sm" value="変更">
                                        </div>
                                    </form>
                                </td>

                                @elseif($one_record->note == "メール")
                                <td>
                                    <form action="/note_up" method="get">
                                        @csrf
                                        <select class="note my-1" name="note">
                                            <option value={{$one_record->note}}>{{$one_record->note}}　</option>
                                            <option value="通所">通所　</option>
                                            <option value="Skype">Skype　</option>
                                            <option value="訪問">訪問　</option>
                                        </select>
                                        <div class="note_up my-1">
                                            <input type="hidden" name="user_id" value={{$user->id}}>
                                            <input type="submit" class="btn btn-outline-secondary btn-sm" value="変更">
                                        </div>
                                    </form>
                                </td>

                                @elseif($one_record->note == "訪問")
                                <td>
                                    <form action="/note_up" method="get">
                                        @csrf
                                        <select class="note my-1" name="note">
                                            <option value={{$one_record->note}}>{{$one_record->note}}　</option>
                                            <option value="通所">通所　</option>
                                            <option value="Skype">Skype　</option>
                                            <option value="メール">メール　</option>
                                        </select>
                                        <div class="note_up my-1">
                                            <input type="hidden" name="user_id" value={{$user->id}}>
                                            <input type="submit" class="btn btn-outline-secondary btn-sm" value="変更">
                                        </div>
                                    </form>
                                </td>
                                @endif
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>
            @endif

            <div class="card  my-4 mx-3">
                <div class="card-header mb-2">
                    <div class="row">
                        <ul class="nav card-header-tabs mr-auto">
                            <li class="user_name mx-3">
                                <font class="name_title">支給決定障害者氏名</font>
                                <br>
                                <font class="name">{{$user->first_name}} {{$user->last_name}}</font>
                            </li>
                            <li class="month-select mx-1 mt-2">
                                <form action="/achievement" method="get">
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
                            <li class="business_operator mx-3">
                                <font class="business">事業者及びその事業所</font>
                                <br>
                                @if ($user->school_id == 1)
                                <font class="school_name">未来のかたち　本町本校</font>
                                @elseif($user->school_id == 2)
                                <font class="school_name">未来のかたち　本町2校</font>
                                @endif
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

                                @endif
                            </tr>
                            @endforeach
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
    <div class="col-1"></div>
</div>
@endsection
