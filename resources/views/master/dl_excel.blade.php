@extends('layouts.master_app')
@section('title', '実績編集')
@yield('css')
@section('content')
<div class="row justify-content-center my-3">
    <div class="col-md-10">
        <div class="card">
            <div class="card-header">
                <nav class="navbar navbar-expand-lg navbar-light bg-light">
                    <a class="navbar-brand" href="/master">
                        <font class="master_title">実績を一括出力</font>
                    </a>
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
            <div class="row justify-content-center">
                <div class="card-body">
                    <form method="POST" action="/bulk_creation">
                        @csrf

                        <div class="form-group row">
                            @if ($user->school_id == 1)
                            <label for="school_name" class="col-md-5 col-form-label text-md-right">未来のかたち本町本校</label>
                            @elseif ($user->school_id == 2)
                            <label for="school_name" class="col-md-5 col-form-label text-md-right">未来のかたち本町２校</label>
                            @endif


                            <div class="col-md-5">
                            </div>
                        </div>

                        <div class="form-group row">
                            <label for="start_time" class="col-md-5 col-form-label text-md-right">
                                月を選択してください
                            </label>

                            <div class="col-md-4 mt-1">
                                <select class="past my-1" name="month">
                                    @foreach ($months as $month)
                                    <option value={{$month}}>{{$month->isoformat('Y年M月')}}</option>
                                    @endforeach
                                </select>
                            </div>
                        </div>
                        <input type="hidden" name="school_id" value={{$user->school_id}}>
                        <div class="form-group row mb-0">
                            <div class="col-md-6 offset-md-4">
                                <button type="submit" class="btn btn-primary">ダウンロード</button>
                                <button type="button" class="btn btn-outline-primary mx-3" onclick="location.href='/master';">キャンセル</button>
                            </div>
                        </div>
                    </form>
                </div>

            </div>
        </div>
    </div>
</div>
@endsection
