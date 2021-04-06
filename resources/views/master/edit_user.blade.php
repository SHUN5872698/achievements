@extends('layouts.master_app')
@yield('css')
@section('title', '利用者情報の編集')
@section('content')
<div class="row justify-content-center my-3">
    <div class="col-md-10">
        <div class="card">
            <div class="card-header">
                <nav class="navbar navbar-expand-lg navbar-light bg-light">
                    <font class="master_title">利用者情報編集</font>
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
                                <a id="navbarDropdown" class="nav-link dropdown-toggle mt-1 mx-2" href="#" role="button"
                                    data-toggle="dropdown" aria-haspopup="true" aria-expanded="false" v-pre>
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

                                    <form id="logout-form" action="{{ route('logout') }}" method="POST"
                                        style="display: none;">
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
                <form method="GET" action="/edit_user">
                    @csrf
                    <div class="form-group row">
                        <label for="school-1" class="col-md-4 col-form-label text-md-center">本校の利用者</label>
                        <div class="col-md-4">
                            <select class="form-control" name="user_id">
                                @foreach ($school_1 as $sch1)
                                <option value={{$sch1->id}}>{{$sch1->first_name}}　{{$sch1->last_name}}</option>
                                @endforeach
                            </select>
                            <input type="hidden" name="month" value=0>
                            <input type="submit" value="選択" class="button">
                        </div>
                    </div>
                </form>

                <form method="GET" action="/edit_user">
                    @csrf
                    <div class="form-group row">
                        <label for="school-2" class="col-md-4 col-form-label text-md-center">２校の利用者</label>
                        <div class="col-md-4">
                            <select class="form-control" name="user_id">
                                @foreach ($school_2 as $sch2)
                                <option value={{$sch2->id}}>{{$sch2->first_name}}　{{$sch2->last_name}}</option>
                                @endforeach
                            </select>
                            <input type="hidden" name="month" value=0>
                            <input type="submit" value="選択" class="button">
                        </div>
                    </div>
                </form>
            </div>
        </div>
    </div>
    @if (isset($user))
    <div class="col-md-10 mt-5">
        <div class="card">
            <div class="card-header">
                <font class="master_title">利用者情報</font>
            </div>
            <div class="row justify-content-center">
                <div class="col-md-8">
                    <div class="card-body">
                        <form method="POST" action="/edit_user">
                            @csrf
                            <div class="form-group row">
                                <label for="first_name" class="col-md-4 col-form-label text-md-right">利用者の名字</label>

                                <div class="col-md-6">
                                    <input id="first_name" type="text"
                                        class="form-control @error('first_name') is-invalid @enderror" name="first_name"
                                        value="{{ empty(old('first_name')) ? ($user->first_name) : old('first_name')}}"
                                        required autocomplete="first_name" autofocus>
                                    @error('first_name')
                                    <span class="invalid-feedback" role="alert">
                                        <strong>{{ $message }}</strong>
                                    </span>
                                    @enderror
                                </div>
                            </div>


                            <div class="form-group row">
                                <label for="last_name" class="col-md-4 col-form-label text-md-right">利用者の名前</label>

                                <div class="col-md-6">
                                    <input id="last_name" type="text"
                                        class="form-control @error('last_name') is-invalid @enderror" name="last_name"
                                        value="{{ empty(old('last_name')) ? ($user->last_name) : old('last_name')}}"
                                        required autocomplete="last_name" autofocus>
                                    @error('last_name')
                                    <span class="invalid-feedback" role="alert">
                                        <strong>{{ $message }}</strong>
                                    </span>
                                    @enderror
                                </div>
                            </div>

                            <div class="form-group row">
                                <label for="full_name" class="col-md-4 col-form-label text-md-right">
                                    <font class=>利用者のフルネーム</font>
                                    <br>
                                    <font class="mr-5">(カナ)</font>
                                </label>

                                <div class="col-md-6 mt-1">
                                    <input id="full_name" type="text"
                                        class="form-control @error('full_name') is-invalid @enderror" name="full_name"
                                        value="{{ empty(old('full_name')) ? ($user->full_name) : old('full_name')}}"
                                        required autocomplete="full_name" autofocus>
                                    @error('full_name')
                                    <span class="invalid-feedback" role="alert">
                                        <strong>{{ $message }}</strong>
                                    </span>
                                    @enderror
                                </div>
                            </div>

                            <div class="form-group row">
                                <label for="age" class="col-md-4 col-form-label text-md-right">年齢</label>

                                <div class="col-md-2 mt-1">
                                    <input id="age" type="number"
                                        class="form-control @error('age') is-invalid @enderror" name="age"
                                        value="{{ empty(old('age')) ? ($user->age) : old('age')}}" required
                                        autocomplete="age" autofocus>
                                    @error('age')
                                    <span class="invalid-feedback" role="alert">
                                        <strong>{{ $message }}</strong>
                                    </span>
                                    @enderror
                                </div>
                            </div>

                            <div class="form-group row">
                                <label for="school_id" class="col-md-4 col-form-label text-md-right">所属校</label>

                                @if ($user->school_id == 1)
                                <div class="col-md-6 mt-1">
                                    <label>
                                        <input class="@error('school_id') is-invalid @enderror" type="radio"
                                            name="school_id" value=1 checked="checked" @if(old('school_id')==1) checked
                                            @endif>本校
                                    </label>
                                    <label>
                                        <input class="@error('school_id') is-invalid @enderror" type="radio"
                                            name="school_id" value=2 @if(old('school_id')==2) checked @endif>２校
                                    </label>
                                </div>
                                @elseif($user->school_id == 2)
                                <div class="col-md-6 mt-1">
                                    <label>
                                        <input class="@error('school_id') is-invalid @enderror" type="radio"
                                            name="school_id" value=1 @if(old('school_id')==1) checked @endif>本校
                                    </label>
                                    <label>
                                        <input class="@error('school_id') is-invalid @enderror" type="radio"
                                            name="school_id" value=2 checked="checked" @if(old('school_id')==2) checked
                                            @endif>２校
                                    </label>
                                </div>
                                @endif
                            </div>
                            <input type="hidden" name="user_id" value={{$user->id}}>

                            <div class="form-group row mb-0">
                                <div class="col-md-6 offset-md-4">
                                    <button type="submit" class="btn btn-primary">変更</button>
                                    <button type="button" class="btn btn-outline-primary"
                                        onclick="location.href='/master';">キャンセル</button>
                                </div>
                            </div>
                        </form>
                    </div>
                </div>
            </div>
            @endif
        </div>
    </div>
    @endsection
