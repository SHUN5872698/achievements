@extends('layouts.ach_app')
@section('title', '利用者選択画面')
@yield('css')
@section('content')
<div class="container">
    <div class="row justify-content-center">
        <div class="col-md-8">
            <div class="card">
                <div class="card-header">
                    <div class="row">
                        <p class="title mt-2 ml-2">利用者確認ページ</p>
                        <ul class="navbar-nav ml-auto">
                            <li class="nav-item">
                                <a class="nav-link mr-3" href="{{ route('login') }}">管理者ページ</a>
                            </li>
                        </ul>
                    </div>
                </div>
                <div class="card-body">
                    <form method="GET" action="/achievement">
                        @csrf
                        <div class="form-group row">
                            <label for="school-1" class="col-md-4 col-form-label text-md-center">本校の方はこちら</label>

                            <div class="col-md-6">
                                <select class="form-control" name="user_id">
                                    @foreach ($school_1 as $item1)
                                    <option value={{$item1->id}}>{{$item1->first_name}}　{{$item1->last_name}}</option>
                                    @endforeach
                                </select>
                                <input type="hidden" name="month" value={{$bmonth}}>
                                <input type="submit" value="出欠ページへ" class="button">
                            </div>
                        </div>
                    </form>


                    <form method="GET" action="/achievement">
                        @csrf
                        <div class="form-group row">
                            <label for="school-2" class="col-md-4 col-form-label text-md-center">２校の方はこちら</label>
                            <div class="col-md-6">
                                <select class="form-control" name="user_id">
                                    @foreach ($school_2 as $item2)
                                    <option value={{$item2->id}}>{{$item2->first_name}}　{{$item2->last_name}}</option>
                                    @endforeach
                                </select>
                                <input type="hidden" name="month" value={{$bmonth}}>
                                <input type="submit" value="出欠ページへ" class="button">
                            </div>
                        </div>
                    </form>
                </div>
            </div>
        </div>
    </div>
</div>
@endsection
