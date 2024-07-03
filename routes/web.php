<?php

use Illuminate\Support\Facades\Auth;
use Illuminate\Support\Facades\Route;

/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider and all of them will
| be assigned to the "web" middleware group. Make something great!
|
*/

Route::get('/', function () {
    return view('auth/login');
});

Auth::routes();

Route::get('/home', 'HomeController@index')->name('home');

Route::get('/profile', 'ProfileController@index')->name('profile');
Route::put('/profile', 'ProfileController@update')->name('profile.update');

Route::get('/data', 'DataController@index')->name('data');
Route::get('/data/getdata', 'DataController@getdata')->name('data.getdata');
Route::post('/data/import', 'DataController@import')->name('data.import');
Route::post('/data/delete', 'DataController@delete')->name('data.delete');
Route::get('/data/cetak_pdf', 'DataController@cetak_pdf')->name('data.cetakpdf');


//Charts
Route::get('/chart/view', 'ChartsController@index')->name('chart.index');
Route::get('/chart/weekly', 'ChartsController@weekly')->name('chart.weekly');
Route::get('/chart/total', 'ChartsController@total')->name('chart.total');
Route::post('/chart/print', 'ChartsController@print')->name('chart.print');

//monthly
Route::get('/monthly', 'MonthlyController@index')->name('monthly');
Route::post('/monthly/import', 'MonthlyController@import')->name('monthly.import');
Route::post('/monthly/delete', 'MonthlyController@delete')->name('monthly.delete');
Route::get('/monthly/chart', 'MonthlyController@chart')->name('monthly.chart');
Route::get('/monthly/export', 'MonthlyController@export')->name('monthly.export');

Route::get('/about', function () {
    return view('about');
})->name('about');
