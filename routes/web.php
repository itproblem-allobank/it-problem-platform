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

Route::get('/profile', 'ProfileController@index')->name('profile');
Route::put('/profile', 'ProfileController@update')->name('profile.update');

// Page Jiras
Route::get('/jira', 'Problem\JiraController@index')->name('jira.index');
Route::post('jira/import', 'Problem\JiraController@import')->name('jira.import');
Route::post('/jira/delete', 'Problem\JiraController@delete')->name('jira.delete');

// Page Service Request
Route::get('/service', 'Problem\ServiceController@index')->name('service.index');
Route::post('/service/import', 'Problem\ServiceController@import')->name('service.import');
Route::post('/service/delete', 'Problem\ServiceController@delete')->name('service.delete');

// Page weekly Powerpoint
Route::get('/weekly', 'Problem\WeeklyController@index')->name('weekly.index');
Route::get('/weekly/download', 'Problem\WeeklyController@download')->name('weekly.download');

// Page monthly Powerpoint
Route::get('/monthly', 'Problem\MonthlyController@index')->name('monthly.index');
Route::get('/monthly/download', 'Problem\MonthlyController@download')->name('monthly.download');

