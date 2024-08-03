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
Route::get('/jira', 'JiraController@index')->name('jira.index');
Route::post('jira/import', 'JiraController@import')->name('jira.import');
Route::post('/jira/delete', 'JiraController@delete')->name('jira.delete');

// Page Service Request
Route::get('/service', 'ServiceController@index')->name('service.index');
Route::post('/service/import', 'ServiceController@import')->name('service.import');
Route::post('/service/delete', 'ServiceController@delete')->name('service.delete');

// Page weekly Powerpoint
Route::get('/weekly', 'WeeklyController@index')->name('weekly.index');
Route::get('/weekly/download', 'WeeklyController@download')->name('weekly.download');

// Page monthly Powerpoint
Route::get('/monthly', 'MonthlyController@index')->name('monthly.index');
Route::get('/monthly/download', 'MonthlyController@download')->name('monthly.download');

