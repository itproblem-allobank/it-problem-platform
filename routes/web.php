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

// ---------------------------- IT Problem Page -----------------------------
// Add Ticket Problem
Route::get('/problem', 'Problem\ProblemController@index')->name('problem.index');
Route::post('problem/import', 'Problem\ProblemController@import')->name('problem.import');
Route::post('/problem/delete', 'Problem\ProblemController@delete')->name('problem.delete');
// Add Service Request
Route::get('/service', 'Problem\ServiceController@index')->name('service.index');
Route::post('/service/import', 'Problem\ServiceController@import')->name('service.import');
Route::post('/service/delete', 'Problem\ServiceController@delete')->name('service.delete');
// Generate Weekly
Route::get('/p-weekly', 'Problem\WeeklyController@index')->name('p-weekly.index');
Route::get('/p-weekly/download', 'Problem\WeeklyController@download')->name('p-weekly.download');
// Generate Monthly
Route::get('/p-monthly', 'Problem\MonthlyController@index')->name('p-monthly.index');
Route::get('/p-monthly/download', 'Problem\MonthlyController@download')->name('p-monthly.download');

// ---------------------------- IT Incident Page ------------------------------
// Add Ticket Incident
Route::get('/incident', 'Incident\IncidentController@index')->name('incident.index');
Route::post('/incident/import', 'Incident\IncidentController@import')->name('incident.import');
Route::post('/incident/delete', 'Incident\IncidentController@delete')->name('incident.delete');
// Generate Weekly
Route::get('/i-weekly', 'Incident\WeeklyController@index')->name('i-weekly.index');
Route::get('/i-weekly/download', 'Incident\WeeklyController@download')->name('i-weekly.download');
// Generate Monthly
Route::get('/i-monthly', 'Incident\MonthlyController@index')->name('i-monthly.index');
Route::get('/i-monthly/download', 'Incident\MonthlyController@download')->name('i-monthly.download');
