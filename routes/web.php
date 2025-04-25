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
Route::get('/jira', 'Problem\ProblemController@index')->name('jira.index');
Route::post('jira/import', 'Problem\ProblemController@import')->name('jira.import');
Route::post('/jira/delete', 'Problem\ProblemController@delete')->name('jira.delete');
// Add Service Request
Route::get('/service', 'Problem\ServiceController@index')->name('service.index');
Route::post('/service/import', 'Problem\ServiceController@import')->name('service.import');
Route::post('/service/delete', 'Problem\ServiceController@delete')->name('service.delete');
// Generate Weekly
Route::get('/weekly', 'Problem\WeeklyController@index')->name('weekly.index');
Route::get('/weekly/download', 'Problem\WeeklyController@download')->name('weekly.download');
// Generate Monthly
Route::get('/monthly', 'Problem\MonthlyController@index')->name('monthly.index');
Route::get('/monthly/download', 'Problem\MonthlyController@download')->name('monthly.download');

// ---------------------------- IT Incident Page ------------------------------
// Add Ticket Incident
Route::get('/incident', 'Incident\IncidentController@index')->name('incident.index');
Route::post('/incident/import', 'Incident\IncidentController@import')->name('incident.import');
Route::post('/incident/delete', 'Incident\IncidentController@delete')->name('incident.delete');

