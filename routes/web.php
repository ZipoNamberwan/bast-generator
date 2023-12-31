<?php

use App\Http\Controllers\GeneratorController;
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

Route::get('/generate', [App\Http\Controllers\GeneratorController::class, 'generate']);
Route::get('/generate-juli', [App\Http\Controllers\GeneratorController::class, 'generateJuli']);
Route::get('/generate-pml', [App\Http\Controllers\GeneratorController::class, 'generatePML']);
Route::get('/test', [App\Http\Controllers\GeneratorController::class, 'test']);
