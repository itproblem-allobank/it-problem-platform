<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

return new class extends Migration
{
    /**
     * Run the migrations.
     */
    public function up(): void
    {
        Schema::create('excel_data', function (Blueprint $table) {
            $table->increments('id');
            $table->string('code_jira');
            $table->string('environment');
            $table->string('problem');
            $table->string('category');
            $table->string('summary');
            $table->string('zentao_link');
            $table->string('priority');
            $table->string('status');
            $table->string('pending_reason');
            $table->string('target_version');
            $table->string('impact_analyst');
            $table->string('root_cause');
            $table->string('work_around');
            $table->string('reporter');
            $table->string('assignee_to');
            $table->string('created');
            $table->string('updated');
            $table->string('changed_at');
        });
    }

    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::dropIfExists('excel_data');
    }
};
