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
            $table->integer('id', true);
            $table->string('code_jira')->nullable();
            $table->string('environment')->nullable();
            $table->string('problem')->nullable();
            $table->string('category')->nullable();
            $table->text('summary')->nullable();
            $table->string('zentao_link')->nullable();
            $table->string('priority')->nullable();
            $table->string('status')->nullable();
            $table->text('pending_reason')->nullable();
            $table->string('target_version')->nullable();
            $table->text('impact_analyst')->nullable();
            $table->text('root_cause')->nullable();
            $table->text('work_around')->nullable();
            $table->string('reporter')->nullable();
            $table->string('assignee_to')->nullable();
            $table->string('nickname', 100)->nullable();
            $table->text('description')->nullable();
            $table->string('frequent', 100)->nullable();
            $table->string('complain_info')->nullable();
            $table->timestamp('created')->nullable();
            $table->timestamp('updated')->nullable();
            $table->timestamp('changed_at')->nullable();
            $table->timestamp('rca_time')->nullable();
            $table->timestamp('closed_time')->nullable();
            $table->timestamp('updated_at')->nullable();
            $table->timestamp('created_at')->nullable();
            $table->integer('resolved_days')->nullable();
            $table->integer('rca_days')->nullable();
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
