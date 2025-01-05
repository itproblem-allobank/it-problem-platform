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
        Schema::create('service_data', function (Blueprint $table) {
            $table->integer('id', true);
            $table->string('issue_type')->nullable();
            $table->string('code_jira')->nullable();
            $table->text('summary')->nullable();
            $table->string('assignee')->nullable();
            $table->string('reporter')->nullable();
            $table->string('status')->nullable();
            $table->timestamp('created')->nullable();
            $table->timestamp('updated')->nullable();
            $table->string('priority')->nullable();
            $table->string('sub_category')->nullable();
            $table->string('ticket_number')->nullable();
            $table->timestamp('updated_at')->nullable();
            $table->timestamp('created_at')->nullable();
        });
    }

    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::dropIfExists('service_data');
    }
};
