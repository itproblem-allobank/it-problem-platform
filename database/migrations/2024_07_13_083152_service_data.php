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
            $table->increments('id');
            $table->string('issue_type');
            $table->string('code_jira');
            $table->string('summary');
            $table->string('assignee');
            $table->string('reporter');
            $table->string('status');
            $table->timestamp('created');
            $table->timestamp('updated');
            $table->string('priority');
            $table->string('sub_category');
            $table->string('ticket_number');

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
