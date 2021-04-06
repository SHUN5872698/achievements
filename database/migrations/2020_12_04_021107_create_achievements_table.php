<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class CreateAchievementsTable extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()

    {
        Schema::create('achievements', function (Blueprint $table) {
            $table->increments('id');
            $table->integer('user_id')->unsigned()->unique();
            $table->date('insert_date');
            $table->time('start_time')->nullable();
            $table->time('end_time')->nullable();
            $table->string('visit_support')->nullable();
            $table->tinyInteger('food')->default(0);
            $table->tinyInteger('outside_support')->default(0);
            $table->tinyInteger('medical_support')->default(0);
            $table->string('note')->nullable();
            $table->string('stamp')->nullable();
            $table->timestamps();

            // $table->foreign('user_id')
            // ->references('id')
            //     ->on('users')
            //     ->onDelete('cascade');
        });
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::dropIfExists('achievements');
    }
}
