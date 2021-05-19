import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { BrowserAnimationsModule } from '@angular/platform-browser/animations';
import { FormsModule } from '@angular/forms';
import { GridModule } from '@progress/kendo-angular-grid';

import { AppComponent } from './app.component';

@NgModule({
  imports: [ BrowserModule, BrowserAnimationsModule, FormsModule, GridModule ],
  declarations: [ AppComponent ],
  bootstrap: [ AppComponent ]
})

export class AppModule { }
