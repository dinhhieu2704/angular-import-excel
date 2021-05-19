import { Component, ViewChild, ElementRef } from '@angular/core';
import * as XLSX from 'xlsx';
import { Subject } from 'rxjs/Subject';

@Component({
  selector: 'my-app',
  template: `
    <input type="file" (change)="onChange($event)" #inputFile />
    <button (click)="removeData()">Remove Data</button>
    <div *ngIf="isExcelFile === false">
      This is not an Excel file
    </div>
    <span *ngIf="spinnerEnabled" class="k-i-loading k-icon"></span>
    <table>
      <th *ngFor="let key of keys">
        {{ key }}
      </th>
      <tr *ngFor="let item of (dataSheet | async)">
        <td *ngFor="let key of keys">
          {{ item[key] }}
        </td>
      </tr>
    </table>
  `
})
export class AppComponent {
  spinnerEnabled = false;
  keys: string[];
  dataSheet = new Subject();
  @ViewChild('inputFile') inputFile: ElementRef;
  isExcelFile: boolean;

  onChange(evt) {
    let data, header;
    const target: DataTransfer = <DataTransfer>evt.target;
    this.isExcelFile = !!target.files[0].name.match(/(.xls|.xlsx)/);
    if (target.files.length > 1) {
      this.inputFile.nativeElement.value = '';
    }
    if (this.isExcelFile) {
      this.spinnerEnabled = true;
      const reader: FileReader = new FileReader();
      reader.onload = (e: any) => {
        /* read workbook */
        const bstr: string = e.target.result;
        const wb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });

        /* grab first sheet */
        const wsname: string = wb.SheetNames[0];
        const ws: XLSX.WorkSheet = wb.Sheets[wsname];

        /* save data */
        data = XLSX.utils.sheet_to_json(ws);
        console.log(data);
      };

      reader.readAsBinaryString(target.files[0]);

      reader.onloadend = e => {
        this.spinnerEnabled = false;
        this.keys = Object.keys(data[0]);
        this.dataSheet.next(data);
      };
    } else {
      this.inputFile.nativeElement.value = '';
    }
  }

  removeData() {
    this.inputFile.nativeElement.value = '';
    this.dataSheet.next(null);
    this.keys = null;
  }
}
