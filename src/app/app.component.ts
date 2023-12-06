import { Component } from '@angular/core';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss'],
})
export class AppComponent {
  constructor() {}
  excelData: any;
  title = 'upload-edit-download-excel';

  readExcel(event: any) {
    let file = event.target.files[0];
    let fileReader = new FileReader();
    fileReader.readAsBinaryString(file);
    fileReader.onload = (e) => {
      var workBook = XLSX.read(fileReader.result, { type: 'binary' });
      var sheetName = workBook.SheetNames;
      this.excelData = XLSX.utils.sheet_to_json(workBook.Sheets[sheetName[0]]);

      this.excelData.map((response: any) => {
        let newCods = this.ParseDMS(
          response.Latitude + ' ' + response.Longitude
        );
        (response.decLatitude = newCods?.lat),
          (response.decLongitude = newCods?.lng);
      });
    };
  }

  // parse your input
  ParseDMS(input: any) {
    var parts = input.split(/[^\d\w]+/);
    var lat = this.ConvertDMSToDD(
      parts[1],
      parts[2],
      (Number(parts[3]) / 100) * 60,
      parts[0]
    );
    var lng = this.ConvertDMSToDD(
      parts[5],
      parts[6],
      (Number(parts[7]) / 100) * 60,
      parts[4]
    );
    return { lat: lat, lng: lng };
  }

  // The following will convert your DMS to DD
  ConvertDMSToDD(
    degrees: number,
    minutes: number,
    seconds: number,
    direction: string
  ) {
    var dd =
      Number(degrees) + Number(minutes) / 60 + Number(seconds) / (60 * 60);
    if (direction == 'S' || direction == 'W') {
      dd = dd * -1;
    } // Don't do anything for N or E
    //convert all to 10 decimal places
    return dd.toFixed(10);
  }

  //export data   to excel
  exportExcel() {
    const workSheet = XLSX.utils.json_to_sheet(this.excelData, {
      // header: ['Description', 'Name', 'DMS Lat', 'DMS Lng', 'DD Lat', 'DD Lng'],
    });
    const workBook: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workBook, workSheet, 'coordinates');
    XLSX.writeFile(workBook, 'updated coordinates t - combined.xlsx');
  }
}
