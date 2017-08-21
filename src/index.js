import Handsontable from '../hot-pro-trimmed/handsontable.full';
import '../hot-pro-trimmed/handsontable.full.css';
import XlsxPopulate from 'xlsx-populate';

class GanttExport {
  constructor(element) {
    this.hot = null;

    this.createGantt(element);
  }

  createGantt(element) {
    const dataObj = [
      {
        startDate: '1/5/2017',
        endDate: '1/20/2017',
        additionalData: {label: 'Label 1'}
      },
      {
        startDate: '1/11/2017',
        endDate: '1/29/2017',
        additionalData: {label: 'Label 2'}
      },
      {
        startDate: '2/1/2017',
        endDate: '2/26/2017',
        additionalData: {label: 'Label 3'}
      },
      {
        startDate: '2/15/2017',
        endDate: '3/26/2017',
        additionalData: {label: 'Label 4'}
      }
    ];

    this.hot = new Handsontable(element, {
      data: [],
      hiddenColumns: true,
      colHeaders: true,
      ganttChart: {
        dataSource: dataObj,
        firstWeekDay: 'monday',
        startYear: 2017
      },
      width: 600,
      height: 165
    });
  }

}

class Exporter {
  constructor(ganttHot) {
    this.hot = ganttHot;
  }

  generateXLS() {
    XlsxPopulate.fromBlankAsync()
      .then(workbook => {
        // Modify the workbook.
        // workbook.sheet("Sheet1").cell("A1").value("This is neat!");

        this.fillSheet(workbook);

        // Write to file.
        return workbook.outputAsync();
      }).then(function(blob) {
      if (window.navigator && window.navigator.msSaveOrOpenBlob) {
        window.navigator.msSaveOrOpenBlob(blob, "out.xlsx");
      } else {
        let url = window.URL.createObjectURL(blob);
        let a = document.createElement("a");
        document.body.appendChild(a);
        a.href = url;
        a.download = "out.xlsx";
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
      }
    })
      .catch(function(err) {
        alert(err.message || err);
        throw err;
      });
  }

  fillSheet(workbook) {
    const range = workbook.sheet(0).range("A2:B5");
    const tableData = [['Sample data 1', 'Sample 1'], ['Sample data 2', 'Sample 2'], ['Sample data 3', 'Sample 3'], ['Sample data 4', 'Sample 4']];
    range.value(tableData);

    for (let r = 0; r < this.hot.countRows(); r ++) {
      let metas = this.hot.getCellMetaAtRow(r);

      for (let c = 0; c < 20; c++) {
        let color = 'ffffff';
        if (metas[c] && metas[c].className && metas[c].className.indexOf('rangeBar') > -1) {
          if(metas[c].className.indexOf('partial') > -1) {
            color = '8edf5a';
          } else {
            color = '48b703';
          }
        }

        workbook.sheet(0).cell(r + 2, 4 + c).value('').style('fill', color);
      }
    }

    workbook.sheet(0).range(1, 1, 1, 2).merged(true);
    workbook.sheet(0).range(1, 4, 1, 23).merged(true);
    workbook.sheet(0).cell(1, 1).value('Sample dataset:').style('bold', true);
    workbook.sheet(0).cell(1, 4).value('Sample Gantt Data:').style('bold', true);
  }
}

window.Exporter = Exporter;
window.GanttExport = GanttExport;