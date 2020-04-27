import { AppService } from './app.service';
import { Component, OnInit, ViewChild, ElementRef } from '@angular/core';

declare const jexcel: any;

@Component({
    selector: 'app-root',
    templateUrl: './app.component.html',
    styleUrls: ['./app.component.scss']
})
export class AppComponent implements OnInit {
    title = 'jexcel-ng';

    constructor(
        private service: AppService
    ) {}
    @ViewChild('excelElement', { static: false }) private excelElement: ElementRef;
    ngOnInit() {
        const selectionActive = (instance, x1, y1, x2, y2, origin) => {
            const cellName1 = jexcel.getColumnNameFromId([x1, y1]);
            const cellName2 = jexcel.getColumnNameFromId([x2, y2]);
            console.log('The selection from ' + cellName1 + ' to ' + cellName2 + '');
        };
        jexcel.fromSpreadsheet('assets/公务员数量变化情况.xls', (result) => {
            if (!result.length) {
                console.error('JEXCEL: Something went wrong.');
            } else {
                if (result.length === 1) {
                    result[0].allowComments = true;
                    result[0].columnSorting = false;
                    result[0].contextMenu = () => {
                        window.event.returnValue = false;
                        return false;
                    };
                    result[0].onselection = selectionActive;
                    result[0].isSelectedMerge = false;
                    result[0].freezeColumns = 3;
                    result[0].freezeRows = 5;
                    result[0].tableOverflow = true;
                    result[0].tableWidth = '1000px';
                    result[0].tableHeight = '500px';
                    result[0].columns = result[0].columns.filter((cell, index) => index < result[0].data[0].length);
                    jexcel(this.excelElement.nativeElement, result[0]);
                } else {
                    jexcel.createTabs(this.excelElement.nativeElement, result);
                }
            }
        });
    }
}
