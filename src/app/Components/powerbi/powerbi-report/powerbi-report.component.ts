import { Component, Input } from '@angular/core';
import { SafeUrlPipe } from "../safe-url.pipe";

@Component({
  selector: 'app-powerbi-report',
  imports: [SafeUrlPipe],
  templateUrl: './powerbi-report.component.html',
  styleUrl: './powerbi-report.component.css'
})
export class PowerbiReportComponent {
 @Input() reportUrl: string = ''; // recibe el link público del Power BI
}