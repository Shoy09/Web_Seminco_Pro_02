import { Component } from '@angular/core';
import { DomSanitizer, SafeResourceUrl } from '@angular/platform-browser';

@Component({
  selector: 'app-power-bi-public',
  imports: [],
  templateUrl: './power-bi-public.component.html',
  styleUrl: './power-bi-public.component.css'
})
export class PowerBiPublicComponent {
  reportUrl: SafeResourceUrl;

  constructor(private sanitizer: DomSanitizer) {
    // 👉 Pega aquí tu link público de Power BI
    const url = 'https://app.powerbi.com/view?r=eyJrIjoiNDFkYTZkYjgtYTIxMS00ODJmLWFiMzQtOGY0ZmYzNWFlZWM3IiwidCI6IjY4NmQ2YWVkLWU4YmQtNDFhNS1iZTdkLTRmNzYxN2UxYzE5MSIsImMiOjR9';
    this.reportUrl = this.sanitizer.bypassSecurityTrustResourceUrl(url);
  }
}