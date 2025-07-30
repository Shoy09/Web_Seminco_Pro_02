// src/app/models/pdf.model.ts
export interface Pdf {
  id?: number;           // Opcional para cuando creas uno nuevo
  proceso: string;
  mes: 'ENERO' | 'FEBRERO' | 'MARZO' | 'ABRIL' | 'MAYO' | 'JUNIO' |
       'JULIO' | 'AGOSTO' | 'SEPTIEMBRE' | 'OCTUBRE' | 'NOVIEMBRE' | 'DICIEMBRE';
  url_pdf: string;
  createdAt?: string;    // Opcionales si vienen del backend
  updatedAt?: string;
}
