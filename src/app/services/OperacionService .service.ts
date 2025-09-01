import { Injectable } from '@angular/core';
import { Observable } from 'rxjs';
import { ApiService } from './api.service';
import { NubePerforacionHorizontal, NubePerforacionTaladroLargo, NubeSostenimiento } from '../models/operaciones.models';

@Injectable({
  providedIn: 'root'
})
export class OperacionService {
  private readonly endpoints = {
    largo: 'operacion/largo',
    horizontal: 'operacion/horizontal',
    sostenimiento: 'operacion/sostenimiento',
    carguio: 'operacion/carguio'
  };

  constructor(private apiService: ApiService) {}

  // **Operaciones de Taladro Largo**
  getOperacionesLargo(): Observable<any> {
    return this.apiService.getDatos(this.endpoints.largo);
  }

  postOperacionesLargo(data: any): Observable<any> {
  return this.apiService.postDatos(this.endpoints.largo, data);
}

  // **Operaciones Horizontales**
  getOperacionesHorizontal(): Observable<any> {
    return this.apiService.getDatos(this.endpoints.horizontal);
  }

postOperacionesHorizontal(data: any): Observable<any> {
  return this.apiService.postDatos(this.endpoints.horizontal, data);
}

  // **Operaciones de Sostenimiento**
  getOperacionesSostenimiento(): Observable<any> {
    return this.apiService.getDatos(this.endpoints.sostenimiento);
  }

postOperacionesSostenimiento(data: any): Observable<any> {
  return this.apiService.postDatos(this.endpoints.sostenimiento, data);
}

  getOperacionesCargui(): Observable<any> {
    return this.apiService.getDatos(this.endpoints.carguio);
  }

  postOperacionesCargui(data: any): Observable<any> {
  return this.apiService.postDatos(this.endpoints.carguio, data);
}
}
