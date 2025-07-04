import { Injectable } from '@angular/core';
import { Observable } from 'rxjs';
import { ApiService } from './api.service';

@Injectable({
  providedIn: 'root'
})
export class NubeDatosTrabajoExploracionesService  {
  private readonly endpoints = {
    explosivos: 'NubeDatosExploraciones',

  };

  constructor(private apiService: ApiService) {}

  getExplosivos(): Observable<any> {
    return this.apiService.getDatos(this.endpoints.explosivos);
  }
}