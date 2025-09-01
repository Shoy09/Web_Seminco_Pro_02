import { Injectable } from '@angular/core';
import { ApiService } from './api.service';
import { BehaviorSubject, Observable, tap } from 'rxjs';
import { OrigenDestino } from '../models/origen-destino.model';

@Injectable({
  providedIn: 'root'
})
export class OrigenDestinoService {
  private baseUrl = 'origen-destino'; // 👈 ruta del backend (coincide con Express)
  private origenDestinoActualizados = new BehaviorSubject<boolean>(false);

  constructor(private apiService: ApiService) {}

  // ✅ Obtener todos
  getOrigenesDestinos(): Observable<OrigenDestino[]> {
    return this.apiService.getDatos(`${this.baseUrl}/`);
  }

  // ✅ Obtener por ID
  getOrigenDestinoById(id: number): Observable<OrigenDestino> {
    return this.apiService.getDatos(`${this.baseUrl}/${id}`);
  }

  // ✅ Crear
  createOrigenDestino(data: OrigenDestino): Observable<OrigenDestino> {
    return this.apiService.postDatos(`${this.baseUrl}/`, data).pipe(
      tap(() => {
        this.origenDestinoActualizados.next(true); // Notificar actualización
      })
    );
  }

  // ✅ Actualizar
  updateOrigenDestino(id: number, data: OrigenDestino): Observable<OrigenDestino> {
    return this.apiService.putDatos(`${this.baseUrl}/${id}`, data).pipe(
      tap(() => {
        this.origenDestinoActualizados.next(true); // Notificar actualización
      })
    );
  }

  // ✅ Eliminar
  deleteOrigenDestino(id: number): Observable<any> {
    return this.apiService.deleteDatos(`${this.baseUrl}/${id}`).pipe(
      tap(() => {
        this.origenDestinoActualizados.next(true); // Notificar actualización
      })
    );
  }

  // ✅ Observable para escuchar actualizaciones
  getOrigenDestinoActualizados(): Observable<boolean> {
    return this.origenDestinoActualizados.asObservable();
  }
}
