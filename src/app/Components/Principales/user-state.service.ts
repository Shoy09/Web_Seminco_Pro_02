// user-state.service.ts
import { Injectable } from '@angular/core';
import { BehaviorSubject } from 'rxjs';

@Injectable({
  providedIn: 'root'
})
export class UserStateService {
  private rolUsuarioSource = new BehaviorSubject<string>('');
  rolUsuario$ = this.rolUsuarioSource.asObservable();

  setRolUsuario(rol: string) {
    this.rolUsuarioSource.next(rol);
    localStorage.setItem('rolUsuario', rol);
  }

  getRolUsuario(): string {
    return localStorage.getItem('rolUsuario') || '';
  }
}