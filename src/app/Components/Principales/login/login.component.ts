import { Component } from '@angular/core';
import { Router } from '@angular/router';
import { FormsModule } from '@angular/forms';
import { CommonModule } from '@angular/common';
import { ToastrService } from 'ngx-toastr';
import { AuthService } from '../../../services/auth-service.service';
import { UsuarioService } from '../../../services/usuario.service';
import { UserStateService } from '../user-state.service';

@Component({
  selector: 'app-login',
  standalone: true, // Marca el componente como standalone
  imports: [FormsModule, CommonModule], // Importa los módulos necesarios
  templateUrl: './login.component.html',
  styleUrl: './login.component.css'
})
export class LoginComponent {
  showPassword: boolean = false;
  codigo_dni: string = ''; 
  password: string = ''; 
  errorMessage: string = ''; // Para mostrar mensajes de error

  constructor(
    private readonly router: Router,
    private authService: AuthService,
    private usuarioService: UsuarioService,
    private _toastr: ToastrService,
    private userStateService: UserStateService
  ) {}

  togglePassword() {
    this.showPassword = !this.showPassword;
  }

login() {
  if (!this.codigo_dni || !this.password) {
    this.errorMessage = 'Por favor, ingresa todos los campos.';
    this._toastr.warning(this.errorMessage, 'Advertencia');
    return;
  }

  this._toastr.info('Iniciando sesión...', 'Por favor espera');

  this.authService.login(this.codigo_dni, this.password).subscribe({
    next: (response) => {
      if (response.token) {
        this.authService.setToken(response.token);

        // 🚀 Todos entran al Dashboard
        this.router.navigate(['/Dashboard']);
        this._toastr.success('Sesión iniciada con éxito', 'Bienvenido');

        // 👇 Si quieres guardar el rol aunque no decida la ruta
        this.usuarioService.obtenerRol().subscribe({
          next: (rolResponse) => {
            console.log('Rol del usuario:', rolResponse.rol);
            this.userStateService.setRolUsuario(rolResponse.rol); // ← Usar el servicio
          },
          error: (err) => {
            console.error('Error al obtener rol', err);
          }
        });

      } else {
        this.errorMessage = 'Error en la autenticación. Token no recibido.';
        this._toastr.error(this.errorMessage, 'Error');
      }
    },
    error: (error) => {
      console.error('Error en el login', error);
      this.errorMessage = 'Credenciales incorrectas o problema con el servidor.';
      this._toastr.error(this.errorMessage, 'Error de autenticación');
    }
  });
}

}
