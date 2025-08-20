import { CommonModule } from '@angular/common';
import { Component, OnDestroy, OnInit } from '@angular/core';
import { Router, RouterModule } from '@angular/router';
import { Subscription } from 'rxjs';
import { UserStateService } from '../user-state.service';

@Component({
  selector: 'app-menu',
  imports: [CommonModule, RouterModule],
  templateUrl: './menu.component.html',
  styleUrl: './menu.component.css',
})
export class MenuComponent implements OnInit, OnDestroy {
private subscription: Subscription = new Subscription();

  menus: any[] = []; // lo definimos vacío
rolUsuario: string | null = null; 
  private menusCompletos = [
    {
      title: 'Dashboard',
      icon: 'das.svg',
      subItems: [
        { title: 'Perforación Taladros Largos', path: 'taladro-largo-grafico' },
        { title: 'Perforación Horizontal', path: 'taladro-horizontal-grafico' },
        { title: 'Sostenimiento', path: 'sostenimiento' },
        { title: 'Explosivos', path: 'explosivos-graficos' },
        { title: 'Reporte Indicadores', path: 'power-bi' },
      ],
    },
    {
      title: 'Carga de Datos',
      icon: 'data.svg',
      subItems: [
        { title: 'Explosivos', path: 'explosivos' },
        { title: 'Estados', path: 'estados' },
        { title: 'Crear Data', path: 'crear-data' },
        { title: 'Plan de Avance', path: 'plan-avance' },
        { title: 'Plan de Metraje', path: 'plan-metraje' },
        { title: 'Plan de Producción', path: 'plan-produccion' },
        { title: 'Metas', path: 'metas' },
        { title: 'Checklist', path: 'checklist' },
        { title: 'Semanas', path: 'semana-personali' },
        { title: 'PDF', path: 'pdf' },
      ],
    },
    {
      title: 'Roles',
      icon: 'usuario.png',
      subItems: [
        { title: 'Usuarios', path: 'usuarios' },
        { title: 'Perfil', path: 'perfil' },
      ],
    },
  ];

  private menuSoloDashboard = [
    {
      title: 'Dashboard',
      icon: 'das.svg',
      subItems: [
        { title: 'Perforación Taladros Largos', path: 'taladro-largo-grafico' },
        { title: 'Perforación Horizontal', path: 'taladro-horizontal-grafico' },
        { title: 'Sostenimiento', path: 'sostenimiento' },
        { title: 'Explosivos', path: 'explosivos-graficos' },
        { title: 'Reporte Indicadores', path: 'power-bi' },
      ],
    },
  ];

  menuOpenIndex: number | null = null;
  selectedSubItemIndex: number | null = null;
  selectedSubItem: string | null = null;

  constructor(private router: Router, private userStateService: UserStateService ) {
    if (this.router.url === '/' || this.router.url === '/Dashboard') {
      this.router.navigate(['/Dashboard/taladro-largo-grafico']);
    }
  }

 ngOnInit(): void {
    // Suscribirse a cambios en el rol
    this.subscription.add(
      this.userStateService.rolUsuario$.subscribe(rol => {
        this.actualizarMenu(rol);
      })
    );

    // Cargar valor inicial
    const rolInicial = this.userStateService.getRolUsuario();
    this.actualizarMenu(rolInicial);
  }

  private actualizarMenu(rol: string): void {
    this.rolUsuario = rol;
    console.log('Rol del usuario en menu:', this.rolUsuario);
    
    if (this.rolUsuario === 'Visualizador') {
      this.menus = this.menuSoloDashboard;
    } else {
      this.menus = this.menusCompletos;
    }
  }

  ngOnDestroy(): void {
    this.subscription.unsubscribe();
  }


  AbrirCerrar(index: number, menu: any) {
    if (menu.title === 'Home') {
      this.router.navigate(['/Dashboard/taladro-largo-grafico']);
    } else if (this.menuColapsado) {
      if (menu.subItems && menu.subItems.length > 0) {
        const ruta = `/Dashboard/${menu.subItems[0].path}`;
        this.router.navigate([ruta]);
        this.selectedSubItemIndex = 0;
        this.selectedSubItem = menu.subItems[0].path;
      }
    } else {
      this.menuOpenIndex = this.menuOpenIndex === index ? null : index;
    }
  }

  selectSubItem(index: number, subItem: any) {
    this.selectedSubItemIndex = index;
    this.selectedSubItem = subItem.path;

    const ruta = `/Dashboard/${subItem.path}`;
    this.router.navigate([ruta]);
  }

  convertirRuta(subItem: string): string {
    return subItem.toLowerCase().replace(/ /g, '-');
  }

  mostrarCerrarSesion = false;
  menuColapsado = false;

  toggleMenu() {
    this.menuColapsado = !this.menuColapsado;
  }
  toggleCerrarSesion() {
    this.mostrarCerrarSesion = !this.mostrarCerrarSesion;
  }

  cerrarSesion() {
  localStorage.removeItem('rolUsuario'); // 👈 limpiamos rol
  localStorage.clear();
  this.router.navigate(['/login']);
}
}
