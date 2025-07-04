import { Routes } from '@angular/router';
import { LoginComponent } from './Components/Principales/login/login.component';
import { PrincipalComponent } from './Components/Principales/principal/principal.component';
import { HomeComponent } from './Components/Principales/home/home.component';
import { EstadosComponent } from './Components/Estado/estados/estados.component';
import { UsuariosComponent } from './Components/Usuario/usuarios/usuarios.component';
import { CrearDataComponent } from './Components/Crear datos/crear-data/crear-data.component';
import { PlanMensualListComponent } from './Components/Planes mensuales/Plan mensual avance/plan-mensual-list/plan-mensual-list.component';
import { PlanMetrajeListComponent } from './Components/Planes mensuales/Plan mensual metraje/plan-metraje-list/plan-metraje-list.component';
import { PlanProduccionListComponent } from './Components/Planes mensuales/Plan mensual produccion/plan-produccion-list/plan-produccion-list.component';
import { ExplosivosComponent } from './Components/Crear datos/explosivos/explosivos.component';
import { UsuarioComponent } from './Components/Usuario/usuario/usuario.component';
import { MostrarGraficosComponent } from './Components/Dato Movil/mostrar-graficos/mostrar-graficos.component';
import { TaladroHorizontalGraficaComponent } from './Components/Dashboard/Horizontal/taladro-horizontal-grafica/taladro-horizontal-grafica.component';
import { TaladroLargoGraficaComponent } from './Components/Dashboard/Largo/taladro-largo-grafica/taladro-largo-grafica.component';
import { SostenimientoGraficosComponent } from './Components/Dashboard/Sostenimiento/sostenimiento-graficos/sostenimiento-graficos.component';
import { MetasComponent } from './Components/Crear datos/Metas/metas/metas.component';
import { ExplosivosGraficosComponent } from './Components/Dashboard/Explosivos/explosivos-graficos/explosivos-graficos.component';

export const routes: Routes = [
  { path: '', redirectTo: '/login', pathMatch: 'full' },
  { path: 'login', component: LoginComponent },

  {
    path: 'Dashboard',
    component: PrincipalComponent, // Layout principal con menú
    children: [
      { path: 'Home', component: HomeComponent },
      { path: 'estados', component: EstadosComponent },
      { path: 'crear-data', component: CrearDataComponent },
      { path: 'plan-avance', component: PlanMensualListComponent },
      { path: 'plan-metraje', component: PlanMetrajeListComponent },
      { path: 'plan-produccion', component: PlanProduccionListComponent },
      { path: 'explosivos', component: ExplosivosComponent },
      { path: 'usuarios', component: UsuariosComponent },
      { path: 'perfil', component: UsuarioComponent },
      { path: 'graficos', component: MostrarGraficosComponent },
      { path: 'taladro-largo-grafico', component: TaladroLargoGraficaComponent },
      { path: 'taladro-horizontal-grafico', component: TaladroHorizontalGraficaComponent },
      { path: 'sostenimiento', component: SostenimientoGraficosComponent },
      { path: 'metas', component: MetasComponent },
      { path: 'explosivos-graficos', component: ExplosivosGraficosComponent },
    ],
  },

  { path: '**', redirectTo: '/login' },
];
 