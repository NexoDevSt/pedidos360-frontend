import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { HttpClient, HttpHeaders } from '@angular/common/http';
import { MsalService, MsalBroadcastService } from '@azure/msal-angular';
import { InteractionStatus } from '@azure/msal-browser';
import { filter } from 'rxjs/operators';
import { FormsModule } from '@angular/forms';
import { apiConfig } from './auth-config';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, FormsModule],
  template: `
    <div style="padding: 30px; font-family: Arial, sans-serif; max-width: 950px; margin: auto;">
      <h2>Sistema Pedidos360 - Duoc UC</h2>

      <div *ngIf="!isLoggedIn" style="margin-top: 20px;">
        <p>Inicie sesión con su cuenta institucional de Azure AD para continuar.</p>
        <button (click)="login()" style="padding: 10px 18px; font-size: 14px; cursor: pointer; background-color: #0078d4; color: white; border: none; border-radius: 4px;">
          Iniciar Sesión con Azure AD
        </button>
      </div>

      <div *ngIf="isLoggedIn" style="margin-top: 20px;">
        <p>Usuario conectado: <strong>{{ userName }}</strong></p>
        <button (click)="logout()" style="padding: 8px 14px; margin-right: 10px; cursor: pointer;">Cerrar Sesión</button>
        <button (click)="consultarPedidos()" style="padding: 8px 14px; margin-right: 10px; cursor: pointer; background-color: #107c41; color: white; border: none; border-radius: 4px;">
          Consultar Órdenes (GET)
        </button>
        <button (click)="consultarNotificaciones()" style="padding: 8px 14px; margin-right: 10px; cursor: pointer; background-color: #d83b01; color: white; border: none; border-radius: 4px;">
          Consultar Notificaciones (GET)
        </button>

        <hr style="margin: 25px 0;">

        <!-- Búsqueda específica por ID de Orden -->
        <h3>Consultar Orden por ID (GET /api/pedidos/&#123;id&#125;):</h3>
        <div style="margin-bottom: 15px;">
          <input [(ngModel)]="searchId" placeholder="Ej: OT-2026-000001" style="padding: 8px; width: 220px; margin-right: 10px;" />
          <button (click)="consultarPedidoPorId()" style="padding: 8px 14px; cursor: pointer; background-color: #0078d4; color: white; border: none; border-radius: 4px;">
            Buscar Orden
          </button>
        </div>

        <div *ngIf="pedidoDetalle" style="background-color: #f9f9f9; padding: 15px; border: 1px solid #ddd; margin-bottom: 20px; border-radius: 4px;">
          <h4>Detalle de Orden:</h4>
          <p><strong>ID OT:</strong> {{ pedidoDetalle.otId }}</p>
          <p><strong>Cliente:</strong> {{ pedidoDetalle.clienteId }}</p>
          <p><strong>Patente:</strong> {{ pedidoDetalle.patente }}</p>
          <p><strong>Descripción:</strong> {{ pedidoDetalle.descripcion }}</p>
          <p><strong>Total:</strong> \${{ pedidoDetalle.total | number }}</p>
        </div>

        <!-- FORMULARIO POST: EMISIÓN DE NOTIFICACIÓN -->
        <div style="background-color: #fff4ce; padding: 15px; border: 1px solid #f2c80f; margin-bottom: 25px; border-radius: 4px;">
          <h4 style="margin-top: 0; color: #795b00;">Crear Nueva Notificación (POST /api/notificaciones):</h4>
          <div style="display: flex; gap: 10px; flex-wrap: wrap; align-items: center;">
            <input [(ngModel)]="nuevaNotif.otId" placeholder="OT ID (Ej: OT-2026-000001)" style="padding: 6px; width: 170px;" />
            <input [(ngModel)]="nuevaNotif.clienteId" placeholder="Cliente ID (Ej: CLI-001)" style="padding: 6px; width: 150px;" />
            <select [(ngModel)]="nuevaNotif.canal" style="padding: 6px;">
              <option value="email">Email</option>
              <option value="sms">SMS</option>
              <option value="push">Push</option>
            </select>
            <input [(ngModel)]="nuevaNotif.payloadJson" placeholder="Mensaje / Payload JSON" style="padding: 6px; width: 250px;" />
            <button (click)="crearNotificacion()" style="padding: 7px 15px; cursor: pointer; background-color: #795b00; color: white; border: none; border-radius: 4px;">
              Enviar Notificación
            </button>
          </div>
          <p *ngIf="notifCreadaMsg" style="color: green; margin-bottom: 0; font-weight: bold;">{{ notifCreadaMsg }}</p>
        </div>

        <!-- TABLA: Notificaciones -->
        <div *ngIf="notificaciones.length > 0">
          <h3 style="margin-top: 20px; color: #d83b01;">Registro de Notificaciones (Microservicio de Eventos/Mensajería):</h3>
          <table border="1" cellpadding="10" style="border-collapse: collapse; width: 100%;">
            <thead style="background-color: #fcebe5;">
              <tr>
                <th>Log ID</th>
                <th>OT Asociada</th>
                <th>Cliente ID</th>
                <th>Canal</th>
                <th>Payload JSON</th>
              </tr>
            </thead>
            <tbody>
              <tr *ngFor="let notif of notificaciones">
                <td>{{ notif.logId }}</td>
                <td>{{ notif.otId }}</td>
                <td>{{ notif.clienteId }}</td>
                <td><span style="text-transform: uppercase; font-weight: bold;">{{ notif.canal }}</span></td>
                <td><code style="font-size: 12px; background: #eee; padding: 2px 4px;">{{ notif.payloadJson }}</code></td>
              </tr>
            </tbody>
          </table>
        </div>

        <!-- TABLA: Órdenes de Trabajo -->
        <div *ngIf="pedidos.length > 0">
          <h3 style="margin-top: 20px;">Listado de Órdenes de Trabajo (Oracle Cloud DB):</h3>
          <table border="1" cellpadding="10" style="border-collapse: collapse; width: 100%;">
            <thead style="background-color: #f2f2f2;">
              <tr>
                <th>ID OT</th>
                <th>Cliente</th>
                <th>Patente</th>
                <th>Descripción</th>
                <th>Total (CLP)</th>
              </tr>
            </thead>
            <tbody>
              <tr *ngFor="let ot of pedidos">
                <td>{{ ot.otId }}</td>
                <td>{{ ot.clienteId }}</td>
                <td>{{ ot.patente }}</td>
                <td>{{ ot.descripcion }}</td>
                <td>\${{ ot.total | number }}</td>
              </tr>
            </tbody>
          </table>
        </div>

        <p *ngIf="cargando" style="color: #0078d4; font-weight: bold;">Procesando solicitud con API Gateway...</p>
        <p *ngIf="errorMsg" style="color: red;">{{ errorMsg }}</p>
      </div>
    </div>
  `
})
export class AppComponent implements OnInit {
  isLoggedIn = false;
  userName = '';
  pedidos: any[] = [];
  notificaciones: any[] = [];
  pedidoDetalle: any = null;
  searchId: string = 'OT-2026-000001';
  cargando = false;
  errorMsg = '';
  notifCreadaMsg = '';

  nuevaNotif = {
    otId: 'OT-2026-000001',
    clienteId: 'CLI-001',
    canal: 'email',
    payloadJson: '{"mensaje":"Vehiculo listo para entrega"}'
  };

  constructor(
    private authService: MsalService,
    private msalBroadcastService: MsalBroadcastService,
    private http: HttpClient
  ) {}

  ngOnInit(): void {
    this.authService.instance.initialize().then(() => {
      this.authService.handleRedirectObservable().subscribe();
    });

    this.msalBroadcastService.inProgress$
      .pipe(filter((status: InteractionStatus) => status === InteractionStatus.None))
      .subscribe(() => {
        const accounts = this.authService.instance.getAllAccounts();
        this.isLoggedIn = accounts.length > 0;
        if (this.isLoggedIn) {
          this.userName = accounts[0].name || accounts[0].username;
        }
      });
  }

  login(): void {
    this.authService.loginRedirect();
  }

  logout(): void {
    this.authService.logoutRedirect();
  }

  private obtenerToken(callback: (token: string) => void): void {
    const accounts = this.authService.instance.getAllAccounts();
    if (accounts.length === 0) {
      this.errorMsg = 'No hay una sesión activa de Azure AD.';
      this.cargando = false;
      return;
    }

    this.authService.instance.acquireTokenSilent({
      scopes: apiConfig.scopes,
      account: accounts[0]
    }).then(response => {
      callback(response.accessToken);
    }).catch(error => {
      this.authService.instance.acquireTokenPopup({
        scopes: apiConfig.scopes
      }).then(response => {
        callback(response.accessToken);
      }).catch(err => {
        this.errorMsg = 'Error al adquirir el token de acceso.';
        this.cargando = false;
      });
    });
  }

  consultarPedidos(): void {
    this.cargando = true;
    this.errorMsg = '';
    this.notifCreadaMsg = '';
    this.pedidoDetalle = null;
    this.notificaciones = [];

    this.obtenerToken((token) => {
      const headers = new HttpHeaders({ 'Authorization': `Bearer ${token}` });
      this.http.get<any[]>(apiConfig.uri, { headers }).subscribe({
        next: (data) => {
          this.pedidos = data;
          this.cargando = false;
        },
        error: (err) => {
          this.errorMsg = `Error ${err.status}: ${err.statusText || 'Error en API Gateway'}`;
          this.cargando = false;
        }
      });
    });
  }

  consultarPedidoPorId(): void {
    if (!this.searchId.trim()) return;
    this.cargando = true;
    this.errorMsg = '';
    this.notifCreadaMsg = '';
    this.notificaciones = [];
    this.pedidos = [];

    this.obtenerToken((token) => {
      const headers = new HttpHeaders({ 'Authorization': `Bearer ${token}` });
      const url = `${apiConfig.uri}/${this.searchId.trim()}`;
      this.http.get<any>(url, { headers }).subscribe({
        next: (data) => {
          this.pedidoDetalle = data;
          this.cargando = false;
        },
        error: (err) => {
          this.errorMsg = `Error ${err.status}: No se encontró la orden indicada`;
          this.cargando = false;
        }
      });
    });
  }

  consultarNotificaciones(): void {
    this.cargando = true;
    this.errorMsg = '';
    this.notifCreadaMsg = '';
    this.pedidoDetalle = null;
    this.pedidos = [];

    this.obtenerToken((token) => {
      const headers = new HttpHeaders({ 'Authorization': `Bearer ${token}` });
      const urlNotificaciones = apiConfig.uri.replace('/pedidos', '/notificaciones');

      this.http.get<any[]>(urlNotificaciones, { headers }).subscribe({
        next: (data) => {
          this.notificaciones = data;
          this.cargando = false;
        },
        error: (err) => {
          this.errorMsg = `Error ${err.status}: ${err.statusText || 'Error al obtener notificaciones'}`;
          this.cargando = false;
        }
      });
    });
  }

  crearNotificacion(): void {
    this.cargando = true;
    this.errorMsg = '';
    this.notifCreadaMsg = '';

    this.obtenerToken((token) => {
      const headers = new HttpHeaders({
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      });
      const urlNotificaciones = apiConfig.uri.replace('/pedidos', '/notificaciones');

      this.http.post<any>(urlNotificaciones, this.nuevaNotif, { headers }).subscribe({
        next: (creada) => {
          this.cargando = false;
          this.notifCreadaMsg = `¡Notificación #${creada.logId || ''} registrada en Oracle exitosamente!`;
          this.consultarNotificaciones();
        },
        error: (err) => {
          this.cargando = false;
          this.errorMsg = `Error al registrar notificación (${err.status}): ${err.statusText || 'Fallo de inserción'}`;
        }
      });
    });
  }
}