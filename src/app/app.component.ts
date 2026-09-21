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
    <div style="padding: 30px; font-family: Arial, sans-serif; max-width: 900px; margin: auto;">
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
        <button (click)="consultarPedidos()" style="padding: 8px 14px; cursor: pointer; background-color: #107c41; color: white; border: none; border-radius: 4px;">
          Consultar Todas las Órdenes
        </button>

        <hr style="margin: 25px 0;">

        <!-- Búsqueda específica por ID de Orden -->
        <h3>Consultar Orden por ID (Ruta /api/pedidos/&#123;id&#125;):</h3>
        <div style="margin-bottom: 15px;">
          <input [(ngModel)]="searchId" placeholder="Ej: OT-2026-000001" style="padding: 8px; width: 220px; margin-right: 10px;" />
          <button (click)="consultarPedidoPorId()" style="padding: 8px 14px; cursor: pointer; background-color: #0078d4; color: white; border: none; border-radius: 4px;">
            Buscar Orden
          </button>
        </div>

        <div *ngIf="pedidoDetalle" style="background-color: #f9f9f9; padding: 15px; border: 1px solid #ddd; margin-bottom: 20px; border-radius: 4px;">
          <h4>Detalle de Orden Encontrada:</h4>
          <p><strong>ID OT:</strong> {{ pedidoDetalle.otId }}</p>
          <p><strong>Cliente:</strong> {{ pedidoDetalle.clienteId }}</p>
          <p><strong>Patente:</strong> {{ pedidoDetalle.patente }}</p>
          <p><strong>Descripción:</strong> {{ pedidoDetalle.descripcion }}</p>
          <p><strong>Total:</strong> \${{ pedidoDetalle.total | number }}</p>
        </div>

        <h3 style="margin-top: 25px;">Listado General de Órdenes (Desde Oracle Cloud):</h3>
        <table border="1" cellpadding="10" style="border-collapse: collapse; width: 100%;" *ngIf="pedidos.length > 0">
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
  pedidoDetalle: any = null;
  searchId: string = 'OT-2026-000001';
  cargando = false;
  errorMsg = '';

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
    this.pedidoDetalle = null;

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
}