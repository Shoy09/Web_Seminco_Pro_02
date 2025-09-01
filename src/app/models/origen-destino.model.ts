export interface OrigenDestino {
  id?: number;       // opcional porque lo genera MySQL
  operacion: string; // ejemplo: "Carguio"
  tipo: string;      // ejemplo: "Origen" o "Destino"
  nombre: string;    // ejemplo: "Ejemplo"
}
