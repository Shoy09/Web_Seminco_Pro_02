export interface SubEstado {
  id: number;
  codigo: string;
  tipo_estado: string;
  estadoId: number; // Relación con el estado principal
}

export interface Estado {
  id: number;
  estado_principal: string;
  codigo: string;
  tipo_estado: string;
  categoria: string;
  proceso: string;
  subEstados?: SubEstado[]; // Relación opcional: lista de subestados
}

export interface Estado2 {
  estado_principal: string;
  codigo: string;
  tipo_estado: string;
  categoria: string;
  proceso: string; // Nuevo campo agregado
}
