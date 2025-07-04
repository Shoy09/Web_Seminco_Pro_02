export interface TipoPerforacion {
    id?: number;  // `id` es opcional
    nombre: string;
    proceso?: string | null;  // `proceso` puede ser opcional o `null`
}
