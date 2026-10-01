/*----------------------------------------------------------------------
 * Universidad del Valle de Guatemala
 * Curso:     CC3169 - Computacion Paralela y Distribuida
 * Ejercicio: Hoja de Trabajo 02 - Introduccion a Open MPI
 *            Inciso 2
 * Descripcion: simulacion del envio del reporte diario de ventas
 *              desde una sucursal hacia la oficina central.
 *
 *              Cada proceso MPI representa una ubicacion diferente:
 *                  rank 0 -> Oficina central
 *                  rank 1 -> Sucursal 1
 *
 *              La Sucursal 1 envia el total de ventas del dia hacia
 *              la Oficina Central utilizando comunicacion punto a punto.
 *
 *              Utilizar MPI_Send() y MPI_Recv() para realizar
 *              comunicacion directa entre dos procesos.
 *----------------------------------------------------------------------*/

#include <stdio.h>
#include <mpi.h>

int main(int argc, char *argv[]) {

    int rank;
    int size;
    float ventas;

    // Inicializa el entorno MPI: debe ejecutarse antes de utilizar otras funciones MPI
    MPI_Init(&argc, &argv);

    // Obtener el identificador del proceso actual
    MPI_Comm_rank(MPI_COMM_WORLD, &rank);

    // Obtener el numero total de procesos que participan
    MPI_Comm_size(MPI_COMM_WORLD, &size);

    // Verificar que existan al menos dos procesos para realizar la comunicacion
    if (size < 2) {

        if (rank == 0) {
            printf("Este programa requiere al menos 2 procesos.\n");
        }

        MPI_Finalize();
        return 0;
    }

    // La Sucursal 1 genera y envia su reporte de ventas
    if (rank == 1) {

        ventas = 1250.75;

        printf("Sucursal 1: ventas del dia = Q%.2f\n", ventas);

        MPI_Send(&ventas, 1, MPI_FLOAT, 0, 100, MPI_COMM_WORLD);

        printf("Sucursal 1: reporte enviado a Oficina Central.\n");
    }

    // La Oficina Central recibe el reporte enviado por la Sucursal 1
    if (rank == 0) {

        MPI_Recv(&ventas, 1, MPI_FLOAT, 1, 100,MPI_COMM_WORLD, MPI_STATUS_IGNORE);

        printf("Oficina Central: reporte recibido.\n");
        printf("Ventas reportadas por Sucursal 1: Q%.2f\n", ventas);
    }

    // Finaliza correctamente el entorno MPI
    MPI_Finalize();

    return 0;
}