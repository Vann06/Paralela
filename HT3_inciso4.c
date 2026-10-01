/*----------------------------------------------------------------------
 * Universidad del Valle de Guatemala
 * Curso:     CC3169 - Computacion Paralela y Distribuida
 * Ejercicio: Hoja de Trabajo 02 - Introduccion a Open MPI
 *            Inciso 4
 * Descripcion: simulacion de la distribucion de pedidos desde la
 *              Oficina Central hacia las sucursales.
 *
 *              Cada proceso MPI representa una ubicacion diferente:
 *                  rank 0 -> Oficina central
 *                  rank 1 -> Sucursal 1
 *                  rank 2 -> Sucursal 2
 *                  rank 3 -> Sucursal 3
 *
 *              La Oficina Central posee una lista de pedidos y
 *              distribuye una parte a cada proceso utilizando
 *              MPI_Scatter().
 *----------------------------------------------------------------------*/

#include <stdio.h>
#include <mpi.h>

int main(int argc, char *argv[]) {

    int rank;
    int size;
    int pedidos[4];
    int pedido_recibido;

    // Inicializa el entorno MPI
    MPI_Init(&argc, &argv);

    // Obtener el identificador del proceso actual
    MPI_Comm_rank(MPI_COMM_WORLD, &rank);

    // Obtener el numero total de procesos que participan
    MPI_Comm_size(MPI_COMM_WORLD, &size);

    // Este ejercicio requiere exactamente 4 procesos
    if (size != 4) {

        if (rank == 0) {
            printf("Este programa requiere exactamente 4 procesos.\n");
        }

        MPI_Finalize();
        return 0;
    }

    // La Oficina Central define la cantidad de pedidos para cada ubicacion
    if (rank == 0) {

        pedidos[0] = 120;
        pedidos[1] = 95;
        pedidos[2] = 140;
        pedidos[3] = 110;

        printf("Oficina Central: distribuyendo pedidos...\n");
    }

    // Distribuir un valor del arreglo a cada proceso
    MPI_Scatter(
        pedidos,
        1,
        MPI_INT,
        &pedido_recibido,
        1,
        MPI_INT,
        0,
        MPI_COMM_WORLD
    );

    // Cada proceso muestra el valor que recibio
    if (rank == 0) {
        printf("Oficina Central: %d pedidos asignados.\n",
               pedido_recibido);
    } else {
        printf("Sucursal %d: %d pedidos asignados.\n",
               rank, pedido_recibido);
    }

    // Finaliza correctamente el entorno MPI
    MPI_Finalize();

    return 0;
}