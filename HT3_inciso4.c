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

    // 4 procesos x 2 datos para cada proceso
    int datos[8];

    // Cada proceso recibirá 2 datos
    int datos_recibidos[2];

    MPI_Init(&argc, &argv);

    MPI_Comm_rank(MPI_COMM_WORLD, &rank);

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

        // Oficina Central
        datos[0] = 120;  // pedidos
        datos[1] = 8;    // empleados

        // Sucursal 1
        datos[2] = 95;
        datos[3] = 6;

        // Sucursal 2
        datos[4] = 140;
        datos[5] = 10;

        // Sucursal 3
        datos[6] = 110;
        datos[7] = 7;

        printf("Oficina Central: distribuyendo pedidos y empleados...\n");
    }

    // Distribuir dos valores a cada proceso
    MPI_Scatter(
        datos,
        2,
        MPI_INT,
        datos_recibidos,
        2,
        MPI_INT,
        0,
        MPI_COMM_WORLD
    );

    // Mostrar los datos recibidos
    if (rank == 0) {

        printf(
            "Oficina Central: %d pedidos, %d empleados disponibles.\n",
            datos_recibidos[0],
            datos_recibidos[1]
        );

    } else {

        printf(
            "Sucursal %d: %d pedidos, %d empleados disponibles.\n",
            rank,
            datos_recibidos[0],
            datos_recibidos[1]
        );
    }

    MPI_Finalize();

    return 0;
}