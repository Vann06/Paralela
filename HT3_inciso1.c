/*----------------------------------------------------------------------
 * Universidad del Valle de Guatemala
 * Curso:     CC3169 - Computacion Paralela y Distribuida
 * Ejercicio: Hoja de Trabajo 02 - Introduccion a Open MPI 
 *			  Inciso 1
 * Descripcion: simulación de una empresa formada por una 
 *              oficina central y varias sucursales.
 * 				Cada proceso MPI representa una ubicacion diferente:
 *     			 rank 0  -> Oficina central
 *      		 rank 1  -> Sucursal 1
 *      		 rank 2  -> Sucursal 2
 *      		 ...
 *
 * 				Identificar el rank de cada proceso y el numero 
 *              total de procesos que participan en la ejecucion.
 *----------------------------------------------------------------------*/

#include <stdio.h>
#include <mpi.h>

int main(int argc, char *argv[]) {

    int rank;
    int size;

    // Inicializa el entorno MPI: debe ejecutarse antes de utilizar otras funciones MPI
    MPI_Init(&argc, &argv);

	// Obtener el identificador del proceso actual
    MPI_Comm_rank(MPI_COMM_WORLD, &rank);

    // Obtener el numero total de procesos que participan
    MPI_Comm_size(MPI_COMM_WORLD, &size);

    //Cada proceso identifica la ubicacion que representa
    if (rank == 0) {
        printf("Proceso %d: Oficina Central\n", rank);

    } else {
        printf("Proceso %d: Sucursal %d\n", rank, rank);
    }

    // Todos los procesos conocen el numero total de procesos
    printf("Proceso %d: Existen %d procesos en la simulacion.\n", rank, size);

    // Finaliza correctamente el entorno MPI
    MPI_Finalize();

    return 0;
}