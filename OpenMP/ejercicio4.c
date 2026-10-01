#include <stdio.h>
#include <stdlib.h>
#include <omp.h>
int main() {
long num_lanzamientos = 100000000; // 100M de dardos
long puntos_dentro = 0;
double inicio, fin;
inicio = omp_get_wtime();
// Completen con parallel for y la cláusula de reducción para 'puntos_dentro'
#pragma omp parallel for reduction(+:puntos_dentro)
for (long i = 0; i < num_lanzamientos; i++) {
// Creamos una semilla unica para cada hilo en cada iteración
unsigned int seed = i + omp_get_thread_num();
// Generamos coordenadas x e y entre -1.0 y 1.0 usando rand_r
double x = (double)rand_r(&seed) / RAND_MAX * 2.0 - 1.0;
double y = (double)rand_r(&seed) / RAND_MAX * 2.0 - 1.0;
// Ecuación del círculo: x^2 + y^2 <= r^2
if (x * x + y * y <= 1.0) {
puntos_dentro++;
}
}
double pi_estimado = 4.0 * ((double)puntos_dentro / (double)num_lanzamientos);
fin = omp_get_wtime();
printf("Valor de Pi estimado: %.6f\n", pi_estimado);
printf("Tiempo paralelo: %f segundos\n", fin - inicio);
// Para el tiempo serial, comenten el pragma, compilen y midan de nuevo el
timempo.
return 0;
}
