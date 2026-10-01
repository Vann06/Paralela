#include <stdio.h>
#include <stdlib.h>
#include <omp.h>
#define N 100000000 // probar
int main() {
double *arreglo = (double*)malloc(N * sizeof(double));
for(int i = 0; i < N; i++) arreglo[i] = (rand() % 10000) / 100.0;
double suma = 0.0, maximo = -1.0, minimo = 10000.0;
double inicio = omp_get_wtime();
// Iniciar la regin de secciones
//#pragma omp parallel sections
{
// Seccion 1: Suma (Media)
//#pragma omp section
{
double suma_local = 0;
for(int i = 0; i < N; i++) suma_local += arreglo[i];
suma = suma_local;
}
// Seccion 2: Máximo
//#pragma omp section
{
double max_local = -1.0;
for(int i = 0; i < N; i++) {
if(arreglo[i] > max_local) max_local = arreglo[i];
}
maximo = max_local;
}
// Seccion 3: Mínimo
//#pragma omp section
{
double min_local = 10000.0;
for(int i = 0; i < N; i++) {
if(arreglo[i] < min_local) min_local = arreglo[i];
}
minimo = min_local;
}
}
double fin = omp_get_wtime();
printf("Media: %.2f, Max: %.2f, Min: %.2f\n", suma/N, maximo, minimo);
printf("Tiempo serial: %f segundos\n", fin - inicio);
free(arreglo);
return 0;
}
