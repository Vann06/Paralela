#include <stdio.h>
#include <omp.h>

int es_primo(int n) {
    if (n < 2)
        return 0;

    for (int i = 2; i < n; i++) {
        if (n % i == 0)
            return 0;
    }

    return 1;
}

int main() {

    int limite = 500000;
    long cantidad_primos = 0;

    double inicio = omp_get_wtime();

    #pragma omp parallel for schedule(dynamic) reduction(+:cantidad_primos)
    for (int i = 2; i <= limite; i++) {
        if (es_primo(i)) {
            cantidad_primos++;
        }
    }

    double fin = omp_get_wtime();

    printf("Primos encontrados: %ld\n", cantidad_primos);
    printf("Tiempo dynamic: %f segundos\n", fin - inicio);

    return 0;
}
