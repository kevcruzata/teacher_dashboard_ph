#include <stdio.h>
#include <time.h>

// 1. MACRO: Parentesi d'obbligo su parametri ed espressione
#define QUADRATO_MACRO(x) ((x) * (x))

// 2. INLINE: Sicura su tipi ed effetti collaterali (es. ++x)
static inline int quadrato_inline(int x) { return x * x; }

// 3. HEADER GUARD PATTERN (per file .h)
#ifndef CHEATSHEET_H
#define CHEATSHEET_H
// Dichiarazioni header...
#endif

int main(void) {
  // --- TIME.H ---
  clock_t start = clock();   // Misura tempo CPU
  time_t t_ora = time(NULL); // Secondi da Epoch (1/1/1970)
  struct tm *loc =
      localtime(&t_ora); // Scompone in data (mon: 0-11, year: +1900)

  char buf[32];
  strftime(buf, sizeof(buf), "%Y-%m-%d %H:%M:%S", loc); // Formattazione custom
  printf("Data: %s | CPU: %.4fs\n", buf,
         (double)(clock() - start) / CLOCKS_PER_SEC);

  // --- PREPROCESSORE: INLINE VS MACRO ---
  int a = 3, b = 3;
  printf("Inline (++a): %d\n", quadrato_inline(++a)); // Sicuro -> 4*4 = 16
  printf("Macro  (val): %d\n",
         QUADRATO_MACRO(
             b)); // Usare ++b dentro la macro causerebbe doppie valutazioni!

  return 0;
}