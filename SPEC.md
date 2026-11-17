Todo servidor tem um horário de trabalho diferenciado. Uns começam cedo, outros começam mais tarde.

Relação de horas extras:

| Dia | Hora | HE |
| --- | --- | --- |
| Seg-Sáb | 05:00 - 22:00 | .5 |
| Seg-Sáb | 22:00 - 05:00 | .75 |
| Dom e Feriado | 05:00 - 22:00 | 1 |
| Dom e Feriado | 22:00 - 05:00 | 1.25 |

Entrada:

- PDF com jornadas de todos os colaboradores (1 por página)

Saída:

- Planilha excel com jornadas organizadas

  - jornadas organizadas consirando o horário de cada colaborador, com as horas extras calculadas

- Colunas:

  - Entrada, Saída, Almoço, Retorno, Entrada, Saída, Entrada (HE), Saída (HE)
  - Totalizadores: somatório de cada HE: 50%, 75%, 100%, 125% e total de horas trabalhadas

Fatos:

- Ninguém começa de madrugada, o sistema "pula" erroneamente. Considerar que início de jornada de trabalho de madrugada é na verdade a saída do dia anterior. Exemplo:

  - Domingo, 11/09, 03:00: Considerar como saída de sábado, 10/09, 03:00

- Folha fecha dia 11 e vai até dia 10 de cada mês

Requisitos:

- Finais de semana e feriados em cores diferenciadas
- Horas extras em cores diferenciadas

Exceções:

- gente que não bate o ponto. Eles esquecem

- em plantões de final de semana e feriados o horário do almoço, deve ser preenchido automaticamente. Dia de semana, eles batem. Se eles esquecerem de bater, deixar horário livre para preenchimento manual

- FALTA é falta

- COMPENSADO ignorar

- LICENÇA ignorar

- RSR ignorar

- marcações sem sentido, deixar para revisão manual. Exemplos:

  - Segunda: 06:40 12:10 12:37 15:56 16:26
  - Terça:   02:00 06:25 12:00 13:01 15:50 18:09 00:30

  Aqui, o horário 02:00 de terça deveria ser a saída de segunda.

  - Segunda: 12:02 00:01

  Considerar 8h trabalhadas, e dali em diante é hora extra.


