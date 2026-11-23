# Computadores para Inclusão – dados

Robôs (GitHub Actions):
- **Atualizar dados** – todo dia 6h23 (reserva 9h47): baixa o BI, calcula tudo e grava `data/`.
- **Matriculas QEdu** – todo dia 2h17, só trabalha se houver escola sem matrícula: consulta o QEdu pelo código INEP.

Impacto social (regra da área gestora): instituição de ensino = matrículas no QEdu (conta uma vez por escola);
demais pontos = 5 × computadores doados. Escola sem matrícula confirmada deixa o município pendente (a página mantém o valor anterior).

Arquivos que você mantém em `data/fixos/`:
- `LISTA_ESCOLAS.xlsx` – lista da área (quais pontos são instituição de ensino e o código INEP).
- `matriculas_manual.csv` – só para escola fora do QEdu (ex.: código e-MEC), sempre com fonte e link.
- `municipios_alias.csv` – correções de nome de município (linha só vale com `aplicar = sim`).

Para conferir: `data/impacto_pontos.csv`, `data/escolas_pendentes.csv`, `data/municipios_sugestoes.csv`.
