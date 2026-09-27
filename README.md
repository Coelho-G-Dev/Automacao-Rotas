# Otimizador de Rotas e Alocação de Frota — Setor de Frios

[![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-3776AB.svg?style=flat&logo=python&logoColor=white)](https://www.python.org/)
[![Pandas](https://img.shields.io/badge/Pandas-2.0%2B-150458.svg?style=flat&logo=pandas&logoColor=white)](https://pandas.pydata.org/)
[![OpenPyXL](https://img.shields.io/badge/OpenPyXL-3.1%2B-217346.svg?style=flat&logo=microsoftexcel&logoColor=white)](https://openpyxl.readthedocs.io/)
[![Logistics](https://img.shields.io/badge/Operations-Fleet%20Optimization-EA580C.svg?style=flat)](https://github.com/Coelho-G-Dev/Automacao-Rotas)
[![License](https://img.shields.io/badge/License-MIT-059669.svg?style=flat)](https://opensource.org/licenses/MIT)

Script em Python para automação da roteirização e consolidação de cargas da expedição de produtos frigorificados (resfriados e congelados) a partir dos relatórios de separação do ERP.

O objetivo principal é eliminar a perda de frete (subutilização da capacidade contratada) sem ultrapassar os limites físicos dos veículos, respeitando os corredores rodoviários, restrições físicas de docas e os ciclos de geração de cada filial.

---

## Modelo Matemático e Regras de Decisão

O motor combinatorial busca maximizar a taxa de ocupação dos veículos despachados, priorizando o despacho mandatório dos pedidos gerados no dia corrente e utilizando as sobras de ciclos anteriores para complementar o peso operacional:

$$\max \sum_{v \in V} \left( P_{\text{hoje}}(v) + P_{\text{sobra}}(v) \right)$$

### Restrições Operacionais

1. **Condição de Piso Mínimo (Prevenção de Frete Morto)**:
   $$P_{\text{carga}}(v) \ge P_{\text{piso}}(v)$$
   Garante que nenhum caminhão saia subutilizado (ex: Toco $\ge 8.000\text{ kg}$, Truck $\ge 12.000\text{ kg}$, Carreta $\ge 24.000\text{ kg}$).

2. **Condição de Teto Máximo com Tolerância Segura**:
   $$P_{\text{carga}}(v) \le C_{\text{nominal}}(v) + \Delta_{\text{tol}}$$
   Onde a tolerância máxima permitida para viabilizar o fechamento da rota é de $\Delta_{\text{tol}} = 100\text{ kg}$.
   $$\text{Se } P_{\text{carga}}(v) > C_{\text{nominal}}(v) + 100\text{ kg} \implies \text{Bloqueio Imediato por Risco de Quebra}$$

3. **Limite de Paradas por Veículo**:
   $$N_{\text{lojas}}(v) \le 4 \quad (\forall v \in V)$$
   Na Capital, para Carreta exclusiva de alto volume, adota-se $N_{\text{lojas}}(v) = 1$.

---

## Parâmetros da Frota

| Veículo | Piso Operacional ($P_{\text{piso}}$) | Capacidade Nominal ($C_{\text{nom}}$) | Teto Máximo Permissível ($C_{\text{nom}} + 100\text{ kg}$) | Atuação |
| :--- | :---: | :---: | :---: | :--- |
| **3/4** | $5.000\text{ kg}$ | $6.000\text{ kg}$ | $6.100\text{ kg}$ | Exclusivo Capital (docas estreitas / rotas urbanas) |
| **Toco** | $8.000\text{ kg}$ | $9.000\text{ kg}$ | $9.100\text{ kg}$ | Capital e Interior |
| **Truck** | $12.000\text{ kg}$ | $14.000\text{ kg}$ | $14.100\text{ kg}$ | Capital e Interior |
| **Bitruck** | $15.500\text{ kg}$ | $16.000\text{ kg}$ | $16.100\text{ kg}$ | Capital e Interior (exceto filiais com restrição) |
| **Carreta** | $24.000\text{ kg}$ | $28.000\text{ kg}$ | $28.100\text{ kg}$ | Interior ($\le 4$ paradas) e Capital ($1$ parada) |

---

## Corredores Rodoviários (Interior)

As entregas do Interior são agrupadas respeitando os eixos de transporte:

- **Corredor 1 (Lençóis / BR-402)**: Rosário, Barreirinhas, Tutóia.
- **Corredor 2 (Baixada / MA-014)**: Arari, Viana, Pinheiro, Santa Helena, Cururupu.
- **Corredor 3 (BR-222 / Baixo Parnaíba)**: Itapecuru, Vargem Grande, Chapadinha, Urbano Santos, Coelho Neto e Parnaíba (PI).
- **Corredor 4 (Centro / BR-135 e BR-316)**: Miranda do Norte, São Mateus, Coroatá, Codó.
- **Corredor 5 (Leste / BR-316)**: Caxias, Timon, Teresina (PI), Floriano (PI).
- **Polo Transferência Ceará**: Fortaleza e região metropolitana, Sobral, Piripiri (PI) — exclusivo Carreta ($28.000\text{ kg}$).

---

## Estrutura da Planilha Gerada

O processamento gera o arquivo `Prototipo_Rotas_Frios.xlsx` estruturado em 6 abas operacionais:

- **`ROTAS_CAP`**: Espelho operacional da expedição Capital.
- **`ROTA_INT`**: Espelho operacional do Interior com agrupamento por corredores.
- **`Roteirização Consolidada`**: Resumo das rotas aprovadas, paradas, pesos e percentual de ocupação útil.
- **`Detalhe Lojas e Endereços`**: Rastreabilidade analítica com status de ciclo (`DIAS_GERAÇÃO`), tipo de demanda e restrições veiculares.
- **`Aguardando Geração (Bater Peso)`**: Lojas retidas para aguardar complementação de piso ou prevenir sobrecarga.
- **`Sobras (Não Roteirizadas)`**: Saldo analítico de mercadorias pendentes de expedição.

---

## Instalação e Execução

### Pré-requisitos
- Python 3.10 ou superior

### Instalação
```bash
git clone https://github.com/Coelho-G-Dev/Automacao-Rotas.git
cd Automacao-Rotas
pip install -r requirements.txt
```

### Execução
Pelo terminal:
```bash
python automacao.py
```
Ou no Windows executando o arquivo `executar_roteirizacao.bat`.

---

## Testes e Validação

Os scripts de teste e simulação estão concentrados na pasta `scratch/`:
- `scratch/executar_bateria_testes_rotas.py`: Suíte de testes com validação de restrições veiculares, pontos de peso, ciclos de pedidos e lojas âncora.
- `scratch/benchmark_3dias_completo.py`: Comparativo de aderência entre as rotas automatizadas e o histórico real da operação.

---

## Licença

Distribuído sob licença MIT. Veja `LICENSE` para mais informações.
