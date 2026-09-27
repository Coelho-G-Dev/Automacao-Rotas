# Otimizador de Rotas e Alocação de Frota — Setor de Frios

Script em Python para automação da roteirização e consolidação de cargas da expedição de produtos frigorificados (resfriados e congelados) a partir dos relatórios de separação do ERP.

O objetivo principal é eliminar a perda de frete (subutilização de capacidade contratada) sem ultrapassar o limite físico dos veículos, respeitando os corredores rodoviários, restrições físicas de docas e os ciclos de geração de cada filial.

---

## Como Funciona

1. **Ingestão de Dados**: Lê a base de separação (`Acomp_EvoluSep.xlsx`), filtra o setor de frios e descarta filiais de condomínio/apoio, preservando lojas operacionais.
2. **Enriquecimento Cadastral**: Cruza as filiais com o cadastro mestre (`Cadastro_Lojas_Enderecos.xlsx`) para obter coordenadas, rota padrão, restrições de doca e ciclo de pedidos (`DIAS_GERAÇÃO`).
3. **Agrupamento por Corredores**: Separa a demanda entre Capital e Interior, organizando o Interior pelos eixos rodoviários (Lençóis, Baixada, BR-222, Centro, Leste e Transferência CE).
4. **Alocação e Otimização**:
   - Prioriza o envio de lojas com pedidos gerados no dia (`01-DE HOJE`).
   - Utiliza sobras de ciclos anteriores (`02-NO PRAZO` / `03-URGENTE`) apenas como complemento para atingir o piso operacional do veículo.
   - Aplica a trava de sobrepeso: aceita até 100 kg de excedente para viabilizar fechamento de carga; acima disso, bloqueia o veículo para evitar avaria mecânica.
5. **Geração do Relatório**: Exporta a planilha operacional (`Prototipo_Rotas_Frios.xlsx`) com abas dedicadas para Capital, Interior, visão consolidada e cargas pendentes de piso.

---

## Parâmetros de Frota

| Veículo | Piso Operacional | Capacidade Nominal | Teto Máximo (Tolerância +100 kg) | Aplicação |
| :--- | :---: | :---: | :---: | :--- |
| **3/4** | 5.000 kg | 6.000 kg | 6.100 kg | Exclusivo Capital (docas estreitas / rotas urbanas) |
| **Toco** | 8.000 kg | 9.000 kg | 9.100 kg | Capital e Interior |
| **Truck** | 12.000 kg | 14.000 kg | 14.100 kg | Capital e Interior |
| **Bitruck** | 15.500 kg | 16.000 kg | 16.100 kg | Capital e Interior (exceto filiais com restrição) |
| **Carreta** | 24.000 kg | 28.000 kg | 28.100 kg | Interior (até 4 paradas) e Capital (1 parada) |

---

## Corredores Rodoviários (Interior)

- **Corredor 1 (Lençóis / BR-402)**: Rosário, Barreirinhas, Tutóia.
- **Corredor 2 (Baixada / MA-014)**: Arari, Viana, Pinheiro, Santa Helena, Cururupu.
- **Corredor 3 (BR-222 / Baixo Parnaíba)**: Itapecuru, Vargem Grande, Chapadinha, Urbano Santos, Coelho Neto e Parnaíba (PI).
- **Corredor 4 (Centro / BR-135 e BR-316)**: Miranda do Norte, São Mateus, Coroatá, Codó.
- **Corredor 5 (Leste / BR-316)**: Caxias, Timon, Teresina (PI), Floriano (PI).
- **Transferência Ceará**: Fortaleza e região metropolitana, Sobral, Piripiri (PI) — exclusivo Carreta.

---

## Estrutura da Planilha Gerada

O arquivo `Prototipo_Rotas_Frios.xlsx` contém 6 abas:
- **`ROTAS_CAP`**: Espelho operacional da expedição Capital.
- **`ROTA_INT`**: Espelho operacional do Interior agrupado por corredores.
- **`Roteirização Consolidada`**: Resumo executivo das rotas aprovadas, paradas e ocupação.
- **`Detalhe Lojas e Endereços`**: Relação detalhada de cada entrega com ciclo de geração e restrições.
- **`Aguardando Geração (Bater Peso)`**: Lojas retidas por falta de peso mínimo ou risco de sobrepeso.
- **`Sobras (Não Roteirizadas)`**: Saldo analítico de cargas pendentes.

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

Os scripts de simulação e benchmark ficam localizados na pasta `scratch/`:
- `scratch/executar_bateria_testes_rotas.py`: Bateria com 6 cenários de teste (restrições de veículos, faixas de peso, ciclos de pedidos e lojas âncora).
- `scratch/benchmark_3dias_completo.py`: Comparador de aderência entre as rotas do algoritmo e o histórico real da operação.
