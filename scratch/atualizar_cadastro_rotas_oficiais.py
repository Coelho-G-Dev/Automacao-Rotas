import os, sys
import pandas as pd
import openpyxl

caminho_cad = r"C:\Users\gabri\Desktop\Separacao\Cadastro_Lojas_Enderecos.xlsx"
df_cad = pd.read_excel(caminho_cad)

# Mapeamento Oficial das Rotas 1 a 12 da Capital
mapa_rotas_capital = {
    # Rota 1: Renascença / São Francisco
    "34": "1",
    "2061": "1",
    "61": "1",
    "60": "1",
    "90": "1",
    "225": "1",
    "229": "1",
    
    # Rota 2: Cohama
    "7": "2",
    "1413": "2",
    
    # Rota 3: Calhau / Maranhão Novo / Shopping da Ilha
    "16": "3",
    "19": "3",
    "411": "3",
    "350": "3",
    "1411": "3",
    
    # Rota 4: Itaqui-Bacanga / Anjo da Guarda
    "30": "4",
    "408": "4",
    "539": "4",
    "1418": "4",
    
    # Rota 5: Estrada de Ribamar / Cidade Operária / Santa Clara / SJ Ribamar
    "11": "5",
    "502": "5",
    "425": "5",
    "218": "5",
    "1416": "5",
    
    # Rota 6: Tirirical / Raposa
    "17": "6",
    "217": "6",
    "122": "6",
    "81": "6",
    "356": "6",
    
    # Rota 7: Centro / Anil / Cajazeiras / Camboa
    "23": "7",
    "200": "7",
    "12": "7",
    "526": "7",
    "1417": "7",
    
    # Rota 8: João Paulo / Vinhais
    "96": "8",
    "20": "8",
    "1407": "8",
    
    # Rota 9: Turu / Arthur Carvalho / Divinéia
    "8": "9",
    "27": "9",
    "232": "9",
    "418": "9",
    "477": "9",
    "1410": "9",
    
    # Rota 10: Jardim Tropical / São Cristóvão
    "29": "10",
    "410": "10",
    "415": "10",
    "1415": "10",
    
    # Rota 11: Cohab / Cohatrac
    "5": "11",
    "201": "11",
    "18": "11",
    "412": "11",
    
    # Rota 12: Estrada da Maioba / Maiobão / Paço do Lumiar
    "93": "12",
    "255": "12",
    "259": "12",
    "269": "12",
    "409": "12",
    "554": "12",
    "31": "12",
    "1408": "12",
    "1409": "12",
}

# Mapeamento do Interior
mapa_rotas_interior = {
    # Caxias / Timon / Floriano
    "41": "Rota 1 - Caxias / Floriano",
    "230": "Rota 1 - Caxias / Floriano",
    "271": "Rota 1 - Caxias / Floriano",
    "275": "Rota 1 - Caxias / Floriano",
    
    # Chapadinha / Codó
    "39": "Rota 2 - Chapadinha / Codó",
    "202": "Rota 2 - Chapadinha / Codó",
    "215": "Rota 2 - Chapadinha / Codó",
    "220": "Rota 2 - Chapadinha / Codó",
    
    # Baixada / Rosário
    "222": "Rota 3 - Baixada / Rosário",
    "424": "Rota 3 - Baixada / Rosário",
    "427": "Rota 3 - Baixada / Rosário",
    "429": "Rota 3 - Baixada / Rosário",
    "433": "Rota 3 - Baixada / Rosário",
    "457": "Rota 3 - Baixada / Rosário",
    "459": "Rota 3 - Baixada / Rosário",
    "99": "Rota 3 - Baixada / Rosário",
    "445": "Rota 3 - Baixada / Rosário",
    
    # Barreirinhas / Tutóia
    "434": "Rota 4 - Barreirinhas / Tutóia",
    "447": "Rota 4 - Barreirinhas / Tutóia",
    "537": "Rota 4 - Barreirinhas / Tutóia",
    
    # Piauí
    "97": "Rota 5 - Teresina / Piauí",
    "211": "Rota 5 - Teresina / Piauí",
    "251": "Rota 5 - Teresina / Piauí",
    "252": "Rota 5 - Teresina / Piauí",
    "258": "Rota 5 - Teresina / Piauí",
    "560": "Rota 5 - Teresina / Piauí",
    "563": "Rota 5 - Teresina / Piauí",
    
    # Santa Luzia
    "439": "Rota 6 - Santa Luzia / Miranda",
}

# Aplicar atualizações no DataFrame
atualizados = 0
for idx, row in df_cad.iterrows():
    f = str(row["Filial"]).strip()
    reg = str(row["Região"]).strip().upper()
    
    if reg == "CAPITAL":
        if f in mapa_rotas_capital:
            r_nova = mapa_rotas_capital[f]
            df_cad.at[idx, "Rota Padrão"] = r_nova
            atualizados += 1
    else:
        if f in mapa_rotas_interior:
            r_nova = mapa_rotas_interior[f]
            df_cad.at[idx, "Rota Padrão"] = r_nova
            atualizados += 1

# Verificar se a Loja 2061 está cadastrada
if "2061" not in df_cad["Filial"].astype(str).values:
    nova_linha = {
        "Filial": "2061",
        "Nome Loja": "2061 - MATEUS FOOD RENASCENCA",
        "Região": "CAPITAL",
        "Rota Padrão": "1",
        "Restrição Veículo": "Sem restrição",
        "Bairro": "Jardim Renascença",
        "Cidade": "São Luís",
        "UF": "MA",
        "Endereço / Referência": "Jardim Renascença, São Luís - MA",
        "Latitude": -2.4998,
        "Longitude": -44.2952,
        "Observação": "Cadastrado automaticamente - Rota 1 (Renascença)"
    }
    df_cad = pd.concat([df_cad, pd.DataFrame([nova_linha])], ignore_index=True)
    print("[OK] Loja 2061 inserida no cadastro.")

# Salvar
df_cad.to_excel(caminho_cad, index=False)
print(f"[OK] Cadastro_Lojas_Enderecos.xlsx atualizado com sucesso! ({atualizados} lojas sincronizadas)")
