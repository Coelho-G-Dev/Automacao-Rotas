import sys, os
sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd
import openpyxl

# Corredores Oficiais do Interior alinhados com as diretrizes do usuário
CORREDORES_INTERIOR = {
    "Rota 1 - Lençóis / Litoral Norte": {
        "lojas_fixas": ["222", "424", "434", "537", "359", "447"],
        "permite_429": False,
        "permite_439": False,
        "descricao": "BR-402: Rosário -> Barreirinhas -> Tutóia (Sem Santa Rita)"
    },
    "Rota 2 - Baixada Maranhense": {
        "lojas_fixas": ["427", "433", "457", "99", "454", "459"],
        "permite_429": True,
        "permite_439": True, # Miranda faz rota com lojas sentido Arari 427
        "descricao": "Baixada: Arari, Viana, São Bento, Pinheiro, Santa Helena, Cururupu"
    },
    "Rota 3 - BR-222 / Baixo Parnaíba": {
        "lojas_fixas": ["445", "451", "39", "461", "215"],
        "permite_429": True, # Todas passam por Santa Rita na BR-135 antes do entroncamento
        "permite_439": False,
        "descricao": "BR-222: Itapecuru, Vargem Grande, Chapadinha, Urbano Santos, Coelho Neto"
    },
    "Rota 4 - Centro / Mearim / Cocais": {
        "lojas_fixas": ["439", "220", "435", "202"],
        "permite_429": True, # Passam por Santa Rita na BR-135
        "permite_439": True, # Miranda do Norte é o próprio eixo da BR-135
        "descricao": "BR-135 / BR-316: Miranda do Norte, São Mateus, Coroatá, Codó"
    },
    "Rota 5 - Caxias / Timon / Floriano / Teresina": {
        "lojas_fixas": ["41", "230", "32", "275", "271", "97", "252", "258", "560", "563", "251"],
        "permite_429": False,
        "permite_439": False,
        "descricao": "BR-316 / BR-343: Caxias, Timon, Teresina, Floriano, Parnaíba"
    },
    "Transferência - Ceará": {
        "lojas_fixas": [
            "506", "527", "534", "503", "514", "518", "520", "264", "266", "280",
            "541", "285", "290", "507", "516", "211"
        ],
        "permite_429": False,
        "permite_439": False,
        "descricao": "Ceará / Crossdock Transferência"
    }
}

print("=== CORREDORES LOGÍSTICOS OFICIAIS DEFINIDOS ===")
for nome, c in CORREDORES_INTERIOR.items():
    print(f"\n{nome}:")
    print(f"   Descrição: {c['descricao']}")
    print(f"   Lojas Fixas: {c['lojas_fixas']}")
    print(f"   Permite Loja 429 (Santa Rita): {c['permite_429']}")
    print(f"   Permite Loja 439 (Miranda do Norte): {c['permite_439']}")
