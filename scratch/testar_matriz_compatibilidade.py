import sys, os
sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd

# Matriz de compatibilidade do Interior
COMPATIBILIDADE_INTERIOR = {
    "Rota 1 - Lençóis / Litoral Norte": {
        "lojas_permitidas": {"222", "424", "434", "537", "359", "447"},
        "proibidas": {"429", "457", "459", "99", "427", "433", "454", "439"} # Sem Santa Rita, Sem Baixada
    },
    "Rota 2 - Baixada Maranhense": {
        "lojas_permitidas": {"427", "433", "457", "99", "454", "459", "429", "439"},
        "proibidas": {"222", "424", "434", "537", "447"} # Sem Lençóis
    },
    "Rota 3 - BR-222 / Baixo Parnaíba": {
        "lojas_permitidas": {"445", "451", "39", "461", "215", "429"},
        "proibidas": {"222", "424", "434", "537", "447", "457", "459"}
    },
    "Rota 4 - Centro / Mearim / Cocais": {
        "lojas_permitidas": {"439", "220", "435", "202", "429"},
        "proibidas": {"222", "424", "434", "537", "447", "457", "459"}
    },
    "Rota 5 - Caxias / Timon / Floriano": {
        "lojas_permitidas": {"41", "230", "32", "275", "271", "97", "252", "258", "560", "563", "251"},
        "proibidas": {"222", "424", "434", "537", "447", "429", "457", "459"}
    },
    "Transferência - Ceará (CE)": {
        "lojas_permitidas": {
            "506", "527", "534", "503", "514", "518", "520", "264", "266", "280",
            "541", "285", "290", "507", "516", "211"
        },
        "proibidas": set()
    }
}

print("=== REGRAS DE COMPATIBILIDADE E PROIBIÇÕES DO INTERIOR ===")
for r, regras in COMPATIBILIDADE_INTERIOR.items():
    print(f"\n{r}:")
    print(f"   Permitidas: {sorted(list(regras['lojas_permitidas']))}")
    print(f"   Proibidas:  {sorted(list(regras['proibidas']))}")
