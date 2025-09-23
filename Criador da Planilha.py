import pandas as pd

# Dados fictícios para teste
dados = [
    {"nome": "Maria Clara", "email": "maria@example.com", "senha": "senha123"},
    {"nome": "João Silva", "email": "joao@example.com", "senha": "joao@456"},
    {"nome": "Ana Paula", "email": "ana@example.com", "senha": "ana789"},
]

# Cria o DataFrame
df = pd.DataFrame(dados)

# Salva como Excel
df.to_excel("dados.xlsx", index=False)

print("✅ Planilha 'dados.xlsx' criada com sucesso!")