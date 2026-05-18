# 📊 Configurador de Tributos Protheus

App web para gerar a **Planilha de Apoio do Configurador de Tributos** do TOTVS Protheus (Livros Fiscais) a partir de uma tabela SFT.

O app cruza automaticamente a SFT com:
- **CBENEF** (Códigos de Benefício Fiscal) — 276 códigos de Goiás classificados manualmente
- **Tabela NCM oficial** — para descrição dos produtos
- **Tabela CFOP** — com descrições oficiais

E aplica **filtragem inteligente** baseada em:
1. **NCM** dos produtos da sua SFT (descarta CBENEFs de produtos que você não vende)
2. **Perfil da empresa** (questionário de 31 perguntas — descarta CBENEFs incompatíveis com o seu negócio)

## ✨ O que o app gera

Uma planilha Excel com 10 abas estruturadas:

| Aba | Descrição |
|---|---|
| **RESUMO** | Visão geral dos cenários encontrados |
| **REGRA DE CALCULO** | Códigos de cálculo por imposto/cenário |
| **REGRA VALOR DECLARATÓRIO** | ⭐ CBENEFs sugeridos com SIM/TALVEZ/NÃO |
| **PERFIL ORIGEM DESTINO** | Combinações UF origem/destino |
| **PERFIL PARTICIPANTE** | Tipos de participante (contribuinte, não contrib.) |
| **PERFIL OPERAÇÃO** | Combinações de CFOP × CST |
| **PERFIL PRODUTO** | Combinações de NCM × CST |
| **REGRA BASE / ALÍQUOTA / ESCRITURAÇÃO** | Detalhamento técnico por imposto |

A aba **REGRA VALOR DECLARATÓRIO** vem com codificação por cor:
- 🟢 **SIM** — Candidato forte, use estes códigos
- 🟡 **TALVEZ** — Revisar manualmente
- ⚪ **NÃO** — Descartado pelo perfil (mantido para auditoria)

## 🚀 Como rodar localmente

### Pré-requisitos
- Python 3.10+
- pip

### Instalação
```bash
# 1. Clone o repositório
git clone https://github.com/SEU_USUARIO/configurador-tributos-protheus.git
cd configurador-tributos-protheus

# 2. (Opcional, recomendado) Crie um ambiente virtual
python -m venv .venv
source .venv/bin/activate   # Linux/Mac
# .venv\Scripts\activate    # Windows

# 3. Instale as dependências
pip install -r requirements.txt

# 4. Execute o app
streamlit run app.py
```

O navegador abrirá automaticamente em `http://localhost:8501`.

## 🌐 Deploy no Streamlit Cloud (grátis)

1. Faça push do repo para o GitHub
2. Acesse [share.streamlit.io](https://share.streamlit.io)
3. Conecte sua conta GitHub
4. Clique em **New app** → selecione o repositório → branch `main` → arquivo `app.py`
5. Clique em **Deploy** ✨

Em ~1 minuto o app estará público em uma URL tipo `https://configurador-tributos.streamlit.app`.

## 📁 Estrutura do projeto

```
configurador-tributos-protheus/
├── app.py                          ← Interface Streamlit
├── processador.py                  ← Lógica de negócio (cruzamento SFT × CBENEF × NCM)
├── cbenef_classificacao_GO.py      ← Banco de classificação dos 276 CBENEFs de GO
├── requirements.txt                ← Dependências Python
├── README.md                       ← Este arquivo
└── data/                           ← Tabelas oficiais (fixas no app)
    ├── CBENEF_GO_x_CFOP.xlsx       ← 276 códigos CBENEF de Goiás
    ├── Tabela_NCM_Vigente.xlsx     ← Tabela NCM oficial
    ├── Tabela_CFOP.xlsx            ← Descrições de CFOPs
    └── Modelo_planilha_apoio.xlsx  ← Template da planilha gerada
```

## 🔧 Como usar (passo a passo)

### Passo 1 — Upload da SFT
Faça upload do arquivo `SFT.xlsx` extraído do Protheus. O app detecta a aba automaticamente.

### Passo 2 — Perfil da empresa
Responda às 31 perguntas agrupadas em:
- **Operações comuns** (exporta, vende para ZFM, faz doações, etc.)
- **Regime e localização** (agropecuário, Simples Nacional, RIDE)
- **Setores específicos** (farmacêutico, telecom, energia, sucroenergético, etc.)

Em caso de dúvida, deixe desmarcado — o código aparecerá como ⚠ TALVEZ para revisão posterior.

### Passo 3 — Download
Clique em **Baixar Planilha** para receber o arquivo final. Pronto!

## 🧠 Como funciona a filtragem inteligente?

Os 276 CBENEFs de Goiás foram classificados manualmente em 4 categorias:

| Categoria | Comportamento |
|---|---|
| **GENÉRICO** (13) | Aplica a qualquer empresa — sempre mantém |
| **PRODUTO** (151) | Específico de tipo de produto — só aplica se o NCM bate com sua SFT |
| **CONTEXTO** (107) | Específico de contexto (doação, órgão público, etc.) — só aplica se o perfil confirma |
| **IGNORE** (5) | "Reconhecida judicialmente" — nunca aplica em cadastro padrão |

Além disso, cada CBENEF tem **flags** (ex.: `farmaceutica`, `exporta`, `agropecuario`). Se o perfil da empresa nega TODAS as flags relevantes, o CBENEF é marcado como ❌ NÃO automaticamente.

## 🛠️ Uso via linha de comando (sem app web)

Também é possível usar como script CLI:

```bash
python processador.py minha_sft.xlsx -o saida.xlsx
```

Ele vai pedir as 31 perguntas no terminal e gerar a planilha.

## 📋 Versão

**v22** — Filtragem inteligente por NCM + perfil da empresa

## 📝 Licença

MIT — use à vontade, contribuições bem-vindas.
