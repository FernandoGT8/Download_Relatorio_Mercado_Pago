# Mercado Pago - Download de Vendas

Automação em Python para acessar o extrato do Mercado Pago via navegador, filtrar movimentações por período e exportar as vendas encontradas para uma planilha Excel.

O projeto usa **Playwright** porque, neste fluxo específico, a API do Mercado Pago não entrega todas as informações necessárias para conciliação, principalmente detalhes operacionais acessíveis apenas pela interface web.

---

## Funcionalidades

O robô executa o fluxo completo de extração:

- Abre o Mercado Pago no navegador.
- Utiliza uma sessão já autenticada do usuário.
- Permite selecionar o período de busca:
  - mês inteiro;
  - ontem / pendente;
  - data específica;
  - período específico.
- Filtra o extrato para exibir entradas.
- Percorre as movimentações encontradas.
- Abre os detalhes das vendas quando necessário.
- Extrai os seguintes dados:
  - data da venda;
  - número da venda;
  - categoria;
  - valor líquido;
  - valor bruto.
- Exporta os dados para Excel.
- Evita duplicidade na planilha.
- Gera log dos itens que não puderem ser processados.

---

## Requisitos

- Python 3.10 ou superior
- Google Chrome instalado
- Conta Mercado Pago com acesso ao extrato
- Login manual realizado na primeira execução

---

## Instalação

Clone o repositório:

```bash
git clone https://github.com/seu-usuario/seu-repositorio.git
cd seu-repositorio
```

Crie e ative um ambiente virtual:

### Windows

```bash
python -m venv .venv
.venv\Scripts\activate
```

### Linux / macOS

```bash
python3 -m venv .venv
source .venv/bin/activate
```

Instale as dependências:

```bash
pip install -r requirements.txt
```

Caso ainda não tenha o Playwright configurado:

```bash
playwright install
```

---

## Dependências

O projeto utiliza:

```txt
playwright
openpyxl
rich
```

Sugestão de `requirements.txt`:

```txt
playwright>=1.40.0
openpyxl>=3.1.0
rich>=13.0.0
```

---

## Como executar

Execute:

```bash
python mercado_pago_robo.py
```

Na primeira execução, o navegador será aberto e você deverá fazer login manualmente no Mercado Pago.

Depois disso, a sessão será reaproveitada por meio do perfil local do navegador.

---

## Configuração

Por padrão, o projeto usa:

| Variável | Padrão | Descrição |
|---|---|---|
| `PERFIL_AUTOMACAO` | `perfil_google` | Pasta usada para armazenar a sessão do navegador |
| `PASTA_BASE_RELATORIOS` | `output` | Pasta onde os relatórios Excel serão salvos |

### Windows CMD

```bash
set PERFIL_AUTOMACAO=perfil_google
set PASTA_BASE_RELATORIOS=output
```

### PowerShell

```powershell
$env:PERFIL_AUTOMACAO="perfil_google"
$env:PASTA_BASE_RELATORIOS="output"
```

### Linux / macOS

```bash
export PERFIL_AUTOMACAO=perfil_google
export PASTA_BASE_RELATORIOS=output
```

---

## Estrutura esperada

Exemplo de estrutura do projeto:

```text
.
├── mercado_pago_robo.py
├── requirements.txt
├── README.md
├── .gitignore
├── perfil_google/        # não versionar
└── output/               # não versionar
```

---

## Arquivos gerados

O robô pode gerar:

- planilhas `.xlsx`;
- logs de erro;
- arquivos temporários do perfil do navegador.

Esses arquivos não devem ser versionados.

Exemplo recomendado no `.gitignore`:

```gitignore
perfil_google/
chrome-profile/
output/
logs/
*.xlsx
*.xls
*.xlsm
*.log
.env
```

---

## Observações importantes

Este projeto depende da interface web do Mercado Pago. Alterações no layout, textos, botões ou fluxo de navegação podem exigir ajustes no código.

A automação não é afiliada, mantida ou homologada pelo Mercado Pago.

Os dados exportados devem ser conferidos pelo usuário antes de qualquer uso contábil, financeiro ou fiscal.

---

## Limitações

O robô não substitui conciliação financeira oficial.

Ele automatiza a coleta de dados exibidos na interface do Mercado Pago, mas ainda depende de:

- sessão ativa;
- estabilidade da página;
- layout atual do Mercado Pago;
- permissões da conta logada;
- disponibilidade dos detalhes da movimentação.

---

## Segurança

Não publique arquivos que contenham sessão, cookies, planilhas reais, logs ou dados de clientes.

Nunca versionar:

```text
perfil_google/
output/
.env
*.xlsx
*.log
```

Se algum desses arquivos já foi commitado, remova do controle de versão:

```bash
git rm -r --cached perfil_google output
git rm --cached .env
git rm --cached *.xlsx *.log
```

Depois faça um novo commit:

```bash
git add .gitignore
git commit -m "chore: remove local files from version control"
```

---

## Status do projeto

Projeto funcional para uso operacional interno.

Como a automação depende da interface do Mercado Pago, eventuais mudanças na plataforma podem quebrar seletores ou etapas do fluxo.

---

## Licença

Defina a licença conforme o objetivo do repositório.

Para uso interno ou privado, uma licença pública não é obrigatória.

Para uso aberto, considere adicionar um arquivo `LICENSE`.