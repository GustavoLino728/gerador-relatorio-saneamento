# 📄 Automação de Relatórios de Fiscalização com Python

Este projeto realiza a **automação da geração de relatórios técnicos** a partir de uma planilha Excel contendo dados de fiscalizações e de pastas com fotos associadas às não conformidades identificadas. O relatório final é gerado em formato `.docx` com formatação e organização automáticas.

---

## 🚀 Funcionalidades

- 📥 Leitura da planilha Excel com dados de fiscalizações e não conformidades
- 🧠 Identificação automática da próxima fiscalização que ainda não teve relatório gerado
- 🖼️ Extração e inserção automática de imagens em tabelas formatadas
- 📝 Geração de legenda automática para cada imagem com base na planilha
- 🗂️ Organização das imagens em blocos, com até 6 por tabela (3 linhas × 2 colunas)
- 📄 Criação automática do relatório Word baseado em um modelo (`RELATÓRIO MODELO.docx`)
- 📊 Inclusão de tabelas com bordas e espaçamento adequado

---

## 📁 Estrutura do Projeto

Automacao-Relatorios-ARPE/
├── assets/ # Pasta com as imagens das não conformidades
├── data/
│ ├── Listagem das NC's.xlsx # Planilha com dados de fiscalizações
│ └── RELATÓRIO MODELO.docx # Modelo de documento base
├── src/
│ ├── excel.py # Módulo de leitura e filtragem de dados da planilha
│ ├── images.py # Módulo de processamento das imagens e geração de tabelas
│ ├── report.py # Lógica principal de geração do relatório
│ ├── utils.py # Funções auxiliares (ex: substituição de variáveis)
│ └── main.py # Script principal que executa toda a automação
├── README.md
└── requirements.txt

yaml
Copiar
Editar

---

## 📊 Requisitos da Planilha

- A aba `Fiscalizações` deve conter a coluna: `Relatório Gerado` (valores booleanos).
- A aba `Nao-conformidades` deve conter:
  - `ID da Fiscalização`
  - `Unidade`
  - `Não Conformidade`
  - `Nome da Foto`

---

## 🛠️ Como Usar

1. **Clone o repositório:**

```bash
git clone https://github.com/seu-usuario/Automacao-Relatorios-ARPE.git
cd Automacao-Relatorios-ARPE

Crie e ative um ambiente virtual (recomendado):

bash
Copiar
Editar
python -m venv venv
venv\Scripts\activate no Windows  # ou source/venv/bin/activate no Linux
Instale as dependências:

bash
Copiar
Editar
pip install -r requirements.txt
Coloque suas imagens em assets/ e sua planilha atualizada em data/.

Execute o script principal:

bash
Copiar
Editar
python src/main.py
📦 Dependências
pandas

openpyxl

python-docx

Pillow

Todas estão listadas no arquivo requirements.txt.

✅ Resultados Esperados
Um documento Word (.docx) preenchido automaticamente com:

Dados da planilha substituindo variáveis do modelo

Tabelas com fotos das não conformidades e suas respectivas legendas

Layout e espaçamento adequados para impressão ou compartilhamento digital

💡 Melhorias Futuras
Exportação direta em PDF

Interface gráfica (GUI) para seleção de planilha e pasta

Validação automática de correspondência entre fotos e nomes da planilha

Integração com Google Drive ou Google Sheets

🧑‍💻 Autor
Desenvolvido por Gustavo Lino · https://www.linkedin.com/in/gustavolinoaraujo ·

📄 Licença
Este projeto está sob a licença MIT.

yaml
Copiar
Editar
