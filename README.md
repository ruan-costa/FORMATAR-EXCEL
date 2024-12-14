# Automação de Formatação de Arquivos Excel - Módulo FORMATAR_ARQUIVO.py

O módulo `FORMATAR_ARQUIVO.py` realiza a formatação automática e pré-definida de arquivos Excel. Ele permite que você abra um arquivo Excel, aplique diversas formatações e, finalmente, salve o arquivo com o layout desejado. As funcionalidades incluem a formatação dos cabeçalhos das colunas, ajuste de bordas, ajuste de largura das colunas e ocultação das linhas de grade.

## 🚀 Funcionalidades

- **Pintar Cabeçalhos das Colunas**: Aplica a cor verde ao fundo dos cabeçalhos e define a fonte como branca e em negrito.
- **Bordas nas Células**: Aplica bordas de largura média e na cor preta em todas as células.
- **Ajuste de Largura das Colunas**: Ajusta a largura de todas as colunas automaticamente para se adequar ao conteúdo.
- **Ocultar Linhas de Grade**: Oculta as linhas de grade, mostrando apenas o conteúdo dentro das células delimitadas.

## 📂 Estrutura do Projeto

Abaixo está a estrutura do projeto com a descrição dos arquivos principais:

```bash
.
├── Log/                          # Pasta onde os arquivos de log serão armazenados
├── FORMATAR_ARQUIVO.py           # Script para formatação automática de arquivos Excel
├── README.md                     # Este arquivo README
└── arquivo_de_log.txt            # Arquivo onde as informações detalhadas de log serão salvas

📝 Detalhes do Arquivo FORMATAR_ARQUIVO.py
Função formatar_arquivo(caminho_arquivo):
Esta função é responsável por abrir o arquivo Excel especificado, aplicar todas as formatações automáticas e salvar o arquivo de volta.

Passos de Formatação:
Cabeçalhos das Colunas:
A cor de fundo é definida como verde.
A fonte é alterada para branca e em negrito.
Bordas:
Bordas de largura média e cor preta são aplicadas em todas as células do arquivo Excel.
Ajuste de Largura das Colunas:
A largura das colunas é ajustada automaticamente de acordo com o conteúdo das células.
Ocultar Linhas de Grade:
As linhas de grade são ocultadas, deixando visíveis apenas as células com conteúdo.
⚠️ Observações Importantes
Formato de Arquivo: Este módulo é projetado para funcionar com arquivos Excel suportados pela biblioteca openpyxl.
Log de Execução: O processo de formatação será registrado em um arquivo de log, armazenado na pasta Log/.
📧 Contato
Se tiver dúvidas ou precisar de mais informações, entre em contato:

Email: ruan.xxx1@gmail.com
Tel: 81 9918-44940 / 98761-4812

