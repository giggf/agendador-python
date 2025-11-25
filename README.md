# Agendador de Tarefas em Python 🐍
Um aplicativo desktop para agendar execução de programas e scripts no Windows, com suporte a intervalos recorrentes e ancoragem de horário.

## ✨ Funcionalidades
- Interface Gráfica (GUI) simples e nativa.
- Agendamento por intervalo (ex: A cada 30 minutos, a cada 24 horas).
- **Ancoragem de Tempo:** Se você definir o início para 10:00, ele manterá a execução sempre às 10:00, independente de quando o programa foi aberto.
- Histórico de execução visível na tabela.
- Logs de erro e sucesso em arquivo de texto

## 🚀 Como rodar o projeto (Via Clone)
### Pré-requisitos
Apenas o [Python](https://www.python.org/) instalado no Windows.

### Passo a Passo
1. Clone este repositório:
  ```bash
  git clone + link desse respositório
  ```
2. Entre na pasta:
   ```bash
   cd "NOME DA PASTA DO REPOSITÓRIO"
   ```
3. Execute o script:
   - Apertando 'F5' com o arquivo `agendador.pyw` aberto
   - Clique duplo no arquivo `agendador.pyw`
   - Via terminal:
     ```bash
     python agendador.pyw
     ```
Um arquivo chamado `tarefas.json` será criado, nele todas as tarefas criadas ficaram salvas e apareceram no dashboard assim que o o código for executado.
